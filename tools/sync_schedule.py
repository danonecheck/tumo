#!/usr/bin/env python3
"""Sync student-data in index.html from the coach's schedule-changes xlsx.

Usage:
  python3 tools/sync_schedule.py path/to/file.xlsx            # dry-run (default)
  python3 tools/sync_schedule.py path/to/file.xlsx --apply     # write changes
"""
import argparse
import csv
import datetime
import difflib
import json
import re
import subprocess
import sys
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

REPO_ROOT = Path(__file__).resolve().parent.parent
INDEX_HTML = REPO_ROOT / "index.html"
REPORTS_DIR = REPO_ROOT / "reports"

DAYS_OK = ["Понедельник", "Вторник", "Четверг", "Пятница"]
TIME_MAP = {
    datetime.time(10, 30): "10:30-12:30",
    datetime.time(14, 30): "14:30-16:30",
    datetime.time(16, 30): "16:30-18:30",
}
TIMES_OK = set(TIME_MAP.values())

FIELD_LABELS = {1: "День SL", 2: "Время", 3: "Новый коуч SL", 4: "ex-Coach"}

KZ_MAP = {"ә": "а", "ғ": "г", "қ": "к", "ң": "н", "ө": "о", "ұ": "у", "ү": "у", "һ": "х", "і": "и"}

# Cyrillic spellings of coach names seen in the xlsx -> canonical Latin form.
CYR_TO_LATIN_COACH = {
    "нурай": "Nurai",
    "гульназ": "Gulnaz",
    "эльмира": "Elmira",
    "алинур": "Alinur",
    "алмас": "Almas",
    "аружан": "Aruzhan",
    "акбота": "Akbota",
    "дана ахметова": "Dana Akhmetova",
    "дана телжан": "Dana Telzhan",
    "дана курак": "Dana Kurak",
    "гулим": "Gulim",
    "назым": "Nazym",
    "молдир": "Moldir",
    "жанет": "Zhanet",
    "дильназ": "Dilnaz",
    "дарига": "Dariga",
    "ахмет": "Akhmet",
}

# Canonical casing for the 12 COACH_EMAILS keys (both name aliases collapse to one spelling).
CANON_PRIMARY_COACH = {
    "gulnaz": "Gulnaz", "dana akhmetova": "Dana Akhmetova", "dana ahmetova": "Dana Akhmetova",
    "dana telzhan": "Dana Telzhan", "dana kurak": "Dana Kurak", "nazym": "Nazym",
    "alinur": "Alinur", "aruzhan": "Aruzhan", "nurai": "Nurai", "almas": "Almas",
    "elmira": "Elmira", "dariga": "Dariga",
}


def norm(s):
    s = (s or "").lower()
    s = s.replace("ё", "е").replace("й", "и")
    s = "".join(KZ_MAP.get(ch, ch) for ch in s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def clean_ws(s):
    if s is None:
        return ""
    s = str(s).replace("\t", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def load_coach_emails(html):
    m = re.search(r"const COACH_EMAILS = \{(.*?)\n  \};", html, re.S)
    if not m:
        raise RuntimeError("COACH_EMAILS block not found in index.html")
    pairs = re.findall(r"'([^']+)':\s*'([^']+)'", m.group(1))
    return {k: v for k, v in pairs}


def canon_new_coach(raw, coach_emails):
    """Clean + translit + canonical-case a 'Новый коуч SL' value.
    Returns (value_or_None, ok). ok=False means not resolvable to COACH_EMAILS."""
    s = clean_ws(raw)
    if not s:
        return None, False
    key = s.lower()
    if key in CYR_TO_LATIN_COACH:
        key = CYR_TO_LATIN_COACH[key].lower()
    if key in CANON_PRIMARY_COACH and key in coach_emails:
        return CANON_PRIMARY_COACH[key], True
    return s, False


def canon_ex_coach(raw):
    """Clean whitespace + translit cyrillic for ex-Coach (unrestricted field)."""
    s = clean_ws(raw)
    if not s:
        return ""
    key = s.lower()
    if key in CYR_TO_LATIN_COACH:
        return CYR_TO_LATIN_COACH[key]
    return s


def read_excel_rows(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Sheet1"]
    rows = []
    for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        time_, name, day, ex, new = r[0], r[1], r[2], r[3], r[4]
        name_c = clean_ws(name)
        if not name_c:
            continue  # empty schedule slot, not a student record
        rows.append({
            "excel_row": i,
            "time_raw": time_,
            "name": name_c,
            "day": clean_ws(day),
            "ex_raw": ex,
            "new_raw": new,
        })
    return rows


def dedup_excel_rows(rows, coach_emails):
    """Group by cleaned name; identical rows collapse, differing rows -> problems."""
    groups = defaultdict(list)
    for r in rows:
        groups[r["name"]].append(r)

    deduped = []
    dup_problems = []
    for name, group in groups.items():
        if len(group) == 1:
            deduped.append(group[0])
            continue
        signature = set()
        for r in group:
            new_val, _ = canon_new_coach(r["new_raw"], coach_emails)
            ex_val = canon_ex_coach(r["ex_raw"])
            signature.add((r["day"], fmt_time(r["time_raw"]), new_val, ex_val))
        if len(signature) == 1:
            deduped.append(group[0])
        else:
            dup_problems.append({
                "excel_row": ";".join(str(r["excel_row"]) for r in group),
                "reason": "дубль ФИО в файле с разными значениями",
                "candidates": " | ".join(
                    f"row{r['excel_row']}: {r['day']}/{fmt_time(r['time_raw'])}/"
                    f"new={r['new_raw']}/ex={r['ex_raw']}" for r in group
                ),
                "name": name,
            })
    return deduped, dup_problems


def fmt_time(t):
    if isinstance(t, datetime.time):
        return TIME_MAP.get(t)
    return None


def load_student_data(html):
    m = re.search(r'(<script id="student-data" type="application/json">)(.*?)(</script>)', html, re.S)
    if not m:
        raise RuntimeError("student-data script tag not found in index.html")
    raw_json = m.group(2)
    data = json.loads(raw_json)
    return data, m.span(2)


def build_indexes(data):
    norm_to_indices = defaultdict(list)
    tokens_to_indices = defaultdict(list)
    data_tokens = []
    for idx, rec in enumerate(data):
        n = norm(rec[0])
        norm_to_indices[n].append(idx)
        tokens = frozenset(n.split(" ")) if n else frozenset()
        tokens_to_indices[tokens].append(idx)
        data_tokens.append((idx, tokens))
    return norm_to_indices, tokens_to_indices, data_tokens


def find_internal_dupes(data, norm_to_indices):
    problems = []
    for n, idxs in norm_to_indices.items():
        if len(idxs) > 1:
            problems.append({
                "excel_row": "",
                "reason": "дубль ФИО в student-data (index.html)",
                "candidates": " | ".join(f"[{i}] {data[i][0]}" for i in idxs),
                "name": data[idxs[0]][0],
            })
    return problems


def match_student(excel_name, norm_to_indices, tokens_to_indices, data_tokens):
    n = norm(excel_name)
    idxs = norm_to_indices.get(n, [])
    if len(idxs) == 1:
        return idxs[0], "exact", []
    if len(idxs) > 1:
        return None, "ambiguous_dup_in_data", idxs

    tokens = frozenset(n.split(" ")) if n else frozenset()
    cand = tokens_to_indices.get(tokens, [])
    if len(cand) == 1:
        return cand[0], "wordset", []
    if len(cand) > 1:
        return None, "ambiguous_multi", cand

    # Fallback: one side missing a patronymic (or a word swapped in). Require the
    # shorter side to have >=2 tokens so a single stray word can't over-match.
    if len(tokens) >= 2:
        sub_cand = [i for i, dtok in data_tokens
                    if dtok and dtok != tokens and (tokens <= dtok or dtok <= tokens)]
        if len(sub_cand) == 1:
            return sub_cand[0], "wordset_subset", []
        if len(sub_cand) > 1:
            return None, "ambiguous_multi", sub_cand
    return None, "not_found", []


def process(xlsx_path, verbose_examples=10):
    html = INDEX_HTML.read_text(encoding="utf-8")
    coach_emails = load_coach_emails(html)
    data, _ = load_student_data(html)
    norm_to_indices, tokens_to_indices, data_tokens = build_indexes(data)

    dup_name_problems = find_internal_dupes(data, norm_to_indices)

    excel_rows = read_excel_rows(xlsx_path)
    deduped_rows, dup_row_problems = dedup_excel_rows(excel_rows, coach_emails)

    problems = list(dup_name_problems) + list(dup_row_problems)
    diffs = []
    new_students = []
    matched_indices = set()
    unchanged = 0
    field_change_counts = Counter()

    resolved = []  # rows that matched a single student-data idx; may still collide across rows
    idx_to_rows = defaultdict(list)

    for r in deduped_rows:
        day = r["day"]
        time_val = fmt_time(r["time_raw"])
        new_coach, new_ok = canon_new_coach(r["new_raw"], coach_emails)
        ex_coach = canon_ex_coach(r["ex_raw"])

        row_problems = []
        if day not in DAYS_OK:
            row_problems.append(f"День SL вне списка: {day!r}")
        if time_val is None:
            row_problems.append(f"Время не распознано: {r['time_raw']!r}")
        if not new_ok:
            row_problems.append(f"Новый коуч не в COACH_EMAILS: {r['new_raw']!r}")

        if row_problems:
            problems.append({
                "excel_row": r["excel_row"],
                "reason": "; ".join(row_problems),
                "candidates": "",
                "name": r["name"],
            })
            continue

        idx, how, cand_idxs = match_student(r["name"], norm_to_indices, tokens_to_indices, data_tokens)

        if how == "ambiguous_dup_in_data":
            problems.append({
                "excel_row": r["excel_row"],
                "reason": "имя совпадает с дублем в student-data",
                "candidates": " | ".join(data[i][0] for i in cand_idxs),
                "name": r["name"],
            })
            continue
        if how == "ambiguous_multi":
            problems.append({
                "excel_row": r["excel_row"],
                "reason": "несколько кандидатов по множеству слов",
                "candidates": " | ".join(data[i][0] for i in cand_idxs),
                "name": r["name"],
            })
            continue
        if how == "not_found":
            new_students.append({
                "excel_row": r["excel_row"],
                "name": r["name"],
                "day": day,
                "time": time_val,
                "new_coach": new_coach,
                "ex_coach": ex_coach,
            })
            continue

        entry = {"r": r, "idx": idx, "day": day, "time_val": time_val,
                 "new_coach": new_coach, "ex_coach": ex_coach}
        resolved.append(entry)
        idx_to_rows[idx].append(entry)

    for idx, entries in idx_to_rows.items():
        if len(entries) > 1:
            problems.append({
                "excel_row": ";".join(str(e["r"]["excel_row"]) for e in entries),
                "reason": "разные строки файла указывают на одного ученика в student-data",
                "candidates": " | ".join(
                    f"row{e['r']['excel_row']} ({e['r']['name']}): {e['day']}/{e['time_val']}/"
                    f"new={e['new_coach']}/ex={e['ex_coach']}" for e in entries
                ),
                "name": data[idx][0],
            })
            continue

        entry = entries[0]
        matched_indices.add(idx)
        rec = data[idx]
        new_values = {1: entry["day"], 2: entry["time_val"], 3: entry["new_coach"], 4: entry["ex_coach"]}
        row_diffs = []
        for field, new_val in new_values.items():
            old_val = rec[field]
            if old_val != new_val:
                row_diffs.append((field, old_val, new_val))
        if row_diffs:
            for field, old_val, new_val in row_diffs:
                diffs.append({
                    "name": rec[0],
                    "field": FIELD_LABELS[field],
                    "old": old_val,
                    "new": new_val,
                    "excel_row": entry["r"]["excel_row"],
                })
                field_change_counts[FIELD_LABELS[field]] += 1
        else:
            unchanged += 1

    not_in_file = [rec[0] for i, rec in enumerate(data) if i not in matched_indices]

    not_in_file_norm = {name: norm(name) for name in not_in_file}
    for n in new_students:
        query = norm(n["name"])
        close = difflib.get_close_matches(query, list(not_in_file_norm.values()), n=3, cutoff=0.6)
        candidates = [name for name, nv in not_in_file_norm.items() if nv in close]
        n["fuzzy_candidates"] = " | ".join(candidates)

    return {
        "coach_emails": coach_emails,
        "data": data,
        "excel_total": len(excel_rows),
        "excel_deduped": len(deduped_rows),
        "matched": len(matched_indices),
        "unchanged": unchanged,
        "diffs": diffs,
        "field_change_counts": field_change_counts,
        "problems": problems,
        "new_students": new_students,
        "not_in_file": not_in_file,
        "matched_indices": matched_indices,
    }


def write_reports(result):
    REPORTS_DIR.mkdir(exist_ok=True)

    with open(REPORTS_DIR / "diff.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["ФИО", "поле", "было", "стало", "excel_row"])
        for d in result["diffs"]:
            w.writerow([d["name"], d["field"], d["old"], d["new"], d["excel_row"]])

    with open(REPORTS_DIR / "problems.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["excel_row", "ФИО", "причина", "кандидаты"])
        for p in result["problems"]:
            w.writerow([p["excel_row"], p["name"], p["reason"], p["candidates"]])

    with open(REPORTS_DIR / "not_in_file.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["ФИО"])
        for name in result["not_in_file"]:
            w.writerow([name])

    with open(REPORTS_DIR / "new_students.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["excel_row", "ФИО", "День SL", "Время", "Новый коуч SL", "ex-Coach", "похожие_в_student-data"])
        for n in result["new_students"]:
            w.writerow([n["excel_row"], n["name"], n["day"], n["time"], n["new_coach"], n["ex_coach"], n["fuzzy_candidates"]])


def print_summary(result, examples=10):
    print(f"Всего строк в файле (после очистки пустых): {result['excel_total']}")
    print(f"После схлопывания дублей ФИО в файле: {result['excel_deduped']}")
    print(f"Сопоставлено с student-data: {result['matched']}")
    print(f"Без изменений: {result['unchanged']}")
    print(f"Изменений по полям:")
    for field, count in result["field_change_counts"].items():
        print(f"  {field}: {count}")
    print(f"Всего изменённых записей: {len(set((d['name'] for d in result['diffs'])))}")
    print(f"Проблем: {len(result['problems'])}")
    print(f"Новых учеников (не найдены на сайте): {len(result['new_students'])}")
    print(f"Учеников из student-data, которых нет в файле: {len(result['not_in_file'])}")

    print(f"\n--- Примеры diff (до {examples}) ---")
    for d in result["diffs"][:examples]:
        print(f"  {d['name']} | {d['field']}: {d['old']!r} -> {d['new']!r}")

    print(f"\n--- Примеры problems (до {examples}) ---")
    for p in result["problems"][:examples]:
        print(f"  row {p['excel_row']} | {p['name']} | {p['reason']} | {p['candidates']}")

    print(f"\n--- Примеры new_students (до {examples}) ---")
    for n in result["new_students"][:examples]:
        suffix = f" | похоже на: {n['fuzzy_candidates']}" if n["fuzzy_candidates"] else ""
        print(f"  row {n['excel_row']} | {n['name']} | {n['day']} {n['time']} new={n['new_coach']} ex={n['ex_coach']}{suffix}")


def apply_changes(result):
    html = INDEX_HTML.read_text(encoding="utf-8")
    data = result["data"]

    field_name_to_idx = {v: k for k, v in FIELD_LABELS.items()}
    by_name = defaultdict(list)
    for d in result["diffs"]:
        by_name[d["name"]].append(d)

    idx_by_name = {rec[0]: i for i, rec in enumerate(data)}
    for name, changes in by_name.items():
        idx = idx_by_name[name]
        for c in changes:
            field_idx = field_name_to_idx[c["field"]]
            data[idx][field_idx] = c["new"]

    for rec in data:
        for v in rec:
            if isinstance(v, str) and "</script" in v.lower():
                raise RuntimeError(f"Unsafe value contains </script: {v!r}")

    new_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    m = re.search(r'(<script id="student-data" type="application/json">)(.*?)(</script>)', html, re.S)
    new_html = html[:m.start(2)] + new_json + html[m.end(2):]
    INDEX_HTML.write_text(new_html, encoding="utf-8")

    # Post-write verification
    verify_data, _ = load_student_data(INDEX_HTML.read_text(encoding="utf-8"))
    assert len(verify_data) == len(data), "record count changed unexpectedly"
    for rec in verify_data:
        assert rec[1] in DAYS_OK, f"bad day after apply: {rec}"
        assert rec[2] in TIMES_OK, f"bad time after apply: {rec}"
        assert rec[3].lower() in result["coach_emails"], f"bad coach after apply: {rec}"
    print(f"Applied {len(result['diffs'])} field changes across {len(by_name)} students.")
    print("index.html updated and verified.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", help="Path to the Изменения_графика_SL.xlsx file")
    ap.add_argument("--apply", action="store_true", help="Write changes to index.html (default: dry-run)")
    args = ap.parse_args()

    result = process(args.xlsx)
    write_reports(result)
    print_summary(result)

    if args.apply:
        print("\n--- APPLYING CHANGES ---")
        apply_changes(result)
    else:
        print("\nDry-run only. Re-run with --apply to write changes.")


if __name__ == "__main__":
    main()
