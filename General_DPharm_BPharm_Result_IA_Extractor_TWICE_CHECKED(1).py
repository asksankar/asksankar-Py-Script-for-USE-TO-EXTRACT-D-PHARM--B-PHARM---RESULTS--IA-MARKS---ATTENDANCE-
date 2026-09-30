"""
GENERAL D.Pharm / B.Pharm RESULT + IA ANALYZER
Python 3.10+ / Python 3.14 compatible

INPUT 1: University Result PDF file(s)
INPUT 2: IA Marks / IA Attendance PDF, XLSX, XLS, CSV file(s)

OUTPUT: One consolidated Excel workbook containing:
- Semester/Year Summary
- Subject Summary
- FAIL
- ABSENT
- FAIL with IA
- ABSENT with IA
- Detailed Result
- IA Data
- IA Match Log
- Candidate Summary
- Extraction Log
- Read Me

Important:
The IA table may use either:
A) App | IA | Total IA Att.(%) | Credit Points
or
B) App | SUBJECT - ATTENDANCE | SUBJECT - INT | ...
The program detects the layout and does NOT treat App as IA marks.
"""

import os
import re
import math
import traceback
import threading
from pathlib import Path
from collections import defaultdict

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    import pymupdf
except ImportError:
    pymupdf = None

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None


STATUS_WORDS = {"PASS", "FAIL", "ABSENT", "---"}
YEAR_WORDS = [
    "FIRST YEAR", "SECOND YEAR", "THIRD YEAR", "FOURTH YEAR",
    "FINAL YEAR", "FIRST SEMESTER", "SECOND SEMESTER",
    "THIRD SEMESTER", "FOURTH SEMESTER", "FIFTH SEMESTER",
    "SIXTH SEMESTER", "SEVENTH SEMESTER", "EIGHTH SEMESTER",
]

def norm(s):
    s = "" if s is None else str(s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def norm_reg(s):
    s = norm(s)
    s = re.sub(r"\.0$", "", s)
    # retain leading zeroes when supplied as text
    m = re.search(r"\d{5,15}", s)
    return m.group(0) if m else s

def norm_code(s):
    s = norm(s)
    m = re.search(r"\d{5,8}", s)
    return m.group(0) if m else s

def safe_float(s):
    try:
        v = float(str(s).replace(",", "").strip())
        if math.isnan(v):
            return None
        return v
    except Exception:
        return None

def unique_name(base, used):
    b = re.sub(r'[\\/:*?"<>|]+', "_", base)[:25] or "Sheet"
    name = b
    n = 2
    while name in used:
        name = f"{b[:28]}_{n}"
        n += 1
    used.add(name)
    return name

def extract_year_term(text):
    t = norm(text).upper()
    # Prefer explicit "Term:" if present
    m = re.search(r"TERM\s*:\s*([^\\n]+)", t)
    if m:
        term = norm(m.group(1))
    else:
        term = ""
    year = ""
    for y in YEAR_WORDS:
        if y in t:
            year = y
            break
    if not year:
        # common alternatives
        m2 = re.search(r"\b(YEAR|SEMESTER)\s*[-:]?\s*([1-8])\b", t)
        if m2:
            n = m2.group(2)
            year = (["FIRST","SECOND","THIRD","FOURTH","FIFTH","SIXTH","SEVENTH","EIGHTH"][int(n)-1]
                    + " " + m2.group(1))
    return year or term or "", term or year or ""

def detect_subjects_from_header(lines):
    """Return [(subject_code, subject_name)] from the page header."""
    # Header ends at first student S.No row.
    row_start = None
    for i, ln in enumerate(lines):
        if re.fullmatch(r"\s*\d+\s*", ln):
            # Need a plausible register number nearby
            if i + 1 < len(lines) and re.fullmatch(r"\s*\d{5,15}\s*", lines[i+1]):
                row_start = i
                break
    header = lines[:row_start] if row_start is not None else lines[:80]

    found = []
    code_positions = []
    for i, ln in enumerate(header):
        if re.fullmatch(r"\s*\d{5,8}\s*", ln):
            # Avoid date/time / pattern fragments: subject codes are usually 5-8 digits.
            code_positions.append(i)

    for p, idx in enumerate(code_positions):
        code = re.sub(r"\D", "", header[idx])
        nxt = code_positions[p+1] if p+1 < len(code_positions) else len(header)
        name_parts = []
        for ln in header[idx+1:nxt]:
            u = norm(ln)
            if not u:
                continue
            if "PATTERN" in u.upper():
                continue
            # Don't include general header labels.
            if u.upper() in {"A", "P", "P", "REGISTER NO.", "NAME OF THE CANDIDATE"}:
                continue
            if "RESULT GALLEY" in u.upper() or "DIRECTORATE" in u.upper():
                break
            name_parts.append(u)
        # remove obvious boilerplate / numbers
        cleaned = []
        for x in name_parts:
            if re.fullmatch(r"[\d\s.%-]+", x):
                continue
            cleaned.append(x)
        name = norm(" ".join(cleaned))
        if name:
            found.append((code, name))
    return found

def split_student_rows(lines):
    """Split page text into student row chunks using standalone S.No lines."""
    starts = []
    for i, ln in enumerate(lines):
        if re.fullmatch(r"\s*\d+\s*", ln):
            if i + 1 < len(lines) and re.fullmatch(r"\s*\d{5,15}\s*", lines[i+1]):
                starts.append(i)
    chunks = []
    for j, st in enumerate(starts):
        en = starts[j+1] if j+1 < len(starts) else len(lines)
        chunks.append(lines[st:en])
    return chunks

def parse_result_page(page_text):
    lines = [x.strip() for x in page_text.splitlines() if x.strip()]
    subjects = detect_subjects_from_header(lines)
    rows = []
    for chunk in split_student_rows(lines):
        tokens = []
        for ln in chunk:
            tokens.extend(norm(ln).split())
        # tokens: SNo RegNo Name... App PASS...
        if len(tokens) < 5:
            continue
        try:
            sno_i = 0
            reg_i = next(i for i in range(1, min(5, len(tokens))) if re.fullmatch(r"\d{5,15}", tokens[i]))
        except StopIteration:
            continue
        reg = tokens[reg_i]
        status_start = None
        for i in range(reg_i+1, len(tokens)):
            if tokens[i].upper() in STATUS_WORDS:
                status_start = i
                break
        if status_start is None:
            continue
        # App is the last numeric token before first status
        app_i = None
        for i in range(reg_i+1, status_start):
            if re.fullmatch(r"\d+(?:\.\d+)?", tokens[i]):
                app_i = i
        if app_i is None:
            continue
        name = norm(" ".join(tokens[reg_i+1:app_i]))
        statuses = [x.upper() for x in tokens[status_start:]]
        if not subjects:
            continue
        statuses = statuses[:len(subjects)]
        if len(statuses) < len(subjects):
            statuses += ["---"] * (len(subjects) - len(statuses))
        for (code, subname), st in zip(subjects, statuses):
            rows.append({
                "Register No": norm_reg(reg),
                "Student Name": name,
                "Subject Code": code,
                "Subject Name": subname,
                "Status": st,
            })
    return subjects, rows

def parse_result_pdf(files):
    if pymupdf is None:
        raise RuntimeError("PyMuPDF is not installed. Run: py -m pip install --upgrade pymupdf")
    all_records = []
    logs = []
    prev_subjects = []
    prev_year = ""
    for f in files:
        doc = pymupdf.open(f)
        for pno, page in enumerate(doc, 1):
            text = page.get_text("text")
            upper = text.upper()
            year, term = extract_year_term(text)
            if year:
                prev_year = year
            subjects, rows = parse_result_page(text)
            if subjects:
                prev_subjects = subjects
            elif prev_subjects:
                # Some continuation pages may omit the repeated subject heading.
                # Reparse with carried subject list.
                lines = [x.strip() for x in text.splitlines() if x.strip()]
                rows = []
                for chunk in split_student_rows(lines):
                    tokens = []
                    for ln in chunk:
                        tokens.extend(norm(ln).split())
                    if len(tokens) < 5:
                        continue
                    try:
                        reg_i = next(i for i in range(1, min(5, len(tokens))) if re.fullmatch(r"\d{5,15}", tokens[i]))
                    except StopIteration:
                        continue
                    status_start = next((i for i in range(reg_i+1, len(tokens))
                                         if tokens[i].upper() in STATUS_WORDS), None)
                    if status_start is None:
                        continue
                    app_i = None
                    for i in range(reg_i+1, status_start):
                        if re.fullmatch(r"\d+(?:\.\d+)?", tokens[i]):
                            app_i = i
                    if app_i is None:
                        continue
                    name = norm(" ".join(tokens[reg_i+1:app_i]))
                    statuses = [x.upper() for x in tokens[status_start:status_start+len(prev_subjects)]]
                    if len(statuses) < len(prev_subjects):
                        statuses += ["---"]*(len(prev_subjects)-len(statuses))
                    for (code, subname), st in zip(prev_subjects, statuses):
                        rows.append({
                            "Register No": norm_reg(tokens[reg_i]),
                            "Student Name": name,
                            "Subject Code": code,
                            "Subject Name": subname,
                            "Status": st,
                        })
                subjects = prev_subjects
            if rows:
                for r in rows:
                    r["Year/Semester"] = year or prev_year or ""
                    r["Result File"] = os.path.basename(f)
                    r["Result Page"] = pno
                all_records.extend(rows)
                logs.append(f"RESULT OK: {os.path.basename(f)} page {pno}: {len(rows)} subject records")
            else:
                logs.append(f"RESULT WARNING: {os.path.basename(f)} page {pno}: no student rows detected")
        doc.close()
    return all_records, logs

def parse_ia_page(page_text):
    lines = [x.strip() for x in page_text.splitlines() if x.strip()]
    # Subject and term are normally in footer. Search entire page.
    subject_code, subject_name = "", ""
    m = re.search(r"Subject\s*:\s*(\d{5,8})\s*-\s*(.+?)(?:\s+Exam Type\s*:|$)", page_text,
                  flags=re.I | re.S)
    if m:
        subject_code = norm_code(m.group(1))
        subject_name = norm(m.group(2).replace("\n", " "))
    year, term = extract_year_term(page_text)

    # Determine whether header indicates attendance before INT.
    header_u = " ".join(lines[:35]).upper()
    dpharm_layout = ("ATTENDANCE" in header_u and "INT" in header_u
                     and header_u.find("ATTENDANCE") < header_u.find("INT"))
    # B.Pharm-like layout: App, IA, Total/IA Att.(%)
    bpharm_layout = ("IA ATT" in header_u or "TOTAL IA ATT" in header_u)

    records = []
    for chunk in split_student_rows(lines):
        tokens = []
        for ln in chunk:
            tokens.extend(norm(ln).replace("|", " ").split())
        if len(tokens) < 6:
            continue
        try:
            reg_i = next(i for i in range(1, min(5, len(tokens))) if re.fullmatch(r"\d{5,15}", tokens[i]))
        except StopIteration:
            continue
        # Collect numeric values after name; App is first numeric after register/name
        nums = []
        for i in range(reg_i+1, len(tokens)):
            if re.fullmatch(r"-?\d+(?:\.\d+)?", tokens[i]):
                nums.append((i, safe_float(tokens[i])))
        if not nums:
            continue
        app_i, app_val = nums[0]
        # Name is everything between reg and App
        name = norm(" ".join(tokens[reg_i+1:app_i]))
        values = [v for _, v in nums[1:]]
        if not values:
            continue

        ia_marks = None
        ia_att = None

        # Exact layouts supported:
        # D.Pharm example: App, Attendance %, INT marks
        if dpharm_layout:
            if len(values) >= 2:
                ia_att = values[0]
                ia_marks = values[1]
        # B.Pharm example: App, IA marks, Total IA Att(%), Credit
        elif bpharm_layout:
            if len(values) >= 2:
                ia_marks = values[0]
                ia_att = values[1]
        else:
            # Generic fallback: inspect numeric ranges.
            if len(values) >= 2:
                a, b = values[0], values[1]
                # attendance is normally 0-100, IA marks normally <= 100 but
                # the attendance field is much more likely to be a percentage.
                if b is not None and a is not None and b <= 100 and a <= 100:
                    # Prefer a percentage-looking second number when one has decimals
                    # and the other is a small IA score.
                    if a <= 30 < b:
                        ia_marks, ia_att = a, b
                    else:
                        ia_marks, ia_att = a, b

        records.append({
            "Register No": norm_reg(tokens[reg_i]),
            "Student Name": name,
            "Subject Code": subject_code,
            "Subject Name": subject_name,
            "IA Marks": ia_marks,
            "IA Attendance (%)": ia_att,
            "IA File": "",
            "IA Page": "",
            "IA Year/Semester": year,
            "IA Eligibility": "",
        })
    return records

def parse_ia_pdf(files):
    if pymupdf is None:
        raise RuntimeError("PyMuPDF is not installed. Run: py -m pip install --upgrade pymupdf")
    out, logs = [], []
    for f in files:
        doc = pymupdf.open(f)
        for pno, page in enumerate(doc, 1):
            text = page.get_text("text")
            recs = parse_ia_page(text)
            for r in recs:
                r["IA File"] = os.path.basename(f)
                r["IA Page"] = pno
                # Eligibility from page row is handled below from raw text if needed.
                out.append(r)
            if recs:
                logs.append(f"IA PDF OK: {os.path.basename(f)} page {pno}: {len(recs)} records")
        doc.close()
    return out, logs

def find_col(df, aliases):
    cols = list(df.columns)
    nmap = {re.sub(r"[^A-Z0-9]", "", str(c).upper()): c for c in cols}
    for a in aliases:
        k = re.sub(r"[^A-Z0-9]", "", a.upper())
        if k in nmap:
            return nmap[k]
    # partial
    for c in cols:
        k = re.sub(r"[^A-Z0-9]", "", str(c).upper())
        for a in aliases:
            ak = re.sub(r"[^A-Z0-9]", "", a.upper())
            if ak and ak in k:
                return c
    return None

def parse_ia_excel(files):
    if pd is None:
        raise RuntimeError("pandas is not installed. Run: py -m pip install --upgrade pandas openpyxl xlrd")
    out, logs = [], []
    for f in files:
        ext = Path(f).suffix.lower()
        if ext == ".csv":
            sheets = {"CSV": pd.read_csv(f)}
        else:
            xls = pd.ExcelFile(f)
            sheets = {s: pd.read_excel(f, sheet_name=s) for s in xls.sheet_names}

        for sname, df in sheets.items():
            if df.empty:
                continue
            reg_c = find_col(df, ["Register No","Register Number","Reg No","Reg.No","Register"])
            name_c = find_col(df, ["Student Name","Candidate Name","Name"])
            code_c = find_col(df, ["Subject Code","Sub Code","Subject"])
            sub_c = find_col(df, ["Subject Name","Subject Description"])
            ia_c = find_col(df, ["IA Marks","IA Mark","Internal Marks","IA","INT","Internal"])
            att_c = find_col(df, ["IA Attendance","IA Att","IA Attendance (%)","Attendance","Att.(%)","Attendance (%)"])
            elig_c = find_col(df, ["Eligibility","Eligible","Status"])

            if reg_c is None:
                logs.append(f"IA EXCEL WARNING: {os.path.basename(f)} / {sname}: Register No column not found")
                continue

            # If subject code is absent but a single subject is implied by sheet/file name,
            # try to discover a code in that sheet name.
            sheet_code = norm_code(sname) or norm_code(Path(f).stem)
            for _, row in df.iterrows():
                reg = norm_reg(row.get(reg_c, ""))
                if not reg:
                    continue
                name = norm(row.get(name_c, "")) if name_c else ""
                code = norm_code(row.get(code_c, "")) if code_c else sheet_code
                sname_val = norm(row.get(sub_c, "")) if sub_c else ""
                ia = safe_float(row.get(ia_c)) if ia_c else None
                att = safe_float(row.get(att_c)) if att_c else None
                eligibility = norm(row.get(elig_c, "")) if elig_c else ""
                if code or sname_val or ia is not None or att is not None:
                    out.append({
                        "Register No": reg,
                        "Student Name": name,
                        "Subject Code": code,
                        "Subject Name": sname_val,
                        "IA Marks": ia,
                        "IA Attendance (%)": att,
                        "IA File": os.path.basename(f),
                        "IA Page": "",
                        "IA Year/Semester": "",
                        "IA Eligibility": eligibility,
                    })
            logs.append(f"IA EXCEL OK: {os.path.basename(f)} / {sname}: {len(df)} rows read")
    return out, logs

def parse_ia_files(files):
    pdf = [f for f in files if Path(f).suffix.lower() == ".pdf"]
    tab = [f for f in files if Path(f).suffix.lower() in {".xlsx",".xls",".csv"}]
    records, logs = [], []
    if pdf:
        r, l = parse_ia_pdf(pdf)
        records.extend(r); logs.extend(l)
    if tab:
        r, l = parse_ia_excel(tab)
        records.extend(r); logs.extend(l)
    return records, logs

def dedupe_result(records):
    d = {}
    for r in records:
        key = (r["Register No"], r["Subject Code"], r["Year/Semester"], r["Status"])
        d[key] = r
    return list(d.values())


def value_key(v):
    """Strict comparison key for double verification."""
    if v is None:
        return ""
    if isinstance(v, float):
        if math.isnan(v):
            return ""
        return f"{v:.6f}".rstrip("0").rstrip(".")
    return norm(v).upper()


def record_verification_key(r):
    return (
        norm_reg(r.get("Register No", "")),
        norm_code(r.get("Subject Code", "")),
        norm(r.get("Subject Name", "")).upper(),
        value_key(r.get("IA Marks")),
        value_key(r.get("IA Attendance (%)")),
    )


def two_pass_verify_ia(ia_records):
    """Verify that every IA record is internally consistent before export.

    Pass 1: normal parsed records are grouped by the exact identity/value key.
    Pass 2: duplicate records for the same Register+Subject are compared.
    Export is permitted only when no key has conflicting IA values.
    """
    grouped = defaultdict(list)
    for r in ia_records:
        base = (norm_reg(r.get("Register No", "")), norm_code(r.get("Subject Code", "")))
        if base[0] and base[1]:
            grouped[base].append(r)

    verified = []
    conflicts = []
    seen = set()
    for base, recs in grouped.items():
        keys = {record_verification_key(r) for r in recs}
        if len(keys) == 1:
            # keep the first record; all repeated copies agree exactly
            verified.append(recs[0])
            seen.add(base)
        else:
            conflicts.append({
                "Register No": base[0],
                "Subject Code": base[1],
                "Verification": "CONFLICT - EXPORT BLOCKED",
                "Records Found": len(recs),
                "Values": " | ".join(sorted(str(k) for k in keys)),
            })

    # Records without a usable subject code cannot be safely matched.
    for r in ia_records:
        if not norm_code(r.get("Subject Code", "")) or not norm_reg(r.get("Register No", "")):
            conflicts.append({
                "Register No": norm_reg(r.get("Register No", "")),
                "Subject Code": norm_code(r.get("Subject Code", "")),
                "Verification": "MISSING KEY - EXPORT BLOCKED",
                "Records Found": 1,
                "Values": str(record_verification_key(r)),
            })

    # Verify that the surviving records themselves contain the same identity/value
    # when normalized a second time. This catches whitespace/case/decimal artifacts.
    second_check = []
    for r in verified:
        k1 = record_verification_key(r)
        clone = dict(r)
        clone["Register No"] = norm_reg(clone.get("Register No", ""))
        clone["Subject Code"] = norm_code(clone.get("Subject Code", ""))
        clone["Subject Name"] = norm(clone.get("Subject Name", ""))
        clone["IA Marks"] = safe_float(clone.get("IA Marks")) if clone.get("IA Marks") not in (None, "") else None
        clone["IA Attendance (%)"] = safe_float(clone.get("IA Attendance (%)")) if clone.get("IA Attendance (%)") not in (None, "") else None
        k2 = record_verification_key(clone)
        second_check.append((k1, k2))
    for k1, k2 in second_check:
        if k1 != k2:
            conflicts.append({
                "Register No": k1[0], "Subject Code": k1[1],
                "Verification": "SECOND PASS MISMATCH - EXPORT BLOCKED",
                "Records Found": 1, "Values": f"PASS1={k1} ; PASS2={k2}"
            })

    return verified, conflicts

def attach_ia(problem_records, ia_records):
    idx = defaultdict(list)
    for r in ia_records:
        key = (norm_reg(r.get("Register No","")), norm_code(r.get("Subject Code","")))
        if key[0] and key[1]:
            idx[key].append(r)

    output = []
    logs = []
    for r in problem_records:
        key = (r["Register No"], r["Subject Code"])
        matches = idx.get(key, [])
        # If multiple matching IA records exist, prefer one with same year.
        same_year = [m for m in matches if norm(m.get("IA Year/Semester","")) == norm(r.get("Year/Semester",""))
                     and m.get("IA Year/Semester","")]
        if same_year:
            matches = same_year
        status = "NOT FOUND"
        match = None
        if len(matches) == 1:
            status = "MATCHED"
            match = matches[0]
        elif len(matches) > 1:
            status = "AMBIGUOUS"
        rec = dict(r)
        rec.update({
            "IA Marks": match.get("IA Marks") if match else None,
            "IA Attendance (%)": match.get("IA Attendance (%)") if match else None,
            "IA Eligibility": match.get("IA Eligibility","") if match else "",
            "IA Match Status": status,
            "IA Match File": match.get("IA File","") if match else "",
            "IA Match Page": match.get("IA Page","") if match else "",
        })
        output.append(rec)
        logs.append({
            "Register No": r["Register No"],
            "Student Name": r["Student Name"],
            "Year/Semester": r["Year/Semester"],
            "Subject Code": r["Subject Code"],
            "Subject Name": r["Subject Name"],
            "Result Status": r["Status"],
            "IA Marks": rec["IA Marks"],
            "IA Attendance (%)": rec["IA Attendance (%)"],
            "IA Match Status": status,
            "IA File": rec["IA Match File"],
            "IA Page": rec["IA Match Page"],
        })
    return output, logs

def make_summaries(result_records):
    # Group by year/semester and then subject.
    sem = defaultdict(lambda: {"students": set(), "passed_all": set(), "with_result": set()})
    subj = defaultdict(lambda: {"PASS":0, "FAIL":0, "ABSENT":0, "APPLICABLE":set(), "students":set()})

    by_student_year = defaultdict(list)
    for r in result_records:
        y = r.get("Year/Semester","") or "UNSPECIFIED"
        by_student_year[(y, r["Register No"])].append(r)
        s = subj[(y, r["Subject Code"], r["Subject Name"])]
        st = r["Status"]
        s["students"].add(r["Register No"])
        if st in {"PASS","FAIL","ABSENT"}:
            s[st] += 1
            s["APPLICABLE"].add(r["Register No"])
        sem[y]["students"].add(r["Register No"])
        if st in {"PASS","FAIL","ABSENT"}:
            sem[y]["with_result"].add(r["Register No"])

    for (y, reg), rows in by_student_year.items():
        applicable = [r["Status"] for r in rows if r["Status"] in {"PASS","FAIL","ABSENT"}]
        if applicable and all(x == "PASS" for x in applicable):
            sem[y]["passed_all"].add(reg)

    sem_rows = []
    for y, d in sorted(sem.items()):
        n = len(d["with_result"])
        p = len(d["passed_all"])
        sem_rows.append({
            "Year/Semester": y,
            "Candidates With Results": n,
            "Passed ALL Applicable Subjects": p,
            "Class Pass %": round(p/n*100,2) if n else 0,
            "Candidates in Result File": len(d["students"]),
        })

    subj_rows = []
    for (y, code, name), d in sorted(subj.items()):
        attended = d["PASS"] + d["FAIL"]
        subj_rows.append({
            "Year/Semester": y,
            "Subject Code": code,
            "Subject Name": name,
            "PASS": d["PASS"],
            "FAIL": d["FAIL"],
            "ABSENT": d["ABSENT"],
            "Attended (PASS+FAIL)": attended,
            "Subject Pass %": round(d["PASS"]/attended*100,2) if attended else 0,
        })
    return sem_rows, subj_rows

def build_candidate_summary(result_records):
    by = defaultdict(list)
    names = {}
    for r in result_records:
        y = r.get("Year/Semester","") or "UNSPECIFIED"
        by[(y,r["Register No"])].append(r)
        names[r["Register No"]] = r["Student Name"]
    rows = []
    for (y,reg), rs in sorted(by.items()):
        relevant = [r for r in rs if r["Status"] in {"PASS","FAIL","ABSENT"}]
        fail = sum(r["Status"]=="FAIL" for r in relevant)
        absent = sum(r["Status"]=="ABSENT" for r in relevant)
        passed = bool(relevant) and fail==0 and absent==0
        rows.append({
            "Year/Semester": y,
            "Register No": reg,
            "Student Name": names.get(reg,""),
            "Subjects With Result": len(relevant),
            "FAIL Subjects": fail,
            "ABSENT Subjects": absent,
            "PASS ALL": "YES" if passed else "NO",
        })
    return rows

def write_sheet(ws, rows, title=None):
    ws.delete_rows(1, ws.max_row)
    if title:
        ws.cell(1,1,title)
        ws.cell(1,1).font = Font(bold=True, size=14)
        start = 3
    else:
        start = 1
    if not rows:
        ws.cell(start,1,"No records")
        return
    keys = list(rows[0].keys())
    for c,k in enumerate(keys,1):
        cell = ws.cell(start,c,k)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for rno,row in enumerate(rows,start+1):
        for c,k in enumerate(keys,1):
            ws.cell(rno,c,row.get(k,""))
    ws.freeze_panes = ws.cell(start+1,1)
    ws.auto_filter.ref = f"A{start}:{get_column_letter(len(keys))}{start+len(rows)}"
    for col in range(1,len(keys)+1):
        max_len = len(str(ws.cell(start,col).value or ""))
        for rr in range(start+1, min(start+200,len(rows)+start+1)):
            max_len = max(max_len, len(str(ws.cell(rr,col).value or "")))
        ws.column_dimensions[get_column_letter(col)].width = min(max(max_len+2,10),42)

def export_excel(out_file, sem_rows, subj_rows, fail_rows, absent_rows,
                  fail_ia, absent_ia, detailed, ia_records, match_logs,
                  candidate_rows, extraction_logs, verification_rows):
    if Workbook is None:
        raise RuntimeError("openpyxl is not installed. Run: py -m pip install --upgrade openpyxl")
    wb = Workbook()
    ws = wb.active
    ws.title = "Semester Summary"
    write_sheet(ws, sem_rows, "CLASS PASS PERCENTAGE - YEAR / SEMESTER")

    ws = wb.create_sheet("Subject Summary")
    write_sheet(ws, subj_rows, "SUBJECT-WISE PASS PERCENTAGE")

    for nm, rows in [
        ("Fail", fail_rows),
        ("Absent", absent_rows),
        ("Fail with IA", fail_ia),
        ("Absent with IA", absent_ia),
        ("Detailed Result", detailed),
        ("IA Data", ia_records),
        ("IA Match Log", match_logs),
        ("Candidate Summary", candidate_rows),
        ("Extraction Log", [{"Message":x} for x in extraction_logs]),
        ("Two-Pass Verification", verification_rows),
    ]:
        ws = wb.create_sheet(nm)
        write_sheet(ws, rows, nm.upper())

    ws = wb.create_sheet("Read Me")
    readme = [
        {"Item":"Result input","Value":"University result PDF file(s)"},
        {"Item":"IA input","Value":"IA PDF / XLSX / XLS / CSV file(s)"},
        {"Item":"IA matching","Value":"Register No + Subject Code; year is used as a disambiguation when available"},
        {"Item":"IA Marks","Value":"For D.Pharm layout App | Attendance | INT, INT is used as IA Marks"},
        {"Item":"IA Attendance","Value":"D.Pharm layout: attendance is the percentage immediately after App; B.Pharm layout: IA Attendance (%) is detected from the header"},
        {"Item":"Subject pass %","Value":"PASS / (PASS + FAIL) * 100; ABSENT excluded"},
        {"Item":"Class pass %","Value":"Students who passed ALL applicable subjects / students having a result"},
        {"Item":"---","Value":"Treated as not applicable / not reported; it is not counted as FAIL"},
    ]
    write_sheet(ws, readme, "README")
    wb.save(out_file)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("General D.Pharm / B.Pharm Result + IA Analyzer")
        self.root.geometry("900x680")

        self.result_files = []
        self.ia_files = []

        frm = ttk.Frame(root, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="GENERAL RESULT + IA ANALYZER",
                  font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0,8))
        ttk.Label(frm, text="Works with First Year / Second Year / Final Year and semester-style result PDFs.").pack(anchor="w")

        self.make_file_group(frm, "1. RESULT FILE(S) - PDF", self.add_results, self.result_files)
        self.make_file_group(frm, "2. IA MARKS / ATTENDANCE FILE(S) - PDF / Excel / CSV",
                             self.add_ia, self.ia_files)

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=10)
        ttk.Button(btns, text="PROCESS + EXPORT EXCEL", command=self.start).pack(side="left", padx=4)
        ttk.Button(btns, text="CLEAR", command=self.clear).pack(side="left", padx=4)

        self.progress = ttk.Progressbar(frm, mode="indeterminate")
        self.progress.pack(fill="x", pady=5)

        self.status = ttk.Label(frm, text="Ready")
        self.status.pack(anchor="w", pady=4)

        ttk.Label(frm, text="Processing log").pack(anchor="w")
        self.logbox = tk.Text(frm, height=18, wrap="word")
        self.logbox.pack(fill="both", expand=True)

    def make_file_group(self, parent, title, command, store):
        lf = ttk.LabelFrame(parent, text=title, padding=8)
        lf.pack(fill="x", pady=6)
        row = ttk.Frame(lf); row.pack(fill="x")
        ttk.Button(row, text="ADD FILES", command=command).pack(side="left")
        ttk.Button(row, text="REMOVE ALL", command=lambda s=store:s.clear()).pack(side="left", padx=5)
        lb = tk.Listbox(lf, height=4)
        lb.pack(fill="x", pady=6)
        if "IA" in title:
            self.ia_listbox = lb
        else:
            self.result_listbox = lb

    def add_results(self):
        files = filedialog.askopenfilenames(title="Select Result PDF file(s)",
                                            filetypes=[("PDF files","*.pdf")])
        for f in files:
            if f not in self.result_files:
                self.result_files.append(f)
        self.refresh_lists()

    def add_ia(self):
        files = filedialog.askopenfilenames(title="Select IA PDF / Excel / CSV file(s)",
                                            filetypes=[("IA files","*.pdf *.xlsx *.xls *.csv"),
                                                       ("PDF","*.pdf"),("Excel","*.xlsx *.xls"),("CSV","*.csv")])
        for f in files:
            if f not in self.ia_files:
                self.ia_files.append(f)
        self.refresh_lists()

    def refresh_lists(self):
        self.result_listbox.delete(0,"end")
        for f in self.result_files: self.result_listbox.insert("end", f)
        self.ia_listbox.delete(0,"end")
        for f in self.ia_files: self.ia_listbox.insert("end", f)

    def clear(self):
        self.result_files.clear(); self.ia_files.clear()
        self.refresh_lists(); self.logbox.delete("1.0","end")
        self.status.config(text="Ready")

    def log(self, msg):
        self.logbox.insert("end", msg+"\n")
        self.logbox.see("end")
        self.root.update_idletasks()

    def start(self):
        if not self.result_files:
            messagebox.showwarning("Missing Result File","Add at least one result PDF.")
            return
        if not self.ia_files:
            messagebox.showwarning("Missing IA File","Add at least one IA PDF / Excel / CSV file.")
            return
        out = filedialog.asksaveasfilename(
            title="Save consolidated workbook",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook","*.xlsx")],
            initialfile="DPharm_BPharm_Result_IA_Analysis.xlsx")
        if not out:
            return
        self.progress.start(10)
        self.status.config(text="Processing...")
        self.logbox.delete("1.0","end")
        threading.Thread(target=self.worker, args=(out,), daemon=True).start()

    def worker(self, out):
        try:
            self.log("Reading result file(s)...")
            result, rlogs = parse_result_pdf(self.result_files)
            for x in rlogs: self.log(x)
            result = dedupe_result(result)
            self.log(f"Unique result subject records: {len(result)}")

            self.log("Reading IA file(s)...")
            ia, ilogs = parse_ia_files(self.ia_files)
            for x in ilogs: self.log(x)
            self.log(f"IA records: {len(ia)}")

            self.log("Running IA verification pass 1...")
            verified_ia, verification_conflicts = two_pass_verify_ia(ia)
            if verification_conflicts:
                self.progress.stop()
                self.status.config(text="EXPORT BLOCKED - VERIFICATION ERROR")
                self.log(f"TWO-PASS VERIFICATION FAILED: {len(verification_conflicts)} conflict(s)")
                for c in verification_conflicts[:30]:
                    self.log("  " + str(c))
                # Create a diagnostic workbook so the user can inspect errors.
                try:
                    export_excel(
                        out, [], [], [], [], [], [], result, verified_ia,
                        [], [], rlogs + ilogs,
                        verification_conflicts
                    )
                except Exception:
                    pass
                raise RuntimeError(
                    f"Two-pass verification failed for {len(verification_conflicts)} IA record/group(s). "
                    "Excel export of unverified data has been blocked. Check 'Two-Pass Verification'."
                )
            self.log(f"TWO-PASS VERIFICATION PASSED: {len(verified_ia)} verified IA records")
            ia = verified_ia

            fail = [r for r in result if r["Status"]=="FAIL"]
            absent = [r for r in result if r["Status"]=="ABSENT"]

            fail_ia, fail_log = attach_ia(fail, ia)
            absent_ia, absent_log = attach_ia(absent, ia)

            sem_rows, subj_rows = make_summaries(result)
            candidate_rows = build_candidate_summary(result)

            self.log(f"FAIL records: {len(fail)}")
            self.log(f"ABSENT records: {len(absent)}")
            self.log(f"PASS/FAIL subject summaries: {len(subj_rows)}")

            export_excel(
                out,
                sem_rows, subj_rows,
                fail, absent,
                fail_ia, absent_ia,
                result, ia,
                fail_log + absent_log,
                candidate_rows,
                rlogs + ilogs,
                [{"Verification": "PASSED", "Verified IA Records": len(ia), "Conflicts": 0}]
            )
            self.progress.stop()
            self.status.config(text=f"Completed: {out}")
            self.log(f"EXCEL CREATED: {out}")
            messagebox.showinfo("Completed", f"Consolidated workbook created:\n{out}")
        except Exception as e:
            self.progress.stop()
            self.status.config(text="ERROR")
            self.log("ERROR: "+str(e))
            self.log(traceback.format_exc())
            messagebox.showerror("Processing Error", str(e))

def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
