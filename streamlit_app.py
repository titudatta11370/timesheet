import streamlit as st
import pandas as pd
import datetime
import re
import io
import os
import tempfile
from difflib import SequenceMatcher
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Solar Edge Reconciliation",
    page_icon="📋",
    layout="centered"
)

st.title("📋 Solar Edge Roster Reconciliation")
st.markdown("Upload your Solar Edge email and timesheet Excel files to generate a reconciliation report.")

# ─── Email extraction ────────────────────────────────────────────────────────

def extract_email_text(uploaded_file):
    filename = uploaded_file.name.lower()
    if filename.endswith(".msg"):
        try:
            import extract_msg
            with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
                tmp.write(uploaded_file.read())
                tmp_path = tmp.name
            msg = extract_msg.openMsg(tmp_path)
            text = msg.body or ""
            msg.close()
            os.unlink(tmp_path)
            return text
        except ImportError:
            st.error("Missing library: extract-msg. Contact the app admin.")
            return ""
        except Exception as e:
            st.error(f"Could not read .msg file: {e}")
            return ""
    else:
        return uploaded_file.read().decode("utf-8", errors="ignore")

# ─── Helper functions ────────────────────────────────────────────────────────

def normalize_name(name):
    name = re.sub(r"[^a-zA-Z\s]", "", name).lower().strip()
    return " ".join(sorted(name.split()))

def fuzzy_match(name, candidates, threshold=0.75):
    key = normalize_name(name)
    best_score = 0
    best_match = None
    for candidate_key, data in candidates.items():
        score = SequenceMatcher(None, key, candidate_key).ratio()
        if score > best_score:
            best_score = score
            best_match = candidate_key
    if best_score >= threshold:
        return candidates[best_match], best_score
    return None, best_score

def parse_email_text(text):
    employees = []
    shift_date = None
    shift_name = None
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    for line in lines:
        m = re.search(r"Employee Summary for\s+(\d+/\d+/\d+),?\s*(.*?)(?:,\s*Simcoe|$)", line, re.IGNORECASE)
        if m:
            shift_date = m.group(1)
            shift_name = m.group(2).strip()
            continue
        parts = re.split(r"\t|\|", line)
        parts = [p.strip() for p in parts if p.strip()]
        if len(parts) < 3:
            continue
        if re.match(r"employee|emp\s*id|payable", parts[0], re.IGNORECASE):
            continue
        try:
            hours = float(parts[2])
        except ValueError:
            continue
        employees.append({
            "name": parts[0],
            "emp_id": parts[1] if len(parts) > 1 else "—",
            "roster_hours": hours,
        })
    return employees, shift_date, shift_name

def read_timesheets_from_files(uploaded_files):
    results = {}
    for uploaded_file in uploaded_files:
        try:
            xl = pd.read_excel(uploaded_file, sheet_name=None, header=None)
        except Exception as e:
            st.warning(f"Could not read {uploaded_file.name}: {e}")
            continue
        for sheet_name, df in xl.items():
            name = None
            hours = None
            try:
                name_val = df.iloc[8, 5]
                if pd.notna(name_val) and str(name_val).strip():
                    name = str(name_val).strip()
            except (IndexError, KeyError):
                pass
            if not name:
                name = sheet_name.strip()
            try:
                for i, row in df.iterrows():
                    for j, val in enumerate(row):
                        if "Total Hours Worked" in str(val):
                            for c in range(len(df.columns)):
                                cell_val = df.iloc[i, c]
                                if isinstance(cell_val, datetime.time) and (cell_val.hour > 0 or cell_val.minute > 0):
                                    hours = cell_val.hour + cell_val.minute / 60 + cell_val.second / 3600
                                    break
                                elif isinstance(cell_val, datetime.timedelta):
                                    hours = cell_val.total_seconds() / 3600
                                    break
                                elif isinstance(cell_val, (int, float)) and not pd.isna(cell_val) and cell_val > 0:
                                    hours = float(cell_val)
                                    break
                            break
                    if hours is not None:
                        break
            except Exception:
                pass
            if name:
                key = normalize_name(name)
                results[key] = {
                    "name": name,
                    "hours": hours,
                    "file": uploaded_file.name,
                    "sheet": sheet_name,
                }
    return results

def reconcile(roster, timesheets):
    results = []
    for emp in roster:
        ts_data, score = fuzzy_match(emp["name"], timesheets)
        if ts_data and ts_data["hours"] is not None:
            ts_hours = ts_data["hours"]
            diff = ts_hours - emp["roster_hours"]
            status = "MATCH" if abs(diff) < 0.01 else "DISCREPANCY"
            results.append({**emp, "ts_hours": ts_hours, "diff": diff,
                           "status": status, "ts_file": ts_data["file"],
                           "match_score": round(score, 2), "matched_name": ts_data["name"]})
        else:
            results.append({**emp, "ts_hours": None, "diff": None,
                           "status": "MISSING TIMESHEET",
                           "ts_file": ts_data["file"] if ts_data else "—",
                           "match_score": round(score, 2),
                           "matched_name": ts_data["name"] if ts_data else "—"})
    return results

def build_excel_report(results, shift_date=None, shift_name=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reconciliation"

    GREEN = "C6EFCE"; RED = "FFC7CE"; AMBER = "FFEB9C"
    HEADER_BG = "1F3864"; HEADER_FG = "FFFFFF"
    SUBHEADER_BG = "D9E1F2"
    MATCH_FG = "276221"; DISC_FG = "9C0006"; MISS_FG = "7D4C00"

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def cs(row, col, value, bold=False, bg=None, fg=None, align="left", num_fmt=None):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(bold=bold, color=fg or "000000", name="Calibri", size=11)
        if bg:
            c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal=align, vertical="center")
        c.border = border
        if num_fmt:
            c.number_format = num_fmt
        return c

    ws.merge_cells("A1:J1")
    title = "Solar Edge Roster Reconciliation"
    if shift_date:
        title += f"  —  {shift_date}"
    if shift_name:
        title += f"  ({shift_name})"
    c = ws["A1"]
    c.value = title
    c.font = Font(bold=True, color=HEADER_FG, name="Calibri", size=13)
    c.fill = PatternFill("solid", fgColor=HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    total = len(results)
    matched = sum(1 for r in results if r["status"] == "MATCH")
    discs = sum(1 for r in results if r["status"] == "DISCREPANCY")
    missing = sum(1 for r in results if r["status"] == "MISSING TIMESHEET")

    ws.merge_cells("A2:J2")
    summary = f"Total: {total}   |   Matched: {matched}   |   Discrepancies: {discs}   |   Missing: {missing}"
    c = ws["A2"]
    c.value = summary
    c.font = Font(color="1F3864", name="Calibri", size=10)
    c.fill = PatternFill("solid", fgColor=SUBHEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")

    headers = ["Employee", "Emp ID", "Roster Hrs", "Timesheet Hrs", "Difference",
               "Status", "Matched Name", "Match Score", "Timesheet File", "Notes"]
    for col, h in enumerate(headers, 1):
        cs(3, col, h, bold=True, bg=HEADER_BG, fg=HEADER_FG, align="center")
    ws.row_dimensions[3].height = 22

    for row_i, r in enumerate(results, 4):
        ws.row_dimensions[row_i].height = 18
        status = r["status"]
        row_bg = GREEN if status == "MATCH" else RED if status == "DISCREPANCY" else AMBER
        status_fg = MATCH_FG if status == "MATCH" else DISC_FG if status == "DISCREPANCY" else MISS_FG

        cs(row_i, 1, r["name"], bg=row_bg)
        cs(row_i, 2, r["emp_id"], bg=row_bg, align="center")
        cs(row_i, 3, r["roster_hours"], bg=row_bg, align="center", num_fmt="0.00")
        ts_val = r["ts_hours"] if r["ts_hours"] is not None else "—"
        cs(row_i, 4, ts_val, bg=row_bg, align="center", num_fmt="0.00" if r["ts_hours"] is not None else None)
        diff_val = r["diff"]
        if diff_val is not None:
            diff_str = f"+{diff_val:.2f}" if diff_val > 0 else ("0.00" if abs(diff_val) < 0.01 else f"{diff_val:.2f}")
        else:
            diff_str = "—"
        cs(row_i, 5, diff_str, bg=row_bg, align="center")
        cs(row_i, 6, status, bold=True, bg=row_bg, fg=status_fg, align="center")
        cs(row_i, 7, r.get("matched_name", "—"), bg=row_bg)
        cs(row_i, 8, r.get("match_score", "—"), bg=row_bg, align="center")
        cs(row_i, 9, r.get("ts_file", "—"), bg=row_bg)
        cs(row_i, 10, "", bg=row_bg)

    widths = [28, 12, 13, 15, 12, 22, 28, 13, 30, 20]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ─── UI ──────────────────────────────────────────────────────────────────────

st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📧 Step 1 — Solar Edge Email")
    email_file = st.file_uploader(
        "Upload the email as .msg or .txt",
        type=["txt", "msg"],
        help="In Outlook: File → Save As → Outlook Message (.msg) or Text Only (.txt)"
    )

with col2:
    st.subheader("📊 Step 2 — Timesheet Excel Files")
    timesheet_files = st.file_uploader(
        "Upload one or more timesheet .xlsx files",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Upload all the timesheet Excel files for this week"
    )

st.markdown("---")

if st.button("▶ Run Reconciliation", type="primary", use_container_width=True):
    if not email_file:
        st.error("Please upload the Solar Edge email .txt file first.")
    elif not timesheet_files:
        st.error("Please upload at least one timesheet Excel file.")
    else:
        with st.spinner("Running reconciliation..."):
            # Parse email
            email_text = extract_email_text(email_file)
            roster, shift_date, shift_name = parse_email_text(email_text)

            if not roster:
                st.error("Could not find any employees in the email. Check the file format.")
            else:
                # Read timesheets
                timesheets = read_timesheets_from_files(timesheet_files)

                # Reconcile
                results = reconcile(roster, timesheets)

                # Summary metrics
                matched = [r for r in results if r["status"] == "MATCH"]
                discs = [r for r in results if r["status"] == "DISCREPANCY"]
                missing = [r for r in results if r["status"] == "MISSING TIMESHEET"]

                st.success("✅ Reconciliation complete!")

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Total Employees", len(results))
                m2.metric("✅ Matched", len(matched))
                m3.metric("⚠️ Discrepancies", len(discs))
                m4.metric("❌ Missing Timesheet", len(missing))

                st.markdown("---")

                # Results table
                if discs:
                    st.subheader("⚠️ Discrepancies")
                    disc_df = pd.DataFrame([{
                        "Employee": r["name"],
                        "Emp ID": r["emp_id"],
                        "Roster Hrs": r["roster_hours"],
                        "Timesheet Hrs": r["ts_hours"],
                        "Difference": f"+{r['diff']:.2f}" if r["diff"] > 0 else f"{r['diff']:.2f}"
                    } for r in discs])
                    st.dataframe(disc_df, use_container_width=True, hide_index=True)

                if missing:
                    st.subheader("❌ Missing Timesheets")
                    miss_df = pd.DataFrame([{
                        "Employee": r["name"],
                        "Emp ID": r["emp_id"],
                        "Roster Hrs": r["roster_hours"],
                    } for r in missing])
                    st.dataframe(miss_df, use_container_width=True, hide_index=True)

                if matched:
                    with st.expander(f"✅ View {len(matched)} matched employees"):
                        match_df = pd.DataFrame([{
                            "Employee": r["name"],
                            "Emp ID": r["emp_id"],
                            "Hours": r["roster_hours"],
                        } for r in matched])
                        st.dataframe(match_df, use_container_width=True, hide_index=True)

                # Download report
                st.markdown("---")
                today = datetime.date.today().strftime("%Y-%m-%d")
                excel_report = build_excel_report(results, shift_date, shift_name)
                st.download_button(
                    label="⬇️ Download Excel Report",
                    data=excel_report,
                    file_name=f"reconciliation_{today}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
