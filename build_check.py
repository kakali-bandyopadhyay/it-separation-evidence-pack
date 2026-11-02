#!/usr/bin/env python3
"""Independent workbook reconciliations and one-page PDF generation."""
from __future__ import annotations

import csv
import math
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, TableStyle
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
XLSX = ROOT / "it_separation_evidence_pack.xlsx"
PDF = ROOT / "it_separation_one_pager.pdf"
TOL = 0.02


def read_csv(name):
    with (DATA / name).open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def as_int(v):
    return int(float(v))


def as_num(v):
    return float(v)


def as_date(v):
    return date.fromisoformat(v)


def month_count(start, end):
    return (end.year - start.year) * 12 + end.month - start.month + 1


def near(actual, expected, label):
    if actual is None or not math.isclose(float(actual), float(expected), rel_tol=0, abs_tol=TOL):
        raise AssertionError(f"{label}: workbook={actual!r}, independently computed={expected!r}")


def text_equal(actual, expected, label):
    if actual != expected:
        raise AssertionError(f"{label}: workbook={actual!r}, expected={expected!r}")


def excel_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value)[:10])


def formula_cell(wb, sheet, cell):
    formula = wb[sheet][cell].value
    if not isinstance(formula, str) or not formula.startswith("="):
        raise AssertionError(f"Expected live formula at {sheet}!{cell}; got {formula!r}")


apps = read_csv("apps.csv")
tsa = read_csv("tsa.csv")
day1 = read_csv("day1.csv")
synergies = read_csv("synergies.csv")
risks = read_csv("risks.csv")

wb = load_workbook(XLSX, data_only=True, read_only=False)
wf = load_workbook(XLSX, data_only=False, read_only=False)
required_tabs = ["README_Assumptions", "App_Disposition", "TSA_Schedule", "Exit_Plan", "Day1_Checklist", "Cost_Synergy", "Checks", "Dashboard", "PMI_100Day", "DD_Redflags"]
text_equal(wb.sheetnames, required_tabs, "Workbook tab order/names")

# Source blocks preserve the supplied Alder, PMI, risk, synergy and outsourcing CSV rows in meaning.
def check_source_block(sheet_name, csv_name, header_row):
    source = read_csv(csv_name)
    ws = wb[sheet_name]
    text_equal([ws.cell(header_row, col).value for col in range(1, len(source[0]) + 1)], list(source[0].keys()), f"{sheet_name} source headers {csv_name}")
    for row_idx, record in enumerate(source, header_row + 1):
        for col_idx, key in enumerate(record, 1):
            actual = ws.cell(row_idx, col_idx).value
            expected = record[key]
            label = f"{sheet_name} source {csv_name} row {row_idx} column {key}"
            if expected == "":
                text_equal(actual, None, label)
            elif expected.lstrip("-").replace(".", "", 1).isdigit():
                near(actual, float(expected), label)
            else:
                text_equal(str(actual), expected, label)

dd_rows = read_csv("alder_redflags.csv")
dd_domains = read_csv("alder_domains.csv")
dd_questions = read_csv("alder_questions.csv")
for i, row in enumerate(dd_domains, 9):
    for col, key in enumerate(["domain", "rag", "one_line_reason"], 1):
        text_equal(wb["DD_Redflags"].cell(i, col).value, row[key], f"DD domain {row['domain']} {key}")
for i, row in enumerate(dd_rows, 19):
    for col, key in enumerate(["rf_id", "finding", "why_it_matters", "value_at_risk_low_gbp", "value_at_risk_high_gbp", "recommended_action"], 1):
        actual = wb["DD_Redflags"].cell(i, col).value
        if key.endswith("_gbp"):
            near(actual, float(row[key]), f"DD red flag {row['rf_id']} {key}")
        else:
            text_equal(str(actual), row[key], f"DD red flag {row['rf_id']} {key}")
for i, row in enumerate(dd_questions, 27):
    for col, key in enumerate(["q_id", "question", "linked_rf_id"], 1):
        text_equal(wb["DD_Redflags"].cell(i, col).value, row[key], f"DD question {row['q_id']} {key}")
formula_cell(wf, "DD_Redflags", "B6")
near(wb["DD_Redflags"]["B6"].value, sum(float(wb["App_Disposition"].cell(r, 19).value) for r in range(5, 65)) / 60, "DD complexity proxy")
check_source_block("PMI_100Day", "pmi_milestones.csv", 95)
check_source_block("PMI_100Day", "risks.csv", 129)
check_source_block("PMI_100Day", "synergies.csv", 143)
check_source_block("PMI_100Day", "outsourcing.csv", 162)

# Add-on A independently validates milestone and risk links, formula-driven lists and scorecard.
pmi_s = wb["PMI_100Day"]
pmi_f = wf["PMI_100Day"]
pmi_rows = read_csv("pmi_milestones.csv")
risk_rows = read_csv("risks.csv")
outsourcing_rows = read_csv("outsourcing.csv")
exit_ids = {pmi_s.parent["Exit_Plan"].cell(r, 1).value for r in range(5, 40) if pmi_s.parent["Exit_Plan"].cell(r, 1).value}
known_dependency_ids = {r["app_id"] for r in apps} | {r["svc_id"] for r in tsa} | exit_ids | {r["ms_id"] for r in pmi_rows}
for i, row in enumerate(pmi_rows, 20):
    text_equal(pmi_s.cell(i, 1).value, row["ms_id"], f"PMI milestone {row['ms_id']}")
    if row["dependency_id"] not in known_dependency_ids:
        raise AssertionError(f"Unresolved PMI dependency {row['ms_id']} -> {row['dependency_id']}")
    text_equal(pmi_s.cell(i, 9).value, "PASS", f"PMI dependency check {row['ms_id']}")
    if pmi_s.cell(i, 8).value not in {"Green", "Amber", "Red"}:
        raise AssertionError(f"Invalid dependency RAG {row['ms_id']}: {pmi_s.cell(i, 8).value!r}")
    for col in range(1, 8):
        formula_cell(wf, "PMI_100Day", f"{chr(64 + col)}{i}")
    formula_cell(wf, "PMI_100Day", f"H{i}")
    formula_cell(wf, "PMI_100Day", f"I{i}")

for i, row in enumerate(synergies, 53):
    text_equal(pmi_s.cell(i, 1).value, row["syn_id"], f"PMI synergy {row['syn_id']}")
    text_equal(pmi_s.cell(i, 2).value, row["lever"], f"PMI synergy lever {row['syn_id']}")
    text_equal(pmi_s.cell(i, 4).value, row["linked_app_ids"], f"PMI synergy links {row['syn_id']}")
    near(pmi_s.cell(i, 6).value, (as_num(row["annual_value_low_gbp"]) + as_num(row["annual_value_high_gbp"])) / 2, f"PMI synergy midpoint {row['syn_id']}")
    near(pmi_s.cell(i, 7).value, as_num(row["confidence_pct"]) / 100, f"PMI synergy confidence {row['syn_id']}")
    text_equal(pmi_s.cell(i, 10).value, "PASS", f"PMI synergy tie {row['syn_id']}")
    formula_cell(wf, "PMI_100Day", f"J{i}")

for i, row in enumerate(outsourcing_rows, 72):
    text_equal(pmi_s.cell(i, 1).value, row["out_id"], f"PMI outsourcing option {row['out_id']}")
    text_equal(pmi_s.cell(i, 2).value, row["scope"], f"PMI outsourcing scope {row['out_id']}")
    near(pmi_s.cell(i, 4).value, as_num(row["saving_low_gbp"]), f"PMI outsourcing low range {row['out_id']}")
    near(pmi_s.cell(i, 5).value, as_num(row["saving_high_gbp"]), f"PMI outsourcing high range {row['out_id']}")
    formula_cell(wf, "PMI_100Day", f"A{i}")
    formula_cell(wf, "PMI_100Day", f"F{i}")

for i, row in enumerate(risk_rows, 81):
    text_equal(pmi_s.cell(i, 1).value, row["risk_id"], f"PMI risk {row['risk_id']}")
    text_equal(pmi_s.cell(i, 3).value, row["linked_id"], f"PMI risk linked_id {row['risk_id']}")
    if row["linked_id"] not in known_dependency_ids:
        raise AssertionError(f"Unresolved risk linked_id {row['risk_id']} -> {row['linked_id']}")
    text_equal(pmi_s.cell(i, 8).value, "PASS", f"PMI risk linked_id check {row['risk_id']}")
    formula_cell(wf, "PMI_100Day", f"H{i}")

for cell in ["B12", "B13", "B14", "B15"]:
    formula_cell(wf, "PMI_100Day", cell)
near(pmi_s["B12"].value, sum(1 for r in range(20, 50) if pmi_s.cell(r, 8).value != "Red") / len(pmi_rows), "PMI on-track proxy")
near(pmi_s["B13"].value, 1, "PMI identified synergy vs modeled target")
near(pmi_s["B14"].value, sum(as_num(r["annual_value_low_gbp"]) + as_num(r["annual_value_high_gbp"]) for r in synergies) / 2, "PMI identified annual synergy midpoint")
near(pmi_s["B15"].value, wb["Dashboard"]["E8"].value, "PMI open No-go items")

for row in range(193, 197):
    text_equal(wb["Checks"].cell(row, 2).value, "PASS", f"PMI Check tab row {row}")
    formula_cell(wf, "Checks", f"B{row}")

# Source app count, derived disposition, complexity, and app one-off cost totals.
app_s = wb["App_Disposition"]
apps_by_id = {r["app_id"]: r for r in apps}
tsa_by_id = {r["svc_id"]: r for r in tsa}
if len(apps_by_id) != len(apps):
    raise AssertionError("apps.csv contains duplicate app_id values")
linked_app_ids = {token.strip() for row in tsa for token in row["linked_app_ids"].split(";") if token.strip()}
app_disposition_counts = Counter()
app_cost_by_disp = defaultdict(float)
app_cost_by_wave = defaultdict(float)
app_run_cost = sum(as_num(a["annual_run_cost_gbp"]) for a in apps)
for rowno, app in enumerate(apps, 5):
    text_equal(app_s.cell(rowno, 1).value, app["app_id"], f"App ID row {rowno}")
    ent = app["data_entanglement"]
    lic = app["licence_transferable"]
    shared = app["shared_with_seller"] == "Y"
    expected_disp = "TSA-then-replace" if shared else ("Transfer" if lic == "Y" and ent == "Low" else "Clone" if lic == "Consent" else "Replicate" if ent == "High" else "Replace")
    text_equal(app_s.cell(rowno, 15).value, expected_disp, f"Suggested disposition {app['app_id']}")
    expected_final = "Retire" if app["app_id"] == "APP010" else expected_disp
    text_equal(app_s.cell(rowno, 20).value, expected_final, f"Final disposition {app['app_id']}")
    if app["app_id"] == "APP010":
        text_equal(app_s.cell(rowno, 10).value, 2, "APP010 criticality")
        text_equal(expected_disp, "Clone", "APP010 original suggested disposition")
        text_equal(app_s.cell(rowno, 16).value, "Retire", "APP010 synthetic override")
        text_equal(app_s.cell(rowno, 17).value, "Synthetic example: function absorbed by buyer platform", "APP010 override rationale")
    elif app_s.cell(rowno, 16).value not in (None, ""):
        raise AssertionError(f"Unexpected override for {app['app_id']}")
    expected_complexity = min(5, 1 + (2 if ent == "High" else 1 if ent == "Med" else 0) + (1 if lic == "Consent" else 0) + (1 if shared else 0))
    near(app_s.cell(rowno, 19).value, expected_complexity, f"Complexity {app['app_id']}")
    expected_tsa = "N" if expected_final == "Retire" else "Y" if app["app_id"] in linked_app_ids else "N"
    text_equal(app_s.cell(rowno, 18).value, expected_tsa, f"TSA needed {app['app_id']}")
    expected_wave = app["wave"]
    text_equal(app_s.cell(rowno, 14).value, expected_wave, f"Final wave {app['app_id']}")
    app_disposition_counts[expected_final] += 1
    app_cost_by_disp[expected_final] += as_num(app["one_off_cost_gbp"])
    app_cost_by_wave[expected_wave] += as_num(app["one_off_cost_gbp"])
for disp, rowno in zip(["Transfer", "Clone", "Replicate", "Replace", "Retire", "TSA-then-replace"], range(69, 75)):
    text_equal(app_s.cell(rowno, 1).value, disp, f"Disposition summary label {disp}")
    near(app_s.cell(rowno, 2).value, app_disposition_counts[disp], f"Disposition count {disp}")
    near(app_s.cell(rowno, 3).value, app_cost_by_disp[disp], f"Disposition cost {disp}")
near(app_s["B77"].value, 0, "Day 1 critical applications without plan")

# TSA lines, price/extension math, and month-by-month run-off on both directions.
tsa_s = wb["TSA_Schedule"]
markup = as_num(wb["README_Assumptions"]["B14"].value)
stepup = as_num(wb["README_Assumptions"]["B15"].value)
tsa_cost = 0.0
reverse_cost = 0.0
buyer_monthly = defaultdict(float)
reverse_monthly = defaultdict(float)
buyer_cutover = {}
for i, item in enumerate(tsa, 5):
    start, end = as_date(item["start_date"]), as_date(item["planned_end_date"])
    months = month_count(start, end)
    basis = as_num(item["cost_basis_gbp_month"])
    monthly = basis * (1 + markup)
    extension_months = as_int(item["extension_option_months"])
    extension_monthly = monthly * (1 + stepup)
    line_total = monthly * months
    ext_total = line_total + extension_monthly * extension_months
    near(tsa_s.cell(i, 7).value, months, f"TSA months {item['svc_id']}")
    near(tsa_s.cell(i, 8).value, monthly, f"TSA monthly fee {item['svc_id']}")
    near(tsa_s.cell(i, 13).value, extension_monthly, f"TSA step-up fee {item['svc_id']}")
    near(tsa_s.cell(i, 14).value, line_total, f"TSA total line value {item['svc_id']}")
    near(tsa_s.cell(i, 15).value, ext_total, f"TSA extended line value {item['svc_id']}")
    if item["direction"] == "Seller to buyer":
        tsa_cost += line_total
        buyer_cutover[item["svc_id"]] = end - timedelta(days=30)
        cursor = date(start.year, start.month, 1)
        while cursor <= end:
            buyer_monthly[(cursor.year, cursor.month)] += monthly
            cursor = date(cursor.year + (cursor.month == 12), 1 if cursor.month == 12 else cursor.month + 1, 1)
    else:
        reverse_cost += line_total
        cursor = date(start.year, start.month, 1)
        while cursor <= end:
            reverse_monthly[(cursor.year, cursor.month)] += monthly
            cursor = date(cursor.year + (cursor.month == 12), 1 if cursor.month == 12 else cursor.month + 1, 1)
near(tsa_s["R5"].value, tsa_cost, "Buyer TSA total")
near(tsa_s["R6"].value, max(buyer_monthly.values()), "Peak buyer monthly run-rate")
near(tsa_s["R7"].value, reverse_cost, "Reverse line scheduled value")
near(tsa_s["G44"].value, tsa_cost, "Buyer TSA runoff total")
for m in range(24):
    year = 2027 + (3 + m) // 12
    month = (3 + m) % 12 + 1
    rr = 45 + m
    near(tsa_s.cell(rr, 2).value, buyer_monthly[(year, month)], f"Buyer TSA runoff {year}-{month:02d}")
    near(tsa_s.cell(rr, 3).value, reverse_monthly[(year, month)], f"Reverse TSA runoff {year}-{month:02d}")

# Exit milestone formulas, dependency float/RAG, and date statuses.
exit_s = wb["Exit_Plan"]
asof = excel_date(wb["README_Assumptions"]["E13"].value)
day_one = excel_date(wb["README_Assumptions"]["B13"].value)
cut_lead = as_int(wb["README_Assumptions"]["E14"].value)
test_lead = as_int(wb["README_Assumptions"]["E15"].value)
build_lead = as_int(wb["README_Assumptions"]["E16"].value)
design_lead = as_int(wb["README_Assumptions"]["E17"].value)
decom_lag = as_int(wb["README_Assumptions"]["E18"].value)
buyer_lines = [t for t in tsa if t["direction"] == "Seller to buyer"]
exit_row_by_id = {exit_s.cell(r, 1).value: r for r in range(5, 5 + len(buyer_lines))}
for i, item in enumerate(buyer_lines, 5):
    end = as_date(item["planned_end_date"])
    cut = end - timedelta(days=cut_lead)
    test = cut - timedelta(days=test_lead)
    build = test - timedelta(days=build_lead)
    design_raw = build - timedelta(days=design_lead)
    design = max(day_one, design_raw)
    decommission = end + timedelta(days=decom_lag)
    for col, expected, label in [(4, design, "design"), (6, build, "build"), (8, test, "test"), (10, cut, "cutover"), (12, decommission, "decommission")]:
        text_equal(excel_date(exit_s.cell(i, col).value), expected, f"{item['svc_id']} {label} date")
    for date_col, status_col in [(4, 5), (6, 7), (8, 9), (10, 11), (12, 13)]:
        expected_status = "Complete" if excel_date(exit_s.cell(i, date_col).value) <= asof else "Planned"
        text_equal(exit_s.cell(i, status_col).value, expected_status, f"{item['svc_id']} milestone status")
    dep = item["exit_dependency"]
    if dep in buyer_cutover:
        float_days = (cut - buyer_cutover[dep]).days
        critical = "Yes"
    else:
        float_days = (end - cut).days
        critical = "No"
    near(exit_s.cell(i, 14).value, float_days, f"{item['svc_id']} float")
    text_equal(exit_s.cell(i, 15).value, critical, f"{item['svc_id']} critical path")
    expected_rag = "Red" if float_days < as_int(wb["README_Assumptions"]["E19"].value) else "Amber" if float_days <= as_int(wb["README_Assumptions"]["E20"].value) else "Green"
    text_equal(exit_s.cell(i, 16).value, expected_rag, f"{item['svc_id']} float RAG")
    text_equal(exit_s.cell(i, 17).value, "Yes" if design_raw < day_one else "No", f"{item['svc_id']} design clamp")

# Day 1 readiness and open No-go totals.
day_s = wb["Day1_Checklist"]
workstream_rows = defaultdict(list)
for i, item in enumerate(day1, 5):
    workstream_rows[item["workstream"]].append(item)
workstream_first = 6
for offset, (workstream, entries) in enumerate(workstream_rows.items()):
    r = workstream_first + offset
    expected_ready = sum(e["status"] == "Complete" for e in entries) / len(entries)
    expected_nogo = sum(e["gate"] == "No-go" and e["status"] != "Complete" for e in entries)
    near(day_s.cell(r, 12).value, expected_ready, f"Day 1 readiness {workstream}")
    near(day_s.cell(r, 13).value, expected_nogo, f"Open No-go {workstream}")
    for itemrow in range(5, 5 + len(day1)):
        item = day1[itemrow - 5]
        if item["workstream"] == workstream:
            expected_rag = "Green" if item["gate"] == "Go" and item["status"] == "Complete" else "Red" if item["gate"] == "No-go" and item["status"] != "Complete" or item["status"] == "Not started" else "Amber"
            text_equal(day_s.cell(itemrow, 9).value, expected_rag, f"Day1 RAG {item['item_id']}")
near(day_s.cell(6 + len(workstream_rows), 12).value, sum(e["status"] == "Complete" for e in day1) / len(day1), "Total Day 1 readiness")
near(day_s.cell(6 + len(workstream_rows), 13).value, sum(e["gate"] == "No-go" and e["status"] != "Complete" for e in day1), "Total open No-go")

# Cost baseline, wave one-offs, synergy tracker, and bridge.
cost_s = wb["Cost_Synergy"]
check_s = wb["Checks"]
alloc = [app_run_cost, 2100000, 900000, 700000, 800000, 300000]
standalone = [app_run_cost * 1.08, 2600000, 1150000, 850000, 900000, 550000]
for i, (a, s) in enumerate(zip(alloc, standalone), 6):
    near(cost_s.cell(i, 2).value, a, f"Allocated cost row {i}")
    near(cost_s.cell(i, 3).value, s, f"Standalone cost row {i}")
    near(cost_s.cell(i, 4).value, s - a, f"Dis-synergy row {i}")
near(cost_s["B12"].value, sum(alloc), "Allocated cost total")
near(cost_s["C12"].value, sum(standalone), "Standalone cost total")
near(cost_s["D12"].value, sum(s - a for a, s in zip(alloc, standalone)), "Dis-synergy total")
wave_rows = {cost_s.cell(r, 6).value: r for r in range(6, 6 + len({a['wave'] for a in apps}))}
for wave, r in wave_rows.items():
    near(cost_s.cell(r, 7).value, app_cost_by_wave[wave], f"One-off cost wave {wave}")

annual_low = [as_num(s["annual_value_low_gbp"]) for s in synergies]
annual_high = [as_num(s["annual_value_high_gbp"]) for s in synergies]
annual_midpoint = [(lo + hi) / 2 for lo, hi in zip(annual_low, annual_high)]
for i, value in enumerate(annual_midpoint, 6):
    near(cost_s.cell(i, 17).value, value, f"Synergy midpoint {synergies[i-6]['syn_id']}")
planned_months = []
asof_month = as_int(wb["README_Assumptions"]["B28"].value)
actual_split_month = as_int(wb["README_Assumptions"]["E21"].value)
for month in range(1, 25):
    planned = 0.0
    for s, annual in zip(synergies, annual_midpoint):
        start = as_int(s["start_month"])
        ramp = as_int(s["ramp_months"])
        if month >= start:
            factor = 1.0 if ramp == 0 else min(1.0, (month - start + 1) / ramp)
            planned += annual / 12 * factor
    planned_months.append(planned)
    r = 17 + month
    near(cost_s.cell(r, 3).value, planned, f"Synergy planned month {month}")
    if month <= asof_month:
        realization = 0.85 if month <= actual_split_month else 0.92
        actual = planned * realization
        near(cost_s.cell(r, 4).value, actual, f"Synthetic actual month {month}")
        near(cost_s.cell(r, 5).value, actual - planned, f"Variance month {month}")
        if planned:
            ratio = actual / planned
            rag = "Green" if ratio >= 0.95 - 1e-12 else "Amber" if ratio >= 0.85 - 1e-12 else "Red"
            text_equal(cost_s.cell(r, 6).value, rag, f"Synergy RAG month {month}")
        else:
            text_equal(cost_s.cell(r, 6).value, None, f"Synergy RAG month {month} with no planned value")
    elif cost_s.cell(r, 4).value not in (None, ""):
        raise AssertionError(f"Synthetic actual should be blank after as-of month {month}")
near(cost_s["I18"].value, sum(alloc), "Bridge A allocated baseline")
near(cost_s["I19"].value, sum(s - a for a, s in zip(alloc, standalone)), "Bridge A dis-synergy")
near(cost_s["I20"].value, sum(standalone), "Bridge A standalone annual run-rate")
near(cost_s["I21"].value, sum(annual_midpoint), "Bridge A identified annual synergies")
near(cost_s["I22"].value, sum(standalone) - sum(annual_midpoint), "Bridge A target run-rate")
near(cost_s["I26"].value, app_cost_by_wave["Day 1"], "Bridge B Day 1 one-off")
near(cost_s["I27"].value, app_cost_by_wave["Wave 2"], "Bridge B Wave 2 one-off")
near(cost_s["I28"].value, app_cost_by_wave["Wave 3"], "Bridge B Wave 3 one-off")
near(cost_s["I29"].value, tsa_cost, "Bridge B planned buyer TSA")
near(cost_s["I30"].value, sum(as_num(a["one_off_cost_gbp"]) for a in apps), "Bridge B one-off separation cost")
near(cost_s["I31"].value, sum(as_num(a["one_off_cost_gbp"]) for a in apps) + tsa_cost, "Bridge B one-time cost to separate")
for rowno, item in enumerate(tsa, 5):
    near(tsa_s.cell(rowno, 19).value, month_count(as_date(item["start_date"]), as_date(item["planned_end_date"])), f"Independent date months {item['svc_id']}")
if "G" in check_s["B6"].value or "H" in check_s["B6"].value or "N" in check_s["B6"].value:
    raise AssertionError("Checks!B6 must not depend on TSA Months, Monthly Fee, or line total columns")
for row in range(174, 190):
    for col in (3, 4):
        if check_s.cell(row, col).number_format != "yyyy-mm-dd":
            raise AssertionError(f"Checks!{check_s.cell(row, col).coordinate} date format is not yyyy-mm-dd")
override_validation = app_s.data_validations.dataValidation
if not any("P5:P64" in str(v.sqref) and "Transfer" in str(v.formula1) and "TSA-then-replace" in str(v.formula1) for v in override_validation):
    raise AssertionError("Override validation list is missing or incomplete")

# Source link validity, dependency timing, and workbook formula presence.
app_ids = set(apps_by_id)
tsa_ids = set(tsa_by_id)
for s in synergies:
    missing = set(s["linked_app_ids"].split(";")) - app_ids
    if missing:
        raise AssertionError(f"Missing synergy linked app IDs: {sorted(missing)}")
for t in tsa:
    missing = set(t["linked_app_ids"].split(";")) - app_ids
    if missing:
        raise AssertionError(f"Missing TSA linked app IDs: {sorted(missing)}")
    if t["exit_dependency"] not in app_ids | tsa_ids:
        raise AssertionError(f"Unresolved TSA dependency on {t['svc_id']}: {t['exit_dependency']}")
for r in risks:
    if r["linked_id"] not in app_ids | tsa_ids:
        raise AssertionError(f"Unresolved risk link on {r['risk_id']}: {r['linked_id']}")
for t in buyer_lines:
    dep = t["exit_dependency"]
    if dep in buyer_cutover and as_date(t["planned_end_date"]) < buyer_cutover[dep]:
        raise AssertionError(f"TSA {t['svc_id']} ends before dependency {dep} cutover")
check_results = [check_s.cell(r, 2).value for r in range(5, 14)]
if check_results != ["PASS"] * 9:
    raise AssertionError(f"Checks tab is not all PASS: {check_results}")
formula_errors = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!"}
for sheet in wb.worksheets:
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value in formula_errors:
                raise AssertionError(f"Formula error in cached workbook value {sheet.title}!{cell.coordinate}: {cell.value}")
for sheet, cells in {
    "TSA_Schedule": ["G5", "H5", "N5", "O5", "R5", "B45", "G44"],
    "App_Disposition": ["O5", "R5", "S5", "T5", "B69", "C69"],
    "Cost_Synergy": ["B6", "C6", "D6", "Q6", "C18", "D18", "I18", "I22", "I26", "I30", "I31"],
    "Checks": ["B5", "B6", "B10", "B12", "B13"],
}.items():
    for cell in cells:
        formula_cell(wf, sheet, cell)


def money(v):
    return "GBP {:,.0f}".format(float(v or 0))


def draw_table(c, data, x, top, widths, row_heights=None, header=True, font_name="Calibri"):
    t = Table(data, colWidths=widths, rowHeights=row_heights, repeatRows=1 if header else 0)
    commands = [
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, -1), 7.2),
        ("LEADING", (0, 0), (-1, -1), 8.2),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#243247")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#D8DEE8")),
    ]
    if header:
        commands += [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3864")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("FONTNAME", (0, 0), (-1, 0), font_name), ("FONTSIZE", (0, 0), (-1, 0), 7.2)]
    t.setStyle(TableStyle(commands))
    _, h = t.wrapOn(c, sum(widths), 1000)
    t.drawOn(c, x, top - h)
    return top - h


def draw_paragraph(c, text, x, top, width, size=8, color="#243247", leading=None, font_name="Calibri"):
    pstyle = ParagraphStyle("p", fontName=font_name, fontSize=size, leading=leading or size * 1.18, textColor=colors.HexColor(color), alignment=TA_LEFT)
    p = Paragraph(text, pstyle)
    _, h = p.wrap(width, 500)
    p.drawOn(c, x, top - h)
    return top - h


def build_pdf():
    # Pull every displayed numeric value from cached workbook outputs.
    d = wb["Dashboard"]
    disp_s = wb["App_Disposition"]
    tsa_data = wb["TSA_Schedule"]
    day_data = wb["Day1_Checklist"]
    cost_data = wb["Cost_Synergy"]
    app_count = d["B5"].value
    critical_gaps = d["B6"].value
    readiness = d["B7"].value
    buyer_total = d["B8"].value
    peak = d["B9"].value
    synergy_mid = d["B10"].value
    one_off_cost = d["E5"].value
    annual_target = d["E9"].value
    one_time_total = d["E10"].value
    if any(v is None for v in [app_count, critical_gaps, readiness, buyer_total, peak, synergy_mid, one_off_cost, annual_target, one_time_total]):
        raise AssertionError("Dashboard values are not cached; cannot create workbook-sourced PDF")
    font_name = "Helvetica"
    calibri_path = Path(r"C:\Windows\Fonts\calibri.ttf")
    if calibri_path.exists():
        pdfmetrics.registerFont(TTFont("Calibri", str(calibri_path)))
        font_name = "Calibri"
    pdfmetrics.registerFontFamily(font_name, normal=font_name, bold=font_name, italic=font_name, boldItalic=font_name)
    page_w, page_h = A4
    c = canvas.Canvas(str(PDF), pagesize=A4)
    c.setTitle("Project Meridian IT Separation Evidence Pack - Synthetic")
    c.setAuthor("Simulated worked example")
    # Header and scenario line.
    c.setFillColor(colors.HexColor("#1F3864")); c.rect(0, page_h - 56, page_w, 56, stroke=0, fill=1)
    c.setFillColor(colors.white); c.setFont(font_name, 16); c.drawString(32, page_h - 28, "Project Meridian | IT Separation")
    c.setFont(font_name, 9); c.drawRightString(page_w - 32, page_h - 28, "SYNTHETIC")
    draw_paragraph(c, "Simulated divestiture of a specialty-components division to a simulated PE-backed buyer; this is a worked example, not a client deliverable or past work.", 32, page_h - 68, page_w - 64, size=8.5, font_name=font_name)
    # Six summary figures read from Dashboard.
    card_y = page_h - 125
    card_w = (page_w - 64 - 5 * 8) / 6
    metrics = [("Apps", f"{int(app_count)}"), ("Day 1 plan gaps", f"{int(critical_gaps)}"), ("Day 1 ready", f"{readiness:.0%}"), ("One-off separation cost", money(one_off_cost)), ("One-time cost (one-off + TSA)", money(one_time_total)), ("Peak TSA / mo", money(peak))]
    for idx, (label, value) in enumerate(metrics):
        x = 32 + idx * (card_w + 8)
        c.setFillColor(colors.HexColor("#EEF3F8")); c.roundRect(x, card_y - 37, card_w, 43, 4, stroke=0, fill=1)
        c.setFillColor(colors.HexColor("#1F3864")); c.setFont(font_name, 7); c.drawString(x + 5, card_y - 8, label)
        c.setFillColor(colors.HexColor("#243247")); c.setFont(font_name, 9 if idx not in (3, 4, 5) else 7.2); c.drawString(x + 5, card_y - 25, value)
    # Left: disposition and Day 1 readiness.
    left_x, left_w = 32, 253
    y = card_y - 55
    c.setFillColor(colors.HexColor("#1F3864")); c.setFont(font_name, 10); c.drawString(left_x, y, "Disposition mix")
    data = [["Disposition", "Apps", "One-off GBP"]]
    for rr in range(69, 75):
        data.append([disp_s.cell(rr, 1).value, f"{int(disp_s.cell(rr, 2).value or 0)}", money(disp_s.cell(rr, 3).value)])
    y = draw_table(c, data, left_x, y - 6, [114, 38, 101], font_name=font_name) - 16
    c.setFillColor(colors.HexColor("#1F3864")); c.setFont(font_name, 10); c.drawString(left_x, y, "Day 1 readiness by workstream")
    data = [["Workstream", "Ready", "Open No-go"]]
    for rr in range(6, 6 + len(workstream_rows)):
        data.append([day_data.cell(rr, 11).value, f"{day_data.cell(rr, 12).value:.0%}", str(int(day_data.cell(rr, 13).value or 0))])
    y = draw_table(c, data, left_x, y - 6, [145, 47, 61], font_name=font_name)
    # Right: TSA run-off line and cost reference bridge table.
    rx, rw = 310, page_w - 342
    ry = card_y - 55
    c.setFillColor(colors.HexColor("#1F3864")); c.setFont(font_name, 10); c.drawString(rx, ry, "Buyer TSA run-off | GBP / month")
    graph_x, graph_y, graph_w, graph_h = rx, ry - 103, rw, 82
    vals = [float(tsa_data.cell(45 + m, 2).value or 0) for m in range(24)]
    max_v = max(vals) or 1
    c.setStrokeColor(colors.HexColor("#CBD5E1")); c.setLineWidth(0.5)
    for k in range(3):
        yy = graph_y + k * graph_h / 2
        c.line(graph_x, yy, graph_x + graph_w, yy)
    c.setStrokeColor(colors.HexColor("#2E5597")); c.setLineWidth(1.7)
    points = []
    for m, v in enumerate(vals):
        px = graph_x + (m / 23) * graph_w
        py = graph_y + (v / max_v) * graph_h
        points.append((px, py))
    path = c.beginPath(); path.moveTo(*points[0])
    for point in points[1:]: path.lineTo(*point)
    c.drawPath(path, stroke=1, fill=0)
    c.setFont(font_name, 6.5); c.setFillColor(colors.HexColor("#64748B")); c.drawString(graph_x, graph_y - 11, "Apr 2027")
    c.drawRightString(graph_x + graph_w, graph_y - 11, "Mar 2029")
    ry = graph_y - 30
    c.setFillColor(colors.HexColor("#1F3864")); c.setFont(font_name, 9); c.drawString(rx, ry, "Bridge A | Steady-state annual IT run-rate (synthetic)")
    data = [["Component", "GBP / year"]]
    for rr in range(18, 23):
        label = "Identified synergies, annual midpoint" if rr == 21 else cost_data.cell(rr, 8).value
        data.append([label, money(cost_data.cell(rr, 9).value)])
    table_bottom = draw_table(c, data, rx, ry - 5, [rw - 100, 100], font_name=font_name)
    c.setFillColor(colors.HexColor("#1F3864")); c.setFont(font_name, 9); c.drawString(rx, table_bottom - 12, "Bridge B | One-time cost to separate, Day 1 to TSA exit (synthetic)")
    data = [["Component", "GBP"]]
    for rr in range(26, 32):
        data.append([cost_data.cell(rr, 8).value, money(cost_data.cell(rr, 9).value)])
    table_bottom = draw_table(c, data, rx, table_bottom - 17, [rw - 88, 88], font_name=font_name)
    draw_paragraph(c, "TSA fees replace standalone run cost for covered services during the TSA period; they are not added to the run-rate.", rx, table_bottom - 5, rw, size=6.2, color="#64748B", font_name=font_name)
    # Three risk bullets, assumptions, and what this is/isn't.
    bottom_top = 170
    c.setStrokeColor(colors.HexColor("#CBD5E1")); c.line(32, bottom_top + 9, page_w - 32, bottom_top + 9)
    col_w = (page_w - 80) / 3
    col_x = [32, 32 + col_w + 8, 32 + 2 * (col_w + 8)]
    c.setFillColor(colors.HexColor("#1F3864")); c.setFont(font_name, 9); c.drawString(col_x[0], bottom_top - 5, "Three key risks")
    ty = bottom_top - 14
    for risk in risks[:3]:
        ty = draw_paragraph(c, "- " + risk["risk"], col_x[0], ty, col_w, size=6.5, leading=7.4, font_name=font_name) - 3
    c.setFillColor(colors.HexColor("#1F3864")); c.setFont(font_name, 9); c.drawString(col_x[1], bottom_top - 5, "Three assumptions")
    assumption_items = [
        f"TSA mark-up: {as_num(wb['README_Assumptions']['B14'].value):.0%}; extension step-up: {as_num(wb['README_Assumptions']['B15'].value):.0%}.",
        f"Applications standalone run cost: source total plus {as_num(wb['README_Assumptions']['B16'].value):.0%} pricing loss.",
        f"Day 1: {day_one.isoformat()}; tracker as-of month: {as_int(wb['README_Assumptions']['B28'].value)}.",
    ]
    ty = bottom_top - 14
    for item in assumption_items:
        ty = draw_paragraph(c, "- " + item, col_x[1], ty, col_w, size=6.8, leading=8, font_name=font_name) - 5
    c.setFillColor(colors.HexColor("#1F3864")); c.setFont(font_name, 9); c.drawString(col_x[2], bottom_top - 5, "What this is / is not")
    draw_paragraph(c, "Synthetic planning model with formula-linked assumptions, not evidence of an actual transaction or achieved savings. TSA reverse lines and outsourcing alternatives are excluded from buyer TSA cost and synergy totals.", col_x[2], bottom_top - 14, col_w, size=6.8, leading=8, font_name=font_name)
    c.setFillColor(colors.HexColor("#64748B")); c.setFont(font_name, 7); c.drawString(32, 24, "Simulated worked example | synthetic data | GBP | Source values and assumptions from workbook")
    c.drawRightString(page_w - 32, 24, "1 / 1")
    c.showPage(); c.save()
    reader = PdfReader(str(PDF))
    if len(reader.pages) != 1:
        raise AssertionError(f"One-page PDF expected; found {len(reader.pages)} pages")
    pdf_text = "\n".join(page.extract_text() or "" for page in reader.pages)
    if "SYNTHETIC" not in pdf_text or "Project Meridian" not in pdf_text:
        raise AssertionError("PDF is missing its title or synthetic label")
    if "One-off separation cost" not in pdf_text or "One-time cost to separate" not in pdf_text:
        raise AssertionError("PDF must show one-off separation cost and one-time cost to separate separately")


build_pdf()
print("PASS | 13 Checks tab results (9 core + 4 PMI); independent source-to-workbook reconciliations and both bridges passed; one-page PDF generated from workbook values.")
