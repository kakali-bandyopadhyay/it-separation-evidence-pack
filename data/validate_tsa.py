import csv
from datetime import date
from dateutil.relativedelta import relativedelta

apps_path = r"C:\Projects\it-separation-evidence-pack\data\apps.csv"
tsa_path = r"C:\Users\KBand\OneDrive\Documents\deepseek-harness\default-workspace\tsa_services.csv"

with open(apps_path, newline="", encoding="utf-8") as f:
    apps = list(csv.DictReader(f))
with open(tsa_path, newline="", encoding="utf-8") as f:
    tsa = list(csv.DictReader(f))

cols = ["svc_id","service","direction","linked_app_ids","start_date","planned_end_date","cost_basis_gbp_month","exit_dependency","exit_owner_role","extension_option_months"]
print("columns:", list(tsa[0].keys()) == cols)
print("rows:", len(tsa))
print("unique ids:", len({r["svc_id"] for r in tsa}))

app_by_id = {a["app_id"]: a for a in apps}
svc_ids = {r["svc_id"] for r in tsa}
reverse = [r["svc_id"] for r in tsa if r["direction"] == "Buyer to seller"]
print("reverse count:", len(reverse), reverse)
print("bad directions:", [r["svc_id"] for r in tsa if r["direction"] not in {"Seller to buyer","Buyer to seller"}])

bad_links = []
for r in tsa:
    for aid in r["linked_app_ids"].split(";"):
        if aid not in app_by_id:
            bad_links.append((r["svc_id"], aid))
print("bad links:", bad_links)

linked = set()
for r in tsa:
    linked.update(r["linked_app_ids"].split(";"))

mandatory = {a["app_id"] for a in apps if a["shared_with_seller"] == "Y" or (a["wave"] == "Wave 2" and a["day1_criticality"] == "1")}
print("mandatory count:", len(mandatory))
print("missing mandatory:", sorted(mandatory - linked))
print("all wave2:", sum(1 for a in apps if a["wave"] == "Wave 2"), "covered:", sum(1 for a in apps if a["wave"] == "Wave 2" and a["app_id"] in linked))
print("all wave3:", sum(1 for a in apps if a["wave"] == "Wave 3"), "covered:", sum(1 for a in apps if a["wave"] == "Wave 3" and a["app_id"] in linked))
day1_non_shared = [a["app_id"] for a in apps if a["wave"] == "Day 1" and a["shared_with_seller"] != "Y" and a["app_id"] in linked]
print("day1 non-shared covered:", len(day1_non_shared), sorted(day1_non_shared))

bad_dates = []
duration_issues = []
for r in tsa:
    start = date.fromisoformat(r["start_date"])
    end = date.fromisoformat(r["planned_end_date"])
    delta = relativedelta(end + relativedelta(days=1), start)
    months = delta.years * 12 + delta.months
    if start != date(2027,4,1) or not (3 <= months <= 18) or end.day != (end + relativedelta(days=1) - relativedelta(days=1)).day:
        bad_dates.append((r["svc_id"], months, r["start_date"], r["planned_end_date"]))
    for aid in r["linked_app_ids"].split(";"):
        wave = app_by_id[aid]["wave"]
        if wave == "Wave 2" and not (6 <= months <= 12):
            duration_issues.append((r["svc_id"], aid, "Wave 2", months))
        if wave == "Wave 3" and not (9 <= months <= 18):
            duration_issues.append((r["svc_id"], aid, "Wave 3", months))
print("bad dates:", bad_dates)
print("duration issues:", duration_issues)

bad_deps = []
for r in tsa:
    dep = r["exit_dependency"]
    if dep == r["svc_id"] or (dep not in app_by_id and dep not in svc_ids):
        bad_deps.append((r["svc_id"], dep))
print("bad deps:", bad_deps)

total = sum(int(r["cost_basis_gbp_month"]) for r in tsa)
print("total monthly cost:", total)
print("min/max cost:", min(int(r["cost_basis_gbp_month"]) for r in tsa), max(int(r["cost_basis_gbp_month"]) for r in tsa))
print("empty fields:", [(r["svc_id"], k) for r in tsa for k, v in r.items() if v is None or str(v).strip() == ""])
print("svc IDs:", [r["svc_id"] for r in tsa])
