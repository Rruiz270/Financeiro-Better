#!/usr/bin/env python3
"""
Generate receitas_customers.js by:
1. Parsing receitas_all_data.js (Vindi bills)
2. Grouping by customer name
3. Cross-referencing with the sales spreadsheet
"""

import json
import re
import unicodedata
from datetime import datetime

import openpyxl

# ── 1. Parse Vindi data ────────────────────────────────────────────
with open("/Users/Raphael/Financeiro-Better/receitas_all_data.js", "r") as f:
    raw = f.read()

m = re.search(r"var data = (\[.*\]);?\s*$", raw, re.DOTALL)
if not m:
    raise ValueError("Could not parse receitas_all_data.js")

bills = json.loads(m.group(1))
print(f"Loaded {len(bills)} Vindi bills")

# ── 2. Group by customer ───────────────────────────────────────────
customers = {}
for b in bills:
    c = b["cliente"]
    if c not in customers:
        customers[c] = {
            "bills": [],
            "total": 0.0,
            "first": b["data"],
            "last": b["data"],
        }
    rec = customers[c]
    rec["bills"].append(
        {
            "data": b["data"],
            "valor": b["valor"],
            "situacao": b["situacao"],
            "mes": b["mes"],
            "categoria": b["categoria"],
        }
    )
    rec["total"] += b["valor"]
    if b["data"] < rec["first"]:
        rec["first"] = b["data"]
    if b["data"] > rec["last"]:
        rec["last"] = b["data"]

print(f"Unique Vindi customers: {len(customers)}")

# ── 3. Read sales spreadsheet ──────────────────────────────────────
wb = openpyxl.load_workbook(
    "/Users/Raphael/Downloads/Vendas para Emissão de NF - ATUALIZADA_03.xlsx",
    data_only=True,
)
ws = wb["Vendas"]

# Normalize helper
def norm(s):
    if not s:
        return ""
    s = str(s).strip().upper()
    # Remove accents
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s


# Build lookup from spreadsheet: by normalized name
sales_by_name = {}
sales_by_email = {}
all_sales = []

for row_num in range(2, ws.max_row + 1):
    name = ws.cell(row=row_num, column=6).value
    if not name or not str(name).strip():
        continue
    name = str(name).strip()
    email = str(ws.cell(row=row_num, column=7).value or "").strip()
    cpf = str(ws.cell(row=row_num, column=5).value or "").strip()
    dt_raw = ws.cell(row=row_num, column=10).value
    produto = str(ws.cell(row=row_num, column=14).value or "").strip()
    renovacao = str(ws.cell(row=row_num, column=16).value or "").strip()
    valor_total = ws.cell(row=row_num, column=22).value or 0
    cancel_raw = ws.cell(row=row_num, column=27).value
    dt_cancel_raw = ws.cell(row=row_num, column=32).value

    # Parse date
    dt_str = ""
    if isinstance(dt_raw, datetime):
        dt_str = dt_raw.strftime("%Y-%m-%d")
    elif dt_raw:
        dt_str = str(dt_raw)

    # Parse cancellation
    cancelamento = False
    if cancel_raw:
        cv = str(cancel_raw).strip().lower()
        cancelamento = cv in ("true", "sim", "1", "cancelado")

    dt_cancel = ""
    if isinstance(dt_cancel_raw, datetime):
        dt_cancel = dt_cancel_raw.strftime("%Y-%m-%d")
    elif dt_cancel_raw:
        dt_cancel = str(dt_cancel_raw)

    sale = {
        "nome": name,
        "email": email,
        "cpf": cpf,
        "data_transacao": dt_str,
        "produto": produto,
        "renovacao": renovacao,
        "valor_total": float(valor_total) if valor_total else 0,
        "cancelamento": cancelamento,
        "data_cancelamento": dt_cancel,
    }
    all_sales.append(sale)

    nkey = norm(name)
    if nkey not in sales_by_name:
        sales_by_name[nkey] = []
    sales_by_name[nkey].append(sale)

    if email:
        ekey = email.lower()
        if ekey not in sales_by_email:
            sales_by_email[ekey] = []
        sales_by_email[ekey].append(sale)

print(f"Loaded {len(all_sales)} sales from spreadsheet")
print(f"Unique names in spreadsheet: {len(sales_by_name)}")

# ── 4. Cross-reference ─────────────────────────────────────────────
customer_data = {}
match_count = 0

for cliente, cdata in customers.items():
    entry = {
        "nome": cliente,
        "bills": len(cdata["bills"]),
        "total": round(cdata["total"], 2),
        "first": cdata["first"],
        "last": cdata["last"],
        "billsList": cdata["bills"],
        "produto": "",
        "renovacao": False,
        "cancelamento": False,
        "cpf": "",
        "email": "",
    }

    # Try matching by normalized name
    nkey = norm(cliente)
    matched_sales = sales_by_name.get(nkey, [])

    # If no match by name, try by email (in case cliente is an email)
    if not matched_sales and "@" in cliente:
        matched_sales = sales_by_email.get(cliente.lower(), [])

    if matched_sales:
        match_count += 1
        # Use first sale for product info; check if any is renewal
        s = matched_sales[0]
        entry["produto"] = s["produto"]
        entry["cpf"] = s["cpf"]
        entry["email"] = s["email"]
        entry["cancelamento"] = any(ms["cancelamento"] for ms in matched_sales)
        entry["renovacao"] = any(
            ms["renovacao"] in ("Renovação", "Renovacao") for ms in matched_sales
        )
        # If multiple sales for same person, mark as renewal
        if len(matched_sales) > 1:
            entry["renovacao"] = True
            # Collect all products
            prods = list(set(ms["produto"] for ms in matched_sales if ms["produto"]))
            entry["produto"] = ", ".join(prods) if prods else entry["produto"]

    customer_data[cliente] = entry

print(f"Matched {match_count} Vindi customers to spreadsheet")

# ── 5. Compute renewal stats ───────────────────────────────────────
total_customers = len(customer_data)
multi_purchase = sum(1 for c in customer_data.values() if c["bills"] >= 2)

# From spreadsheet
renewals_from_sheet = sum(
    1 for s in all_sales if s["renovacao"] in ("Renovação", "Renovacao")
)
total_renewal_value = sum(
    s["valor_total"]
    for s in all_sales
    if s["renovacao"] in ("Renovação", "Renovacao")
)

renewal_stats = {
    "total_customers": total_customers,
    "multi_purchase": multi_purchase,
    "renewals_from_spreadsheet": renewals_from_sheet,
    "total_renewal_value": round(total_renewal_value, 2),
}

print(f"Stats: {json.dumps(renewal_stats, indent=2)}")

# ── 6. Write output JS ─────────────────────────────────────────────
# Remove billsList from the output (too large), keep it in a separate var
customer_output = {}
bills_output = {}
for k, v in customer_data.items():
    bills_list = v.pop("billsList")
    customer_output[k] = v
    bills_output[k] = bills_list

out = "var customerData = " + json.dumps(customer_output, ensure_ascii=False, indent=1) + ";\n"
out += "var customerBills = " + json.dumps(bills_output, ensure_ascii=False, indent=1) + ";\n"
out += "var renewalStats = " + json.dumps(renewal_stats, ensure_ascii=False, indent=1) + ";\n"

outpath = "/Users/Raphael/Financeiro-Better/receitas_customers.js"
with open(outpath, "w") as f:
    f.write(out)

import os
size = os.path.getsize(outpath)
print(f"\nWrote {outpath} ({size:,} bytes)")
