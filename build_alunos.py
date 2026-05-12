#!/usr/bin/env python3
"""
Build comprehensive student database from:
1. Sales spreadsheet (Vendas tab - 229 records)
2. Vindi customer data (receitas_customers.js - 2852 entries)

Outputs:
- alunos_data.js (for HTML dashboard)
- Alunos_Base_Completa.xlsx (multi-tab Excel)
"""

import json
import re
import os
from datetime import datetime, date
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TODAY = date(2026, 5, 12)
PROJECT = "/Users/Raphael/Financeiro-Better"
SALES_FILE = "/Users/Raphael/Downloads/Vendas para Emissão de NF - ATUALIZADA_03.xlsx"

# ─── Helpers ──────────────────────────────────────────────────────────────────

def normalize_cpf(raw):
    """Strip dots, dashes, slashes, spaces. Return digits only or empty string."""
    if raw is None:
        return ""
    s = re.sub(r"[.\-/ ]", "", str(raw).strip())
    # Remove leading zeros padding issues
    return s if s.isdigit() else ""


def safe_date(val):
    """Convert to date object from datetime, string, or return None."""
    if val is None or val == "" or val == "-" or val == "nan":
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        val = val.strip()
        if val in ("-", "", "nan", "None"):
            return None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%dT%H:%M:%S"):
            try:
                return datetime.strptime(val, fmt).date()
            except ValueError:
                continue
    return None


def safe_float(val):
    if val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def safe_str(val):
    if val is None:
        return ""
    return str(val).strip()


def title_name(name):
    """Title-case a name but respect common exceptions."""
    if not name:
        return ""
    parts = name.strip().split()
    result = []
    lower_words = {"de", "da", "do", "das", "dos", "e", "em"}
    for i, p in enumerate(parts):
        if i > 0 and p.lower() in lower_words:
            result.append(p.lower())
        else:
            result.append(p.capitalize())
    return " ".join(result)


def classify_modalidade(produto, tipo_doc):
    """Classify the modality based on product name and document type."""
    if not produto:
        if tipo_doc == "CNPJ":
            return "In-Company"
        return "Community"
    p = produto.lower()
    if "flow" in p:
        return "Community Flow"
    if "espanhol" in p or "spanish" in p:
        return "Espanhol"
    if "particular" in p or "private" in p:
        return "Private"
    if tipo_doc == "CNPJ":
        return "In-Company"
    # Default for "Inglês 12 meses", "Inglês 06 meses", "Inglês 09 meses"
    return "Community"


def determine_status(cancelou, data_cancel, data_fim, vindi_last):
    """Determine student status.

    Active if:
    - Course end date is in the future, OR
    - Last Vindi payment is within ~3 months of today (Mar 2026+), OR
    - data_fim is None but Vindi payments are recent

    Cancelled if cancelou flag is set.
    Expirado if course ended and no recent Vindi payments.
    """
    if cancelou:
        return "Cancelado"

    # Check Vindi activity: if last payment is recent, student is active
    vl = safe_date(vindi_last) if vindi_last else None
    vindi_active = vl and vl >= date(2026, 3, 1)  # Last 3 months

    if data_fim:
        if data_fim >= TODAY:
            return "Ativo"
        if vindi_active:
            return "Ativo"
        return "Expirado"

    # No end date
    if vindi_active:
        return "Ativo"
    return "Expirado"


def estimate_remaining_classes(data_inicio, duracao_meses, status):
    """Estimate remaining classes based on course duration and elapsed time.
    Community classes = ~2 per week = ~8/month.
    Private = ~4/month.
    """
    if status != "Ativo" or not data_inicio or not duracao_meses:
        return 0
    di = safe_date(data_inicio) if isinstance(data_inicio, str) else data_inicio
    if not di:
        return 0
    dur = int(duracao_meses) if duracao_meses else 12
    total_classes = dur * 8  # ~2/week
    elapsed_months = (TODAY.year - di.year) * 12 + (TODAY.month - di.month)
    used = elapsed_months * 8
    remaining = max(0, total_classes - used)
    return remaining


# ─── Step 1: Read Sales Spreadsheet ──────────────────────────────────────────

print("Reading sales spreadsheet...")
wb = openpyxl.load_workbook(SALES_FILE, data_only=True)
ws = wb["Vendas"]

sales_by_cpf = defaultdict(list)
sales_by_email = defaultdict(list)
sales_records = []

for r in range(2, ws.max_row + 1):
    nome = safe_str(ws.cell(r, 6).value)
    if not nome:
        continue

    cpf_raw = safe_str(ws.cell(r, 5).value)
    cpf = normalize_cpf(cpf_raw)
    email = safe_str(ws.cell(r, 7).value).lower()
    celular = safe_str(ws.cell(r, 8).value)
    endereco = safe_str(ws.cell(r, 9).value)
    data_transacao = safe_date(ws.cell(r, 10).value)
    data_venda = safe_date(ws.cell(r, 11).value)
    ultima_parcela = safe_date(ws.cell(r, 12).value)
    forma = safe_str(ws.cell(r, 13).value)
    produto = safe_str(ws.cell(r, 14).value)
    fonte = safe_str(ws.cell(r, 15).value)
    renovacao = safe_str(ws.cell(r, 16).value)
    nivel = safe_str(ws.cell(r, 17).value)
    desconto = safe_float(ws.cell(r, 18).value)
    duracao = safe_str(ws.cell(r, 19).value)
    valor_total = safe_float(ws.cell(r, 22).value)
    cancel_raw = ws.cell(r, 27).value
    cancelamento = bool(cancel_raw) and str(cancel_raw).strip() not in ("-", "", "False", "0")
    data_cancelamento = safe_date(ws.cell(r, 32).value)

    # Determine doc type from CPF length
    tipo_doc_val = safe_str(ws.cell(r, 4).value)
    if len(cpf) == 14 or len(cpf) > 11:
        tipo_doc = "CNPJ"
    elif len(cpf) == 11:
        tipo_doc = "CPF"
    elif "cnpj" in tipo_doc_val.lower():
        tipo_doc = "CNPJ"
    else:
        tipo_doc = "CPF"

    # Calculate end date: data_venda + duracao_curso months
    # Note: ultima_parcela (col 12) is the last PAYMENT date, not course end
    # The actual course end date = data_venda + duracao_curso
    data_fim = None
    if data_venda and duracao:
        try:
            dur_int = int(duracao)
            year = data_venda.year + (data_venda.month + dur_int - 1) // 12
            month = (data_venda.month + dur_int - 1) % 12 + 1
            day = min(data_venda.day, 28)  # safe day
            data_fim = date(year, month, day)
        except Exception:
            data_fim = ultima_parcela  # fallback
    if not data_fim:
        data_fim = ultima_parcela  # fallback to last installment

    rec = {
        "cpf": cpf,
        "nome": title_name(nome),
        "email": email,
        "celular": celular,
        "endereco": endereco,
        "tipo_doc": tipo_doc,
        "data_transacao": data_transacao,
        "data_venda": data_venda,
        "data_fim": data_fim,
        "forma": forma,
        "produto": produto,
        "fonte": fonte,
        "renovacao": renovacao,
        "nivel": nivel,
        "desconto": desconto,
        "duracao": duracao,
        "valor_total": valor_total,
        "cancelamento": cancelamento,
        "data_cancelamento": data_cancelamento,
    }
    sales_records.append(rec)

    # Index by CPF and email
    if cpf:
        sales_by_cpf[cpf].append(rec)
    if email:
        sales_by_email[email].append(rec)

print(f"  Sales records loaded: {len(sales_records)}")

# ─── Step 2: Read Vindi Customer Data ────────────────────────────────────────

print("Reading Vindi customer data...")
with open(os.path.join(PROJECT, "receitas_customers.js"), "r") as f:
    content = f.read()

# Parse customerData
cd_start = content.index("var customerData = ") + len("var customerData = ")
cd_end = content.index(";\nvar customerBills")
customer_data = json.loads(content[cd_start:cd_end])

# Parse customerBills
cb_start = content.index("var customerBills = ") + len("var customerBills = ")
cb_end = content.index(";\nvar renewalStats")
customer_bills = json.loads(content[cb_start:cb_end])

print(f"  Vindi customers loaded: {len(customer_data)}")

# Build Vindi lookup by CPF and email
vindi_by_cpf = {}
vindi_by_email = {}
vindi_by_name_lower = {}

for name, cust in customer_data.items():
    cpf = normalize_cpf(cust.get("cpf", ""))
    email = (cust.get("email", "") or "").lower().strip()
    bills = customer_bills.get(name, [])

    vindi_rec = {
        "nome_vindi": name,
        "bills": cust.get("bills", 0),
        "total": cust.get("total", 0.0),
        "first": cust.get("first", ""),
        "last": cust.get("last", ""),
        "produto": cust.get("produto", ""),
        "renovacao": cust.get("renovacao", False),
        "cancelamento": cust.get("cancelamento", False),
        "bill_detail": bills,
    }

    if cpf:
        vindi_by_cpf[cpf] = vindi_rec
    if email:
        vindi_by_email[email] = vindi_rec
    vindi_by_name_lower[name.lower().strip()] = vindi_rec

# ─── Step 3: Consolidate into unified student records ─────────────────────────

print("Consolidating student database...")
students = {}  # keyed by CPF or email

# First pass: sales spreadsheet (primary source, 229 records)
for rec in sales_records:
    key = rec["cpf"] if rec["cpf"] else rec["email"]
    if not key:
        key = rec["nome"].lower().replace(" ", "_")

    if key in students:
        student = students[key]
        # Add this sale
        student["vendas"].append({
            "data_venda": str(rec["data_venda"]) if rec["data_venda"] else "",
            "produto": rec["produto"],
            "nivel": rec["nivel"],
            "valor": rec["valor_total"],
            "renovacao": rec["renovacao"],
            "vendedor": rec["fonte"],
            "cancelamento": rec["cancelamento"],
            "data_cancel": str(rec["data_cancelamento"]) if rec["data_cancelamento"] else "",
            "forma": rec["forma"],
            "desconto": rec["desconto"],
            "duracao": rec["duracao"],
        })
        # Update aggregates
        student["total_gasto"] += rec["valor_total"]
        if rec["renovacao"] == "Renovação":
            student["renovacoes"] += 1
        if rec["cancelamento"]:
            student["cancelou"] = True
            if rec["data_cancelamento"]:
                student["data_cancelamento"] = str(rec["data_cancelamento"])
        # Update date range
        if rec["data_venda"]:
            if not student["data_inicio"] or rec["data_venda"] < safe_date(student["data_inicio"]):
                student["data_inicio"] = str(rec["data_venda"])
        if rec["data_fim"]:
            if not student["data_fim"] or rec["data_fim"] > safe_date(student["data_fim"]):
                student["data_fim"] = str(rec["data_fim"])
        # Prefer non-empty fields
        if not student["email"] and rec["email"]:
            student["email"] = rec["email"]
        if not student["celular"] and rec["celular"]:
            student["celular"] = rec["celular"]
        if not student["endereco"] and rec["endereco"]:
            student["endereco"] = rec["endereco"]
        # Keep latest nivel
        if rec["nivel"]:
            student["nivel"] = rec["nivel"]
        if rec["produto"]:
            student["produto_principal"] = rec["produto"]
    else:
        students[key] = {
            "cpf": rec["cpf"],
            "nome": rec["nome"],
            "email": rec["email"],
            "celular": rec["celular"],
            "endereco": rec["endereco"],
            "tipo_doc": rec["tipo_doc"],
            "vendas": [{
                "data_venda": str(rec["data_venda"]) if rec["data_venda"] else "",
                "produto": rec["produto"],
                "nivel": rec["nivel"],
                "valor": rec["valor_total"],
                "renovacao": rec["renovacao"],
                "vendedor": rec["fonte"],
                "cancelamento": rec["cancelamento"],
                "data_cancel": str(rec["data_cancelamento"]) if rec["data_cancelamento"] else "",
                "forma": rec["forma"],
                "desconto": rec["desconto"],
                "duracao": rec["duracao"],
            }],
            "data_inicio": str(rec["data_venda"]) if rec["data_venda"] else "",
            "data_fim": str(rec["data_fim"]) if rec["data_fim"] else "",
            "cancelou": rec["cancelamento"],
            "data_cancelamento": str(rec["data_cancelamento"]) if rec["data_cancelamento"] else "",
            "renovacoes": 1 if rec["renovacao"] == "Renovação" else 0,
            "total_gasto": rec["valor_total"],
            "produto_principal": rec["produto"],
            "nivel": rec["nivel"],
            "duracao": rec["duracao"],
            # Vindi placeholders
            "vindi_bills": 0,
            "vindi_total": 0.0,
            "vindi_last_payment": "",
            "vindi_bill_detail": [],
            # Source tracking
            "fonte_planilha": True,
            "fonte_vindi": False,
        }

print(f"  Unique students from sales: {len(students)}")

# Second pass: match Vindi data to existing students
matched_vindi = set()
for key, student in students.items():
    cpf = student["cpf"]
    email = student["email"]
    nome_lower = student["nome"].lower().strip()

    vindi = None
    matched_name = None

    # Try CPF match first
    if cpf and cpf in vindi_by_cpf:
        vindi = vindi_by_cpf[cpf]
        matched_name = vindi["nome_vindi"]
    # Then email match
    elif email and email in vindi_by_email:
        vindi = vindi_by_email[email]
        matched_name = vindi["nome_vindi"]
    # Then name match (fuzzy)
    else:
        # Try exact name match
        for vname_lower, vrec in vindi_by_name_lower.items():
            if nome_lower == vname_lower:
                vindi = vrec
                matched_name = vrec["nome_vindi"]
                break

    if vindi:
        student["vindi_bills"] = vindi["bills"]
        student["vindi_total"] = vindi["total"]
        student["vindi_last_payment"] = vindi["last"]
        student["vindi_bill_detail"] = vindi["bill_detail"]
        student["fonte_vindi"] = True
        if matched_name:
            matched_vindi.add(matched_name)

print(f"  Matched Vindi records: {len(matched_vindi)}")

# Third pass: add Vindi-only customers (not in sales spreadsheet)
vindi_only = 0
for name, cust in customer_data.items():
    if name in matched_vindi:
        continue
    cpf = normalize_cpf(cust.get("cpf", ""))
    email = (cust.get("email", "") or "").lower().strip()

    # Check if already matched by CPF or email
    already = False
    if cpf and cpf in {s["cpf"] for s in students.values()}:
        already = True
    if not already and email and email in {s["email"] for s in students.values() if s["email"]}:
        already = True

    if already:
        continue

    key = cpf if cpf else (email if email else name.lower().replace(" ", "_"))
    bills = customer_bills.get(name, [])

    # Determine tipo_doc
    tipo_doc = "CNPJ" if cpf and len(cpf) > 11 else "CPF"

    students[key] = {
        "cpf": cpf,
        "nome": title_name(name),
        "email": email,
        "celular": "",
        "endereco": "",
        "tipo_doc": tipo_doc,
        "vendas": [],
        "data_inicio": cust.get("first", ""),
        "data_fim": cust.get("last", ""),
        "cancelou": cust.get("cancelamento", False),
        "data_cancelamento": "",
        "renovacoes": 1 if cust.get("renovacao", False) else 0,
        "total_gasto": 0.0,
        "produto_principal": cust.get("produto", ""),
        "nivel": "",
        "duracao": "",
        "vindi_bills": cust.get("bills", 0),
        "vindi_total": cust.get("total", 0.0),
        "vindi_last_payment": cust.get("last", ""),
        "vindi_bill_detail": bills,
        "fonte_planilha": False,
        "fonte_vindi": True,
    }
    vindi_only += 1

print(f"  Vindi-only customers added: {vindi_only}")
print(f"  Total consolidated students: {len(students)}")

# ─── Step 4: Classify and finalize ───────────────────────────────────────────

print("Classifying and finalizing records...")

# Separate planilha students (229 records) from vindi-only
planilha_students = {}
all_students = {}

for key, s in students.items():
    # Classify modalidade
    s["modalidade"] = classify_modalidade(s["produto_principal"], s["tipo_doc"])

    # Determine tipo_cliente
    s["tipo_cliente"] = "PJ" if s["tipo_doc"] == "CNPJ" else "PF"

    # Determine status
    data_fim = safe_date(s["data_fim"]) if s["data_fim"] else None
    s["contrato_ativo"] = False
    s["status"] = determine_status(
        s["cancelou"],
        s.get("data_cancelamento"),
        data_fim,
        s["vindi_last_payment"]
    )
    if s["status"] == "Ativo":
        s["contrato_ativo"] = True

    # Estimate remaining classes
    dur = 0
    try:
        dur = int(s["duracao"]) if s["duracao"] else 0
    except:
        dur = 12  # default
    s["aulas_remanescentes"] = estimate_remaining_classes(
        s["data_inicio"], dur if dur else 12, s["status"]
    )

    all_students[key] = s
    if s["fonte_planilha"]:
        planilha_students[key] = s

# ─── Compute stats (planilha-based: 229 original records) ────────────────────

# For the stats object we focus on the 229 planilha students
stats_source = planilha_students

total = len(stats_source)
ativos = sum(1 for s in stats_source.values() if s["status"] == "Ativo")
cancelados = sum(1 for s in stats_source.values() if s["status"] == "Cancelado")
concluidos = sum(1 for s in stats_source.values() if s["status"] == "Concluido")
expirados = sum(1 for s in stats_source.values() if s["status"] == "Expirado")
pf = sum(1 for s in stats_source.values() if s["tipo_cliente"] == "PF")
pj = sum(1 for s in stats_source.values() if s["tipo_cliente"] == "PJ")
community = sum(1 for s in stats_source.values() if s["modalidade"] == "Community")
community_flow = sum(1 for s in stats_source.values() if s["modalidade"] == "Community Flow")
espanhol = sum(1 for s in stats_source.values() if s["modalidade"] == "Espanhol")
private = sum(1 for s in stats_source.values() if s["modalidade"] == "Private")
in_company = sum(1 for s in stats_source.values() if s["modalidade"] == "In-Company")
renovacoes_total = sum(s["renovacoes"] for s in stats_source.values())
total_gasto = sum(s["total_gasto"] for s in stats_source.values())

# Also compute total including Vindi-only
total_all = len(all_students)
ativos_all = sum(1 for s in all_students.values() if s["status"] == "Ativo")
vindi_only_count = sum(1 for s in all_students.values() if not s["fonte_planilha"])

print(f"\n=== STATS (229 Planilha) ===")
print(f"  Total: {total}")
print(f"  Ativos: {ativos}, Cancelados: {cancelados}, Expirados: {expirados}")
print(f"  PF: {pf}, PJ: {pj}")
print(f"  Community: {community}, Flow: {community_flow}, Espanhol: {espanhol}, Private: {private}, InCo: {in_company}")
print(f"  Renovacoes: {renovacoes_total}")
print(f"  Total gasto: R$ {total_gasto:,.2f}")
print(f"\n=== ALL (incl Vindi-only) ===")
print(f"  Total: {total_all}, Vindi-only: {vindi_only_count}, Ativos all: {ativos_all}")

# ─── Step 5: Generate alunos_data.js ─────────────────────────────────────────

print("\nGenerating alunos_data.js...")


def serialize_student(s):
    """Convert student dict to JSON-safe dict."""
    return {
        "cpf": s["cpf"],
        "nome": s["nome"],
        "email": s["email"],
        "celular": s["celular"],
        "endereco": s["endereco"],
        "tipo_doc": s["tipo_doc"],
        "tipo_cliente": s["tipo_cliente"],
        "modalidade": s["modalidade"],
        "nivel": s["nivel"],
        "produto_principal": s["produto_principal"],
        "data_inicio": s["data_inicio"],
        "data_fim": s["data_fim"],
        "contrato_ativo": s["contrato_ativo"],
        "cancelou": s["cancelou"],
        "data_cancelamento": s["data_cancelamento"],
        "renovacoes": s["renovacoes"],
        "total_gasto": round(s["total_gasto"], 2),
        "status": s["status"],
        "aulas_remanescentes": s["aulas_remanescentes"],
        "vendas": s["vendas"],
        "vindi_bills": s["vindi_bills"],
        "vindi_total": round(s["vindi_total"], 2),
        "vindi_last_payment": s["vindi_last_payment"],
        "vindi_bill_detail": s["vindi_bill_detail"],
        "fonte_planilha": s["fonte_planilha"],
        "fonte_vindi": s["fonte_vindi"],
    }


# Sort all by name
sorted_students = sorted(all_students.values(), key=lambda s: s["nome"].lower())
alunos_array = [serialize_student(s) for s in sorted_students]

stats_obj = {
    "total": total,
    "total_all": total_all,
    "ativos": ativos,
    "ativos_all": ativos_all,
    "cancelados": cancelados,
    "concluidos": concluidos,
    "expirados": expirados,
    "pf": pf,
    "pj": pj,
    "community": community,
    "community_flow": community_flow,
    "espanhol": espanhol,
    "private": private,
    "in_company": in_company,
    "renovacoes": renovacoes_total,
    "total_gasto": round(total_gasto, 2),
    "vindi_only": vindi_only_count,
}

js_content = "var alunosData = " + json.dumps(alunos_array, ensure_ascii=False, indent=1) + ";\n\n"
js_content += "var alunosStats = " + json.dumps(stats_obj, ensure_ascii=False, indent=1) + ";\n"

with open(os.path.join(PROJECT, "alunos_data.js"), "w") as f:
    f.write(js_content)

print(f"  alunos_data.js written ({len(alunos_array)} students, {len(js_content)} chars)")

# ─── Step 6: Generate Excel ──────────────────────────────────────────────────

print("\nGenerating Alunos_Base_Completa.xlsx...")

# Style definitions
header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
header_fill_blue = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
header_fill_green = PatternFill(start_color="0D5E3A", end_color="0D5E3A", fill_type="solid")
header_fill_red = PatternFill(start_color="7A1F1F", end_color="7A1F1F", fill_type="solid")
header_fill_purple = PatternFill(start_color="4A1D8A", end_color="4A1D8A", fill_type="solid")
header_fill_cyan = PatternFill(start_color="0E7490", end_color="0E7490", fill_type="solid")
header_fill_amber = PatternFill(start_color="92400E", end_color="92400E", fill_type="solid")

fill_active = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
fill_cancel = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
fill_expired = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

thin_border = Border(
    left=Side(style="thin", color="D1D5DB"),
    right=Side(style="thin", color="D1D5DB"),
    top=Side(style="thin", color="D1D5DB"),
    bottom=Side(style="thin", color="D1D5DB"),
)

align_center = Alignment(horizontal="center", vertical="center")
align_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

# Use planilha students only for the Excel (229 core records)
planilha_list = sorted(planilha_students.values(), key=lambda s: s["nome"].lower())

XL_HEADERS = [
    "Nome", "CPF/CNPJ", "Tipo Doc", "Email", "Celular", "Endereco",
    "Modalidade", "Nivel", "Produto", "Status",
    "Data Inicio", "Data Fim", "Contrato Ativo",
    "Cancelou", "Data Cancelamento", "Renovacoes",
    "Total Gasto (R$)", "Vindi Bills", "Vindi Total (R$)", "Vindi Ultimo Pgto",
    "Tipo Cliente", "Aulas Remanescentes",
    "Vendedor", "Forma Pgto",
    "Fonte Planilha", "Fonte Vindi",
]


def write_student_row(ws, row, s, headers_map=None):
    """Write a student row to worksheet."""
    # Get first sale for vendedor/forma
    vendedor = ""
    forma = ""
    if s["vendas"]:
        vendedor = s["vendas"][-1].get("vendedor", "")
        forma = s["vendas"][-1].get("forma", "")

    values = [
        s["nome"],
        s["cpf"],
        s["tipo_doc"],
        s["email"],
        s["celular"],
        s["endereco"],
        s["modalidade"],
        s["nivel"],
        s["produto_principal"],
        s["status"],
        s["data_inicio"],
        s["data_fim"],
        "Sim" if s["contrato_ativo"] else "Nao",
        "Sim" if s["cancelou"] else "Nao",
        s["data_cancelamento"],
        s["renovacoes"],
        round(s["total_gasto"], 2),
        s["vindi_bills"],
        round(s["vindi_total"], 2),
        s["vindi_last_payment"],
        s["tipo_cliente"],
        s["aulas_remanescentes"],
        vendedor,
        forma,
        "Sim" if s["fonte_planilha"] else "Nao",
        "Sim" if s["fonte_vindi"] else "Nao",
    ]

    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.border = thin_border
        cell.alignment = Alignment(vertical="center")
        if col in (11, 12, 15, 20):  # date columns
            cell.alignment = align_center
        if col in (16, 17, 18, 19, 22):  # number columns
            cell.alignment = align_center
            if col in (17, 19):
                cell.number_format = '#,##0.00'

    # Color status cell
    status_cell = ws.cell(row=row, column=10)
    if s["status"] == "Ativo":
        status_cell.fill = fill_active
    elif s["status"] == "Cancelado":
        status_cell.fill = fill_cancel
    elif s["status"] == "Expirado":
        status_cell.fill = fill_expired


def write_headers(ws, headers, fill):
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = fill
        cell.alignment = align_center
        cell.border = thin_border


def auto_width(ws, headers):
    for col, h in enumerate(headers, 1):
        max_len = len(h) + 2
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=col).value
            if val:
                max_len = max(max_len, min(len(str(val)) + 2, 50))
        ws.column_dimensions[get_column_letter(col)].width = max_len


def add_filters(ws, headers):
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"


# Create workbook
xwb = openpyxl.Workbook()

# ─── Tab 1: Base Completa ────────────────────────────────────────────────────
ws1 = xwb.active
ws1.title = "Base Completa"
write_headers(ws1, XL_HEADERS, header_fill_blue)
for i, s in enumerate(planilha_list):
    write_student_row(ws1, i + 2, s)
auto_width(ws1, XL_HEADERS)
add_filters(ws1, XL_HEADERS)
ws1.freeze_panes = "A2"

# ─── Tab 2: Ativos ──────────────────────────────────────────────────────────
ws2 = xwb.create_sheet("Ativos")
ativos_list = [s for s in planilha_list if s["status"] == "Ativo"]
write_headers(ws2, XL_HEADERS, header_fill_green)
for i, s in enumerate(ativos_list):
    write_student_row(ws2, i + 2, s)
auto_width(ws2, XL_HEADERS)
add_filters(ws2, XL_HEADERS)
ws2.freeze_panes = "A2"

# ─── Tab 3: Cancelados ──────────────────────────────────────────────────────
ws3 = xwb.create_sheet("Cancelados")
cancel_list = [s for s in planilha_list if s["status"] == "Cancelado"]
write_headers(ws3, XL_HEADERS, header_fill_red)
for i, s in enumerate(cancel_list):
    write_student_row(ws3, i + 2, s)
auto_width(ws3, XL_HEADERS)
add_filters(ws3, XL_HEADERS)
ws3.freeze_panes = "A2"

# ─── Tab 4: Renovacoes ──────────────────────────────────────────────────────
ws4 = xwb.create_sheet("Renovacoes")
renov_list = [s for s in planilha_list if s["renovacoes"] > 0]
write_headers(ws4, XL_HEADERS, header_fill_purple)
for i, s in enumerate(renov_list):
    write_student_row(ws4, i + 2, s)
auto_width(ws4, XL_HEADERS)
add_filters(ws4, XL_HEADERS)
ws4.freeze_panes = "A2"

# ─── Tab 5: Para Importacao Portal ──────────────────────────────────────────
ws5 = xwb.create_sheet("Para Importacao Portal")
PORTAL_HEADERS = ["Nome", "Email", "CPF", "Celular", "Modalidade", "Nivel", "Status", "Aulas Remanescentes"]
write_headers(ws5, PORTAL_HEADERS, header_fill_cyan)
portal_list = [s for s in planilha_list if s["status"] == "Ativo"]
for i, s in enumerate(portal_list):
    row = i + 2
    values = [s["nome"], s["email"], s["cpf"], s["celular"], s["modalidade"], s["nivel"], s["status"], s["aulas_remanescentes"]]
    for col, val in enumerate(values, 1):
        cell = ws5.cell(row=row, column=col, value=val)
        cell.border = thin_border
        cell.alignment = Alignment(vertical="center")
auto_width(ws5, PORTAL_HEADERS)
add_filters(ws5, PORTAL_HEADERS)
ws5.freeze_panes = "A2"

# ─── Tab 6: Resumo ──────────────────────────────────────────────────────────
ws6 = xwb.create_sheet("Resumo")

summary_data = [
    ("Metrica", "Valor"),
    ("Total Alunos (Planilha)", total),
    ("Total Consolidado (incl. Vindi)", total_all),
    ("", ""),
    ("STATUS", ""),
    ("Ativos", ativos),
    ("Cancelados", cancelados),
    ("Expirados", expirados),
    ("", ""),
    ("TIPO CLIENTE", ""),
    ("Pessoa Fisica (PF)", pf),
    ("Pessoa Juridica (PJ)", pj),
    ("", ""),
    ("MODALIDADE", ""),
    ("Community", community),
    ("Community Flow", community_flow),
    ("Espanhol", espanhol),
    ("Private", private),
    ("In-Company", in_company),
    ("", ""),
    ("FINANCEIRO", ""),
    ("Total Gasto (Planilha)", f"R$ {total_gasto:,.2f}"),
    ("Renovacoes", renovacoes_total),
    ("", ""),
    ("VINDI", ""),
    ("Clientes apenas Vindi", vindi_only_count),
    ("Ativos (todos)", ativos_all),
    ("", ""),
    ("Gerado em", str(TODAY)),
]

for r, (metric, val) in enumerate(summary_data, 1):
    c1 = ws6.cell(row=r, column=1, value=metric)
    c2 = ws6.cell(row=r, column=2, value=val)
    c1.border = thin_border
    c2.border = thin_border
    if r == 1:
        c1.font = header_font
        c1.fill = header_fill_amber
        c2.font = header_font
        c2.fill = header_fill_amber
    elif metric in ("STATUS", "TIPO CLIENTE", "MODALIDADE", "FINANCEIRO", "VINDI"):
        c1.font = Font(bold=True)

ws6.column_dimensions["A"].width = 35
ws6.column_dimensions["B"].width = 25

# Save
excel_path = os.path.join(PROJECT, "Alunos_Base_Completa.xlsx")
xwb.save(excel_path)
print(f"  Excel saved to {excel_path}")
print(f"  Tabs: {xwb.sheetnames}")
print(f"  Base Completa: {len(planilha_list)} rows")
print(f"  Ativos: {len(ativos_list)} rows")
print(f"  Cancelados: {len(cancel_list)} rows")
print(f"  Renovacoes: {len(renov_list)} rows")
print(f"  Portal: {len(portal_list)} rows")

print("\n=== DONE ===")
