#!/usr/bin/env python3
"""
Fetch BMA Solucoes data from OMIE API and generate JS data files.
BMA is the HR/payroll BPO company for Better Education.
"""
import json
import urllib.request
import time
import os
from collections import defaultdict
from datetime import datetime

BMA_KEY = "4602985397010"
BMA_SECRET = "e7e7e8ebffe8c76051459f4dbbb468e5"

MESES = {
    '01': 'Jan', '02': 'Fev', '03': 'Mar', '04': 'Abr',
    '05': 'Mai', '06': 'Jun', '07': 'Jul', '08': 'Ago',
    '09': 'Set', '10': 'Out', '11': 'Nov', '12': 'Dez'
}

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))


def bma_omie_call(endpoint: str, method: str, params: dict) -> dict:
    """Make a single OMIE API call using BMA credentials."""
    body = json.dumps({
        "call": method,
        "app_key": BMA_KEY,
        "app_secret": BMA_SECRET,
        "param": [params]
    }).encode()
    req = urllib.request.Request(
        f"https://app.omie.com.br/api/v1/{endpoint}",
        data=body,
        headers={"Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read())


def bma_omie_paginate(endpoint: str, method: str, list_key: str,
                      per_page: int = 50, delay: int = 3) -> list:
    """Paginate through OMIE API results with delay between pages."""
    results = []
    page = 1
    total_pages = None
    while True:
        for attempt in range(3):
            try:
                d = bma_omie_call(endpoint, method, {
                    "pagina": page,
                    "registros_por_pagina": per_page,
                    "apenas_importado_api": "N"
                })
                break
            except Exception as e:
                print(f"  BMA Omie retry {attempt+1} page {page}: {e}")
                time.sleep(delay * (attempt + 1))
        else:
            print(f"  BMA Omie SKIP page {page}")
            page += 1
            if total_pages and page > total_pages:
                break
            continue

        if total_pages is None:
            total_pages = d.get('total_de_paginas', 0)
            total_records = d.get('total_de_registros', 0)
            print(f"  {method}: {total_records} records, {total_pages} pages")

        results.extend(d.get(list_key, []))
        if page % 10 == 0:
            print(f"    Page {page}/{total_pages}")
        if page >= total_pages:
            break
        page += 1
        time.sleep(delay)
    return results


def parse_br_date(dt: str):
    """Parse dd/mm/yyyy date string -> (year, mes_label like 'Jan/2026')."""
    if not dt or '/' not in dt:
        return None, None
    parts = dt.split('/')
    if len(parts) < 3:
        return None, None
    month_num = parts[1]
    year = parts[2]
    mes_label = MESES.get(month_num, '') + '/' + year
    return year, mes_label


def main():
    print(f"=== BMA Data Fetch - {datetime.now().strftime('%Y-%m-%d %H:%M')} ===\n")

    # 1. Fetch BMA clients for name resolution
    print("[1/2] Fetching BMA clients...")
    clients_raw = bma_omie_paginate(
        "geral/clientes/", "ListarClientes", "clientes_cadastro",
        per_page=50, delay=3
    )
    client_map = {}
    for c in clients_raw:
        cid = c.get("codigo_cliente_omie", 0)
        name = c.get("razao_social", c.get("nome_fantasia", f"Cliente {cid}"))
        client_map[cid] = name
    print(f"  {len(client_map)} clients loaded")

    # 2. Fetch BMA contas a pagar
    print("[2/2] Fetching BMA contas a pagar...")
    cp_raw = bma_omie_paginate(
        "financas/contapagar/", "ListarContasPagar", "conta_pagar_cadastro",
        per_page=50, delay=3
    )
    print(f"  {len(cp_raw)} raw contas a pagar")

    # 3. Process records
    all_records = []
    monthly_desp = defaultdict(lambda: {"total": 0.0, "count": 0})
    monthly_rec = defaultdict(lambda: {"total": 0.0, "count": 0})

    for cp in cp_raw:
        situacao = cp.get("status_titulo", "")
        if situacao == "CANCELADO":
            continue

        # Resolve fornecedor name
        forn_id = cp.get("codigo_cliente_fornecedor", 0)
        forn_name = client_map.get(forn_id, f"Fornecedor {forn_id}")

        # Dates
        data_emissao = cp.get("data_emissao", "")
        data_vencimento = cp.get("data_vencimento", "")
        data_previsao = cp.get("data_previsao", "")

        # Primary date: data_previsao -> data_vencimento -> data_emissao
        primary_date = data_previsao or data_vencimento or data_emissao
        ano, mes = parse_br_date(primary_date)

        if not ano or not mes:
            continue

        valor = cp.get("valor_documento", 0.0)
        try:
            valor = float(valor)
        except (ValueError, TypeError):
            valor = 0.0

        categoria = cp.get("codigo_categoria", "")
        projeto = cp.get("codigo_projeto", "")

        # Map situacao
        sit_map = {
            "LIQUIDADO": "PAGO",
            "ATRASADO": "ATRASADO",
            "A VENCER": "A VENCER",
        }
        situacao_clean = sit_map.get(situacao, situacao)

        record = {
            "data": primary_date,
            "data_emissao": data_emissao,
            "data_vencimento": data_vencimento,
            "data_previsao": data_previsao,
            "mes": mes,
            "ano": ano,
            "fornecedor": forn_name,
            "categoria": categoria,
            "valor_pago": round(valor, 2),
            "situacao": situacao_clean,
            "projeto": str(projeto) if projeto else "",
            "fonte": "BMA"
        }
        all_records.append(record)

        # Aggregate into monthly
        # All BMA records are despesas (contas a pagar)
        # But some categories might be receitas (repasses from Better to BMA and back)
        # For now, treat all as despesas since they come from contas a pagar
        monthly_desp[mes]["total"] += valor
        monthly_desp[mes]["count"] += 1

    print(f"  {len(all_records)} valid records (after removing CANCELADO)")

    # Sort records by date
    def date_sort_key(r):
        d = r.get("data", "01/01/2000")
        parts = d.split("/")
        if len(parts) == 3:
            return f"{parts[2]}{parts[1]}{parts[0]}"
        return "00000000"

    all_records.sort(key=date_sort_key, reverse=True)

    # Build monthly arrays sorted chronologically
    def mes_sort_key(m):
        parts = m.split("/")
        if len(parts) == 2:
            month_num = {v: k for k, v in MESES.items()}.get(parts[0], "00")
            return f"{parts[1]}{month_num}"
        return "000000"

    desp_list = []
    for m in sorted(monthly_desp.keys(), key=mes_sort_key):
        desp_list.append({
            "m": m,
            "total": round(monthly_desp[m]["total"], 2),
            "count": monthly_desp[m]["count"]
        })

    rec_list = []
    for m in sorted(monthly_rec.keys(), key=mes_sort_key):
        rec_list.append({
            "m": m,
            "total": round(monthly_rec[m]["total"], 2),
            "count": monthly_rec[m]["count"]
        })

    # 4. Write bma_data.js
    bma_js_path = os.path.join(OUTPUT_DIR, "bma_data.js")
    with open(bma_js_path, "w", encoding="utf-8") as f:
        f.write("var bmaDesp = ")
        f.write(json.dumps(desp_list, ensure_ascii=False))
        f.write(";\n")
        f.write("var bmaRec = ")
        f.write(json.dumps(rec_list, ensure_ascii=False))
        f.write(";\n")
        f.write("var bmaAll = ")
        f.write(json.dumps(all_records, ensure_ascii=False))
        f.write(";\n")
    print(f"\n  Wrote {bma_js_path}")
    print(f"    bmaDesp: {len(desp_list)} months")
    print(f"    bmaRec: {len(rec_list)} months")
    print(f"    bmaAll: {len(all_records)} records")

    # 5. Update dashboard_data.js with bma_desp/bma_rec fields
    dash_path = os.path.join(OUTPUT_DIR, "dashboard_data.js")
    with open(dash_path, "r", encoding="utf-8") as f:
        dash_content = f.read()

    # Parse existing dashData
    # Extract JSON between first [ and last ];
    start = dash_content.index("[")
    end = dash_content.index("];", start) + 1
    dash_json_str = dash_content[start:end]
    dash_data = json.loads(dash_json_str)

    # Build lookup from monthly BMA data
    bma_desp_lookup = {d["m"]: d for d in desp_list}
    bma_rec_lookup = {d["m"]: d for d in rec_list}

    for entry in dash_data:
        m = entry["m"]
        bd = bma_desp_lookup.get(m, {"total": 0, "count": 0})
        br = bma_rec_lookup.get(m, {"total": 0, "count": 0})
        entry["bma_desp"] = round(bd["total"], 2)
        entry["bma_desp_n"] = bd["count"]
        entry["bma_rec"] = round(br["total"], 2)
        entry["bma_rec_n"] = br["count"]

    # Rewrite dashboard_data.js
    new_dash = "var dashData = " + json.dumps(dash_data, ensure_ascii=False) + ";\n"
    new_dash += "var yearSummary = {};\n"
    with open(dash_path, "w", encoding="utf-8") as f:
        f.write(new_dash)
    print(f"  Updated {dash_path} with bma_desp/bma_rec fields")

    # 6. Update despesas_all_data.js - append BMA records
    desp_path = os.path.join(OUTPUT_DIR, "despesas_all_data.js")
    with open(desp_path, "r", encoding="utf-8") as f:
        desp_content = f.read()

    # Parse existing data array
    d_start = desp_content.index("[")
    d_end = desp_content.rindex("]") + 1
    existing_records = json.loads(desp_content[d_start:d_end])

    # Remove any previous BMA records (in case of re-run)
    existing_records = [r for r in existing_records if r.get("fonte") != "BMA"]

    # Add fonte:"Better" to existing records that don't have a fonte field
    for r in existing_records:
        if "fonte" not in r:
            r["fonte"] = "Better"

    # Append BMA records
    combined = existing_records + all_records
    print(f"  Combined despesas: {len(existing_records)} Better + {len(all_records)} BMA = {len(combined)} total")

    new_desp = "var data = " + json.dumps(combined, ensure_ascii=False) + ";\n"
    with open(desp_path, "w", encoding="utf-8") as f:
        f.write(new_desp)
    print(f"  Updated {desp_path}")

    # Summary
    total_bma_desp = sum(d["total"] for d in desp_list)
    total_bma_rec = sum(d["total"] for d in rec_list)
    atrasados = sum(1 for r in all_records if r["situacao"] == "ATRASADO")
    print(f"\n=== BMA Summary ===")
    print(f"  Total Despesas: R$ {total_bma_desp:,.2f}")
    print(f"  Total Receitas: R$ {total_bma_rec:,.2f}")
    print(f"  Resultado: R$ {total_bma_rec - total_bma_desp:,.2f}")
    print(f"  Atrasados: {atrasados}")
    print(f"  Records: {len(all_records)}")
    print(f"\nDone!")


if __name__ == "__main__":
    main()
