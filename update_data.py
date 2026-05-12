#!/usr/bin/env python3
"""
Better Financial Dashboard - Data Updater
Pulls data from Vindi, Pagar.me and Omie APIs and generates JS data files.
Run daily via cron or GitHub Actions.
"""
import json, urllib.request, base64, time, os
from collections import defaultdict
from datetime import datetime

VINDI_KEY = os.environ.get('VINDI_KEY', 'pXgDGOG6I5xaYamYFgkjkx0vnqO65rksLWBaU3YIZQU')
PAGARME_KEY = os.environ.get('PAGARME_KEY', 'ak_live_k8BpvdV4wbWVr3fkQuBOWPOQ0GKXYz')
OMIE_APP_KEY = os.environ.get('OMIE_APP_KEY', '4340156993172')
OMIE_APP_SECRET = os.environ.get('OMIE_APP_SECRET', 'dd4651357eabc69d5381e8b47a293eb0')

VINDI_AUTH = base64.b64encode(f"{VINDI_KEY}:".encode()).decode()
MESES = {'01':'Jan','02':'Fev','03':'Mar','04':'Abr','05':'Mai','06':'Jun','07':'Jul','08':'Ago','09':'Set','10':'Out','11':'Nov','12':'Dez'}

def vindi_get(endpoint, query="", max_pages=500):
    results = []
    page = 1
    while page <= max_pages:
        q = urllib.parse.quote(query) if query else ""
        url = f"https://app.vindi.com.br/api/v1/{endpoint}?per_page=50&page={page}&sort_by=created_at&sort_order=asc"
        if q: url += f"&query={q}"
        req = urllib.request.Request(url)
        req.add_header('Authorization', f'Basic {VINDI_AUTH}')
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = json.loads(resp.read())
        except Exception as e:
            print(f"  Vindi error page {page}: {e}")
            time.sleep(3)
            continue
        items = data.get(endpoint, [])
        if not items: break
        results.extend(items)
        if len(items) < 50: break
        page += 1
        time.sleep(0.5)
    return results

def omie_call(endpoint, method, params):
    body = json.dumps({"call": method, "app_key": OMIE_APP_KEY, "app_secret": OMIE_APP_SECRET, "param": [params]}).encode()
    req = urllib.request.Request(f"https://app.omie.com.br/api/v1/{endpoint}", data=body, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read())

def omie_paginate(endpoint, method, list_key, per_page=50, delay=5):
    results = []
    page = 1
    total_pages = None
    while True:
        for attempt in range(3):
            try:
                d = omie_call(endpoint, method, {"pagina": page, "registros_por_pagina": per_page, "apenas_importado_api": "N"})
                break
            except Exception as e:
                print(f"  Omie retry {attempt+1} page {page}: {e}")
                time.sleep(delay * (attempt + 1))
        else:
            print(f"  Omie SKIP page {page}")
            page += 1
            if total_pages and page > total_pages: break
            continue
        if total_pages is None:
            total_pages = d.get('total_de_paginas', 0)
            print(f"  {method}: {d.get('total_de_registros',0)} records, {total_pages} pages")
        results.extend(d.get(list_key, []))
        if page % 10 == 0: print(f"    Page {page}/{total_pages}")
        if page >= total_pages: break
        page += 1
        time.sleep(delay)
    return results

def clean_doc(doc):
    return ''.join(c for c in str(doc) if c.isdigit())

def parse_br_date(dt):
    if not dt or '/' not in dt: return None, None
    parts = dt.split('/')
    if len(parts) < 3: return None, None
    return parts[2], MESES.get(parts[1],'') + '/' + parts[2]

def main():
    output_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"=== Better Data Update - {datetime.now().strftime('%Y-%m-%d %H:%M')} ===\n")

    # 1. VINDI
    print("[1/6] Vindi paid bills...")
    vindi_paid = vindi_get("bills", "status:paid")
    print(f"  {len(vindi_paid)} paid bills")

    print("[2/6] Vindi pending + customers...")
    vindi_pending = vindi_get("bills", "status:pending")
    vindi_customers = vindi_get("customers")
    print(f"  {len(vindi_pending)} pending, {len(vindi_customers)} customers")

    # 2. PAGAR.ME
    print("[3/6] Pagar.me refunds...")
    pm_refunds = []
    for status in ['refunded']:
        page = 1
        while True:
            url = f"https://api.pagar.me/1/transactions?api_key={PAGARME_KEY}&count=1000&page={page}&status={status}"
            try:
                with urllib.request.urlopen(url, timeout=30) as resp:
                    data = json.loads(resp.read())
            except: break
            if not data: break
            for t in data:
                cust = t.get('customer') or {}
                pm_refunds.append({'amount': t.get('amount',0)/100, 'date': t.get('date_created','')[:10], 'customer_doc': cust.get('document_number','')})
            if len(data) < 1000: break
            page += 1
    print(f"  {len(pm_refunds)} refunds")

    # 3. OMIE
    print("[4/6] Omie contas a pagar...")
    cp_raw = omie_paginate("financas/contapagar/", "ListarContasPagar", "conta_pagar_cadastro")
    print(f"  {len(cp_raw)} contas a pagar")

    print("[5/6] Omie OS + Pedidos + Clientes...")
    clients_raw = omie_paginate("geral/clientes/", "ListarClientes", "clientes_cadastro")
    os_raw = omie_paginate("servicos/os/", "ListarOS", "osCadastro")
    print(f"  {len(os_raw)} OS, {len(clients_raw)} clients")

    time.sleep(10)
    pedidos_raw = omie_paginate("produtos/pedido/", "ListarPedidos", "pedido_venda_produto")
    print(f"  {len(pedidos_raw)} pedidos")

    # 4. PROCESS DATA
    print("[6/6] Processing...")

    # Build client map
    cli_map = {}
    for c in clients_raw:
        cod = c.get('codigo_cliente_omie')
        cli_map[cod] = {'nome': c.get('nome_fantasia','') or c.get('razao_social',''), 'cpf_cnpj': c.get('cnpj_cpf','')}

    # Dashboard data
    vcust_doc = {}
    for c in vindi_customers:
        d = clean_doc(c.get('registry_code',''))
        if d: vcust_doc[c['id']] = d

    receitas = defaultdict(lambda: {'total':0,'count':0})
    for b in vindi_paid:
        dt = (b.get('created_at') or '')[:10]
        if not dt or len(dt) < 7: continue
        m = MESES.get(dt[5:7],'') + '/' + dt[:4]
        receitas[m]['total'] += float(b.get('amount',0))
        receitas[m]['count'] += 1

    estornos = defaultdict(lambda: {'total':0,'count':0})
    for r in pm_refunds:
        dt = r.get('date','')[:10]
        if not dt or len(dt) < 7: continue
        m = MESES.get(dt[5:7],'') + '/' + dt[:4]
        estornos[m]['total'] += r['amount']
        estornos[m]['count'] += 1

    despesas = defaultdict(lambda: {'total':0,'count':0})
    for c in cp_raw:
        if c.get('status_titulo') == 'CANCELADO': continue
        dt = c.get('data_previsao','') or c.get('data_vencimento','') or c.get('data_emissao','')
        y, m = parse_br_date(dt)
        if y and m:
            despesas[m]['total'] += c.get('valor_documento',0)
            despesas[m]['count'] += 1

    nfse_m = defaultdict(lambda: {'total':0,'count':0})
    for o in os_raw:
        cab = o.get('Cabecalho',{})
        info = o.get('InfoCadastro',{})
        if info.get('cCancelada') == 'S': continue
        dt = info.get('dDtFat','') or info.get('dDtInc','')
        y, m = parse_br_date(dt)
        if y and m:
            nfse_m[m]['total'] += cab.get('nValorTotal',0)
            nfse_m[m]['count'] += 1

    nfe_m = defaultdict(lambda: {'total':0,'count':0})
    for p in pedidos_raw:
        cab = p.get('cabecalho',{})
        inf = p.get('infoCadastro',{})
        if inf.get('cancelado') == 'S': continue
        dt = inf.get('dInc','') or cab.get('data_previsao','')
        y, m = parse_br_date(dt)
        if y and m:
            tot = p.get('total_pedido',{}).get('valor_total_pedido',0)
            nfe_m[m]['total'] += tot
            nfe_m[m]['count'] += 1

    pend_m = defaultdict(lambda: {'total':0,'count':0})
    for b in vindi_pending:
        dt = (b.get('created_at') or '')[:10]
        if not dt or len(dt) < 7: continue
        m = MESES.get(dt[5:7],'') + '/' + dt[:4]
        pend_m[m]['total'] += float(b.get('amount',0))
        pend_m[m]['count'] += 1

    all_months = set(list(receitas.keys()) + list(despesas.keys()) + list(nfse_m.keys()))
    mo = {'Jan':1,'Fev':2,'Mar':3,'Abr':4,'Mai':5,'Jun':6,'Jul':7,'Ago':8,'Set':9,'Out':10,'Nov':11,'Dez':12}
    sorted_months = sorted(all_months, key=lambda m: (int(m.split('/')[1]) if '/' in m else 0, mo.get(m.split('/')[0],0)))

    dash_data = []
    for m in sorted_months:
        parts = m.split('/')
        if len(parts) != 2: continue
        rec = receitas.get(m, {'total':0,'count':0})
        desp = despesas.get(m, {'total':0,'count':0})
        est = estornos.get(m, {'total':0,'count':0})
        nfse = nfse_m.get(m, {'total':0,'count':0})
        nfe = nfe_m.get(m, {'total':0,'count':0})
        pend = pend_m.get(m, {'total':0,'count':0})
        dash_data.append({
            'm': m, 'y': parts[1],
            'rec': round(rec['total'],2), 'rec_n': rec['count'],
            'desp': round(desp['total'],2), 'desp_n': desp['count'],
            'est': round(est['total'],2), 'est_n': est['count'],
            'chb': 0, 'chb_n': 0,
            'pend': round(pend['total'],2), 'pend_n': pend['count'],
            'nfse': round(nfse['total'],2), 'nfse_n': nfse['count'],
            'nfe': round(nfe['total'],2), 'nfe_n': nfe['count'],
            'res': round(rec['total'] - desp['total'] - est['total'], 2),
        })

    # --- CHANGELOG: Read old data before overwriting ---
    old_dash_data = None
    dash_data_path = os.path.join(output_dir, 'dashboard_data.js')
    try:
        with open(dash_data_path, 'r') as f:
            raw = f.read()
        # Extract JSON between "var dashData = " and ";\n"
        start = raw.index('[')
        end = raw.index('];') + 1
        old_dash_data = json.loads(raw[start:end])
        print(f"  Changelog: read {len(old_dash_data)} months from previous dashboard_data.js")
    except Exception as e:
        print(f"  Changelog: no previous dashboard_data.js to compare ({e})")

    # Save dashboard data
    with open(os.path.join(output_dir, 'dashboard_data.js'), 'w') as f:
        f.write('var dashData = ' + json.dumps(dash_data, ensure_ascii=True) + ';\n')
        f.write('var yearSummary = {};\n')

    # Save receitas
    rec_list = []
    for b in vindi_paid:
        dt = (b.get('created_at') or '')[:10]
        if not dt: continue
        y = dt[:4]; mm = dt[5:7]
        cust = b.get('customer') or {}
        rec_list.append({'situacao':'Pago','data':dt,'mes':MESES.get(mm,'')+'/'+y,'ano':y,'cliente':cust.get('name',''),'categoria':'Vindi','valor':round(float(b.get('amount',0)),2),'projeto':'Clientes','obs':''})
    with open(os.path.join(output_dir, 'receitas_all_data.js'), 'w') as f:
        f.write('var data = ' + json.dumps(rec_list, ensure_ascii=True) + ';')

    # Save despesas
    desp_list = []
    for c in cp_raw:
        if c.get('status_titulo') == 'CANCELADO': continue
        dt_prev = c.get('data_previsao','') or c.get('data_vencimento','') or c.get('data_emissao','')
        y, m = parse_br_date(dt_prev)
        if not y: continue
        forn = c.get('codigo_cliente_fornecedor')
        nome = cli_map.get(forn, {}).get('nome', str(forn))
        desp_list.append({'data':dt_prev,'data_emissao':c.get('data_emissao',''),'data_vencimento':c.get('data_vencimento',''),'data_previsao':c.get('data_previsao',''),'mes':m or '','ano':y,'fornecedor':nome,'categoria':c.get('codigo_categoria',''),'valor_pago':round(c.get('valor_documento',0),2),'situacao':c.get('status_titulo',''),'projeto':'','grupo':'Geral'})
    with open(os.path.join(output_dir, 'despesas_all_data.js'), 'w') as f:
        f.write('var data = ' + json.dumps(desp_list, ensure_ascii=True) + ';')

    # --- CHANGELOG: Generate changelog comparing old vs new ---
    changelog_path = os.path.join(output_dir, 'changelog.js')
    changelog_prev_path = os.path.join(output_dir, 'changelog_prev.js')
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M')

    # Read previous changelog to get its last_update as our previous_update
    prev_update_ts = ""
    try:
        with open(changelog_path, 'r') as f:
            cl_raw = f.read()
        cl_start = cl_raw.index('{')
        cl_end = cl_raw.rindex('}') + 1
        old_changelog = json.loads(cl_raw[cl_start:cl_end])
        prev_update_ts = old_changelog.get('last_update', '')
        # Rename existing changelog to changelog_prev.js
        import shutil
        shutil.copy2(changelog_path, changelog_prev_path)
        print(f"  Changelog: previous update was {prev_update_ts}")
    except Exception as e:
        print(f"  Changelog: no previous changelog.js ({e})")

    field_labels = {
        'rec': 'Receitas', 'desp': 'Despesas', 'est': 'Estornos',
        'nfse': 'NFS-e', 'nfe': 'NF-e', 'pend': 'Pendentes'
    }
    compare_fields = ['rec', 'desp', 'est', 'nfse', 'nfe', 'pend']
    changes = []
    atypical = []

    if old_dash_data is not None:
        old_by_month = {d['m']: d for d in old_dash_data}
        new_by_month = {d['m']: d for d in dash_data}

        all_cl_months = set(list(old_by_month.keys()) + list(new_by_month.keys()))
        for m in sorted(all_cl_months, key=lambda x: (int(x.split('/')[1]) if '/' in x else 0, mo.get(x.split('/')[0], 0))):
            old_row = old_by_month.get(m, {})
            new_row = new_by_month.get(m, {})
            for field in compare_fields:
                old_val = round(old_row.get(field, 0), 2)
                new_val = round(new_row.get(field, 0), 2)
                diff = round(new_val - old_val, 2)
                if diff != 0:
                    change_entry = {
                        'month': m, 'field': field,
                        'old': old_val, 'new': new_val,
                        'diff': diff, 'label': field_labels.get(field, field)
                    }
                    changes.append(change_entry)
                    # Check atypical: abs diff > 10k or > 10%
                    pct = 0
                    if old_val != 0:
                        pct = round(abs(diff) / abs(old_val) * 100, 1)
                    is_atypical = abs(diff) > 10000 or pct > 10
                    if is_atypical:
                        sign = '+' if diff > 0 else ''
                        pct_str = f"{sign}{pct}%" if old_val != 0 else "novo"
                        atypical.append({
                            'month': m, 'field': field,
                            'diff': diff, 'pct': pct,
                            'label': f"{field_labels.get(field, field)} {m} {pct_str}"
                        })

    # Build summary
    if changes:
        total_rec_diff = sum(c['diff'] for c in changes if c['field'] == 'rec')
        total_desp_diff = sum(c['diff'] for c in changes if c['field'] == 'desp')
        parts = []
        if total_rec_diff != 0:
            sign = '+' if total_rec_diff > 0 else ''
            parts.append(f"Receitas {sign}R$ {abs(total_rec_diff)/1000:.1f}k")
        if total_desp_diff != 0:
            sign = '+' if total_desp_diff > 0 else ''
            parts.append(f"Despesas {sign}R$ {abs(total_desp_diff)/1000:.1f}k")
        if atypical:
            parts.append(f"{len(atypical)} alteração(ões) atípica(s)")
        summary = ' | '.join(parts) if parts else 'Sem alterações significativas'
    else:
        summary = 'Sem alterações (primeira execução ou dados idênticos)'

    changelog_obj = {
        'last_update': now_str,
        'previous_update': prev_update_ts,
        'changes': changes,
        'atypical': atypical,
        'summary': summary
    }
    with open(changelog_path, 'w') as f:
        f.write('var changelog = ' + json.dumps(changelog_obj, ensure_ascii=False) + ';\n')
    print(f"  Changelog: {len(changes)} changes, {len(atypical)} atypical")

    print(f"\n=== Done! Dashboard: {len(dash_data)} months, Receitas: {len(rec_list)}, Despesas: {len(desp_list)} ===")

if __name__ == '__main__':
    import urllib.parse
    main()
