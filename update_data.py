#!/usr/bin/env python3
"""
Better Financial Dashboard — Incremental Data Updater (Neon PostgreSQL)

Connects to Neon DB, incrementally fetches new/updated records from each API,
upserts into the database, then regenerates all JS data files from SQL queries.

Run daily via GitHub Actions or locally.
"""
import json, urllib.request, urllib.parse, base64, time, os, sys, shutil
from collections import defaultdict
from datetime import datetime

import psycopg2
from psycopg2.extras import execute_values

# ═══════════════════════════════════════════════════════════════════════
#  Configuration
# ═══════════════════════════════════════════════════════════════════════
DB_URL = os.environ.get(
    "DATABASE_URL",
    "postgresql://neondb_owner:npg_Zu1zG2LPUovb@ep-snowy-shadow-a4hoyxtl-pooler.us-east-1.aws.neon.tech/financeiro?sslmode=require",
)

VINDI_KEY       = os.environ.get('VINDI_KEY', 'pXgDGOG6I5xaYamYFgkjkx0vnqO65rksLWBaU3YIZQU')
PAGARME_KEY     = os.environ.get('PAGARME_KEY', 'ak_live_k8BpvdV4wbWVr3fkQuBOWPOQ0GKXYz')
OMIE_APP_KEY    = os.environ.get('OMIE_APP_KEY', '4340156993172')
OMIE_APP_SECRET = os.environ.get('OMIE_APP_SECRET', 'dd4651357eabc69d5381e8b47a293eb0')
BMA_KEY         = os.environ.get('BMA_KEY', '4602985397010')
BMA_SECRET      = os.environ.get('BMA_SECRET', 'e7e7e8ebffe8c76051459f4dbbb468e5')

VINDI_AUTH = base64.b64encode(f"{VINDI_KEY}:".encode()).decode()
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

MESES = {'01':'Jan','02':'Fev','03':'Mar','04':'Abr','05':'Mai','06':'Jun',
         '07':'Jul','08':'Ago','09':'Set','10':'Out','11':'Nov','12':'Dez'}
MESES_REV = {v: k for k, v in MESES.items()}
MO = {'Jan':1,'Fev':2,'Mar':3,'Abr':4,'Mai':5,'Jun':6,
      'Jul':7,'Ago':8,'Set':9,'Out':10,'Nov':11,'Dez':12}


# ═══════════════════════════════════════════════════════════════════════
#  Helpers
# ═══════════════════════════════════════════════════════════════════════
def safe_float(v):
    if v is None: return 0.0
    try: return float(v)
    except (ValueError, TypeError): return 0.0

def safe_str(v):
    if v is None: return ""
    return str(v).strip()

def date_to_br(v):
    """Convert DB date (datetime.date or ISO string) to DD/MM/YYYY."""
    if v is None: return ""
    if hasattr(v, 'strftime'):
        return v.strftime('%d/%m/%Y')
    s = str(v).strip()
    if len(s) == 10 and s[4] == '-':
        return f"{s[8:10]}/{s[5:7]}/{s[:4]}"
    return s

def get_conn():
    conn = psycopg2.connect(DB_URL, keepalives=1, keepalives_idle=30,
                            keepalives_interval=10, keepalives_count=5)
    conn.autocommit = False
    return conn

def parse_br_date(dt):
    """Parse dd/mm/yyyy -> (year, 'Mes/YYYY')."""
    if not dt or '/' not in dt: return None, None
    parts = dt.split('/')
    if len(parts) < 3: return None, None
    return parts[2], MESES.get(parts[1], '') + '/' + parts[2]

def mes_sort_key(m):
    """Sort key for 'Mes/YYYY' strings."""
    if '/' not in m: return (0, 0)
    parts = m.split('/')
    return (int(parts[1]) if parts[1].isdigit() else 0, MO.get(parts[0], 0))


# ═══════════════════════════════════════════════════════════════════════
#  Batch upsert with reconnect (Neon drops SSL on long ops)
# ═══════════════════════════════════════════════════════════════════════
BATCH_SIZE = 200

def batch_upsert(table, cols, rows, conflict_col, update_cols=None):
    """Insert rows in batches of BATCH_SIZE with ON CONFLICT upsert.
    Reconnects on SSL/connection errors (Neon pooler drops long connections)."""
    if not rows:
        return
    total_inserted = 0
    for i in range(0, len(rows), BATCH_SIZE):
        batch = rows[i:i + BATCH_SIZE]
        for attempt in range(3):
            conn = None
            try:
                conn = get_conn()
                cur = conn.cursor()
                if update_cols:
                    set_clause = ', '.join(f'{c}=EXCLUDED.{c}' for c in update_cols)
                    sql = (f"INSERT INTO {table} ({','.join(cols)}) VALUES %s "
                           f"ON CONFLICT ({conflict_col}) DO UPDATE SET {set_clause}")
                else:
                    sql = (f"INSERT INTO {table} ({','.join(cols)}) VALUES %s "
                           f"ON CONFLICT ({conflict_col}) DO NOTHING")
                execute_values(cur, sql, batch, page_size=50)
                conn.commit()
                cur.close()
                conn.close()
                total_inserted += len(batch)
                break
            except Exception as e:
                print(f"  batch_upsert {table} batch {i//BATCH_SIZE+1} attempt {attempt+1}: {e}")
                try:
                    if conn: conn.close()
                except Exception:
                    pass
                if attempt < 2:
                    time.sleep(3)
        else:
            print(f"  FAILED batch {i//BATCH_SIZE+1} for {table} after 3 attempts")
    return total_inserted


# ═══════════════════════════════════════════════════════════════════════
#  API helpers
# ═══════════════════════════════════════════════════════════════════════
def vindi_get(endpoint, query="", max_pages=500):
    results, page = [], 1
    while page <= max_pages:
        q = urllib.parse.quote(query) if query else ""
        url = (f"https://app.vindi.com.br/api/v1/{endpoint}"
               f"?per_page=50&page={page}&sort_by=created_at&sort_order=asc")
        if q: url += f"&query={q}"
        req = urllib.request.Request(url)
        req.add_header('Authorization', f'Basic {VINDI_AUTH}')
        for attempt in range(3):
            try:
                with urllib.request.urlopen(req, timeout=30) as resp:
                    data = json.loads(resp.read())
                break
            except Exception as e:
                print(f"  Vindi error page {page} attempt {attempt+1}: {e}")
                time.sleep(3)
        else:
            page += 1; continue
        items = data.get(endpoint, [])
        if not items: break
        results.extend(items)
        if len(items) < 50: break
        page += 1
        time.sleep(0.5)
    return results


def omie_call(endpoint, method, params, app_key=None, app_secret=None):
    if app_key is None: app_key = OMIE_APP_KEY
    if app_secret is None: app_secret = OMIE_APP_SECRET
    body = json.dumps({
        "call": method, "app_key": app_key, "app_secret": app_secret,
        "param": [params]
    }).encode()
    req = urllib.request.Request(
        f"https://app.omie.com.br/api/v1/{endpoint}",
        data=body, headers={"Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read())


def omie_paginate(endpoint, method, list_key, per_page=50, delay=3,
                  app_key=None, app_secret=None, recent_pages=None):
    """Paginate an Omie API endpoint.
    If recent_pages is set, only fetch the last N pages (for incremental runs)."""
    if app_key is None: app_key = OMIE_APP_KEY
    if app_secret is None: app_secret = OMIE_APP_SECRET
    results, page, total_pages = [], 1, None
    skipped_to = False
    while True:
        for attempt in range(3):
            try:
                d = omie_call(endpoint, method,
                              {"pagina": page, "registros_por_pagina": per_page,
                               "apenas_importado_api": "N"},
                              app_key, app_secret)
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
            total_records = d.get('total_de_registros', 0)
            print(f"  {method}: {total_records} records, {total_pages} pages")
            # Skip to last N pages if recent_pages is set
            if recent_pages and total_pages > recent_pages and not skipped_to:
                page = total_pages - recent_pages + 1
                skipped_to = True
                print(f"  -> Skipping to page {page}/{total_pages} (last {recent_pages} pages)")
                time.sleep(delay)
                continue
        results.extend(d.get(list_key, []))
        if page % 10 == 0: print(f"    Page {page}/{total_pages}")
        if page >= total_pages: break
        page += 1
        time.sleep(delay)
    return results


# ═══════════════════════════════════════════════════════════════════════
#  1. INCREMENTAL FETCH + UPSERT
# ═══════════════════════════════════════════════════════════════════════

def update_vindi_bills():
    """Fetch all Vindi paid + pending bills and upsert into DB."""
    print("[1/8] Vindi bills ...")

    # Get max created_at from DB for incremental fetch
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COALESCE(MAX(created_at::date), '2020-01-01') FROM vindi_bills")
    max_date = str(cur.fetchone()[0])
    cur.close()
    conn.close()
    print(f"  Max existing created_at: {max_date}")

    new_count, updated_count = 0, 0

    # Paid bills: incremental from max_date
    paid_bills = vindi_get("bills", f"status:paid created_at>={max_date}")
    print(f"  Fetched {len(paid_bills)} paid bills (since {max_date})")

    # Pending bills: always fetch all (status changes)
    pending_bills = vindi_get("bills", "status:pending")
    print(f"  Fetched {len(pending_bills)} pending bills")

    all_bills = paid_bills + pending_bills
    rows = []
    seen_ids = set()
    for b in all_bills:
        bid = b.get('id')
        if not bid or bid in seen_ids: continue
        seen_ids.add(bid)
        dt = (b.get('created_at') or '')[:10]
        mm = dt[5:7] if len(dt) >= 7 else ''
        y = dt[:4] if len(dt) >= 4 else ''
        mes = MESES.get(mm, '') + '/' + y if mm and y else ''
        cust = b.get('customer') or {}
        rows.append((
            bid,
            safe_str(cust.get('name', '')),
            safe_str(cust.get('name', '')),
            safe_str(cust.get('email', '')),
            safe_str(cust.get('registry_code', '')),
            safe_float(b.get('amount')),
            b.get('status', 'unknown'),
            dt if dt else None,
            mes,
            y,
        ))

    cols = ['vindi_id', 'cliente', 'customer_name', 'customer_email',
            'customer_cpf', 'amount', 'status', 'created_at', 'mes', 'ano']
    update_cols = ['amount', 'status', 'customer_name', 'customer_email']

    n = batch_upsert('vindi_bills', cols, rows, 'vindi_id', update_cols)
    print(f"  Vindi bills: upserted {n} rows")
    return n, 0


def _cache_fresh(fonte, max_age_days=7):
    """Return True if fonte was updated less than max_age_days ago."""
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT MAX(timestamp) FROM update_log WHERE fonte=%s", (fonte,))
        last = cur.fetchone()[0]
        cur.close()
        conn.close()
        if last and (datetime.now() - last).days < max_age_days:
            return True
    except Exception:
        pass
    return False


def update_vindi_customers():
    """Fetch all Vindi customers and upsert. Skipped if cache <7 days old."""
    print("[2/8] Vindi customers ...")
    if _cache_fresh('vindi_customers', 7):
        print("  Cache fresh (<7 days), skipping")
        return 0, 0
    customers = vindi_get("customers")
    print(f"  Fetched {len(customers)} customers")
    rows = []
    for c in customers:
        phones_raw = c.get('phones', [])
        phone = ''
        if phones_raw and isinstance(phones_raw, list) and len(phones_raw) > 0:
            p = phones_raw[0]
            phone = safe_str(p.get('number', '')) if isinstance(p, dict) else safe_str(p)
        rows.append((
            c.get('id'), safe_str(c.get('name')), safe_str(c.get('email')),
            safe_str(c.get('registry_code')), phone,
            (c.get('created_at') or '')[:10] or None,
        ))

    cols = ['vindi_id', 'name', 'email', 'registry_code', 'phone', 'created_at']
    update_cols = ['name', 'email', 'registry_code', 'phone']
    n = batch_upsert('vindi_customers', cols, rows, 'vindi_id', update_cols)
    print(f"  Upserted {n} vindi_customers")
    log_update('vindi_customers', n, 0, n, 0)
    return n, 0


def update_omie_contas_pagar():
    """Fetch Omie Better contas a pagar and clients, upsert all.
    Clients: skipped if cache <7 days old.
    CP: only last 10 pages (recent records where status changes happen)."""
    print("[3/8] Omie Better contas a pagar + clients ...")

    # Clients — skip if recently fetched
    cli_map = {}
    if _cache_fresh('omie_clientes', 7):
        print("  Better clients cache fresh (<7 days), loading from DB ...")
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT omie_cod, nome_fantasia, razao_social FROM omie_clientes WHERE fonte='Better'")
        for row in cur.fetchall():
            cli_map[row[0]] = safe_str(row[1]) or safe_str(row[2]) or str(row[0])
        cur.close()
        conn.close()
        print(f"  Loaded {len(cli_map)} Better clients from DB cache")
    else:
        print("  Fetching Better clients ...")
        clients_raw = omie_paginate("geral/clientes/", "ListarClientes", "clientes_cadastro",
                                    delay=3)
        cli_rows = []
        for c in clients_raw:
            cod = c.get('codigo_cliente_omie')
            nome_f = safe_str(c.get('nome_fantasia', ''))
            razao = safe_str(c.get('razao_social', ''))
            cnpj = safe_str(c.get('cnpj_cpf', ''))
            cli_map[cod] = nome_f or razao or str(cod)
            cli_rows.append((cod, nome_f, razao, cnpj, 'Better'))

        n_cli = batch_upsert('omie_clientes', ['omie_cod', 'nome_fantasia', 'razao_social', 'cnpj_cpf', 'fonte'],
                             cli_rows, 'omie_cod',
                             ['nome_fantasia', 'razao_social', 'cnpj_cpf'])
        print(f"  Upserted {n_cli} Better clients")
        log_update('omie_clientes', n_cli, 0, n_cli, 0)

    # Contas a pagar — only last 10 pages (status changes happen on recent records)
    print("  Fetching Better contas a pagar (last 10 pages) ...")
    cp_raw = omie_paginate("financas/contapagar/", "ListarContasPagar", "conta_pagar_cadastro",
                           delay=3, recent_pages=10)
    rows = []
    for c in cp_raw:
        forn_cod = c.get('codigo_cliente_fornecedor')
        forn_nome = cli_map.get(forn_cod, safe_str(forn_cod))
        rows.append((
            c.get('codigo_lancamento_omie'), forn_cod, forn_nome,
            safe_str(c.get('data_emissao')), safe_str(c.get('data_vencimento')),
            safe_str(c.get('data_previsao')), safe_float(c.get('valor_documento')),
            safe_str(c.get('status_titulo')), safe_str(c.get('codigo_categoria')),
            'Better',
        ))

    cols = ['omie_id', 'fornecedor_cod', 'fornecedor_nome', 'data_emissao',
            'data_vencimento', 'data_previsao', 'valor_documento',
            'status_titulo', 'codigo_categoria', 'fonte']
    update_cols = ['fornecedor_nome', 'data_previsao', 'valor_documento',
                   'status_titulo', 'codigo_categoria']
    n = batch_upsert('omie_contas_pagar', cols, rows, 'omie_id', update_cols)
    print(f"  Upserted {n} Better contas a pagar")
    return n, 0


def update_bma():
    """Fetch BMA contas a pagar + receber + clients, upsert all.
    Clients: skipped if cache <7 days old.
    CP: only last 5 pages."""
    print("[4/8] BMA data ...")

    # BMA clients — skip if recently fetched
    bma_cli_map = {}
    if _cache_fresh('bma_clientes', 7):
        print("  BMA clients cache fresh (<7 days), loading from DB ...")
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT omie_cod, razao_social, nome_fantasia FROM omie_clientes WHERE fonte='BMA'")
        for row in cur.fetchall():
            bma_cli_map[row[0]] = safe_str(row[1]) or safe_str(row[2]) or str(row[0])
        cur.close()
        conn.close()
        print(f"  Loaded {len(bma_cli_map)} BMA clients from DB cache")
    else:
        print("  Fetching BMA clients ...")
        bma_clients = omie_paginate("geral/clientes/", "ListarClientes", "clientes_cadastro",
                                    delay=3, app_key=BMA_KEY, app_secret=BMA_SECRET)
        cli_rows = []
        for c in bma_clients:
            cod = c.get('codigo_cliente_omie')
            nome_f = safe_str(c.get('nome_fantasia', ''))
            razao = safe_str(c.get('razao_social', ''))
            cnpj = safe_str(c.get('cnpj_cpf', ''))
            bma_cli_map[cod] = razao or nome_f or str(cod)
            cli_rows.append((cod, nome_f, razao, cnpj, 'BMA'))

        n_cli = batch_upsert('omie_clientes', ['omie_cod', 'nome_fantasia', 'razao_social', 'cnpj_cpf', 'fonte'],
                             cli_rows, 'omie_cod',
                             ['nome_fantasia', 'razao_social'])
        print(f"  Upserted {n_cli} BMA clients")
        log_update('bma_clientes', n_cli, 0, n_cli, 0)

    # BMA contas a pagar — only last 5 pages
    print("  Fetching BMA contas a pagar (last 5 pages) ...")
    bma_cp = omie_paginate("financas/contapagar/", "ListarContasPagar", "conta_pagar_cadastro",
                           delay=3, app_key=BMA_KEY, app_secret=BMA_SECRET,
                           recent_pages=5)
    cp_rows = []
    for c in bma_cp:
        forn_cod = c.get('codigo_cliente_fornecedor')
        forn_nome = bma_cli_map.get(forn_cod, safe_str(forn_cod))
        cp_rows.append((
            c.get('codigo_lancamento_omie'), forn_cod, forn_nome,
            safe_str(c.get('data_emissao')), safe_str(c.get('data_vencimento')),
            safe_str(c.get('data_previsao')), safe_float(c.get('valor_documento')),
            safe_str(c.get('status_titulo')), safe_str(c.get('codigo_categoria')),
            safe_str(c.get('codigo_projeto', '')),
        ))

    n_cp = batch_upsert('bma_contas_pagar',
                        ['omie_id', 'fornecedor_cod', 'fornecedor_nome', 'data_emissao',
                         'data_vencimento', 'data_previsao', 'valor_documento',
                         'status_titulo', 'codigo_categoria', 'projeto'],
                        cp_rows, 'omie_id',
                        ['fornecedor_nome', 'valor_documento', 'status_titulo'])
    print(f"  Upserted {n_cp} BMA contas a pagar")

    # BMA contas a receber
    print("  Fetching BMA contas a receber ...")
    bma_cr = omie_paginate("financas/contareceber/", "ListarContasReceber",
                           "conta_receber_cadastro",
                           delay=3, app_key=BMA_KEY, app_secret=BMA_SECRET)
    cr_rows = []
    for c in bma_cr:
        cli_cod = c.get('codigo_cliente_fornecedor')
        cli_nome = bma_cli_map.get(cli_cod, safe_str(cli_cod))
        cr_rows.append((
            c.get('codigo_lancamento_omie'), cli_cod, cli_nome,
            safe_str(c.get('data_emissao')), safe_str(c.get('data_vencimento')),
            safe_float(c.get('valor_documento')),
            safe_str(c.get('status_titulo')), safe_str(c.get('codigo_categoria')),
        ))

    n_cr = batch_upsert('bma_contas_receber',
                        ['omie_id', 'cliente_cod', 'cliente_nome', 'data_emissao',
                         'data_vencimento', 'valor_documento',
                         'status_titulo', 'codigo_categoria'],
                        cr_rows, 'omie_id',
                        ['valor_documento', 'status_titulo'])
    print(f"  Upserted {n_cr} BMA contas a receber")
    return n_cp, n_cr


def update_omie_nfse():
    """Fetch NFS-e (OS) from Omie and upsert. Only last 5 pages."""
    print("[5/8] Omie NFS-e (last 5 pages) ...")
    os_raw = omie_paginate("servicos/os/", "ListarOS", "osCadastro", delay=3,
                           recent_pages=5)
    rows = []
    for o in os_raw:
        cab = o.get('Cabecalho', {})
        info = o.get('InfoCadastro', {})
        rows.append((
            cab.get('nCodOS'),
            safe_str(info.get('dDtFat', '')),
            safe_float(cab.get('nValorTotal')),
            safe_str(info.get('cEtapa', '')),
            info.get('cCancelada', '') == 'S',
        ))

    n = batch_upsert('omie_nfse',
                     ['omie_id', 'data_faturamento', 'valor_total', 'etapa', 'cancelada'],
                     rows, 'omie_id',
                     ['data_faturamento', 'valor_total', 'etapa', 'cancelada'])
    print(f"  Upserted {n} NFS-e")
    return n, 0


def update_omie_nfe():
    """Fetch NF-e (Pedidos) from Omie and upsert. Only last 3 pages."""
    print("[6/8] Omie NF-e (last 3 pages) ...")
    time.sleep(10)  # cool-down before hitting Omie again
    pedidos = omie_paginate("produtos/pedido/", "ListarPedidos", "pedido_venda_produto",
                            delay=3, recent_pages=3)
    rows = []
    for p in pedidos:
        cab = p.get('cabecalho', {})
        inf = p.get('infoCadastro', {})
        tot = p.get('total_pedido', {})
        rows.append((
            cab.get('codigo_pedido'),
            safe_str(inf.get('dInc', '')),
            safe_str(cab.get('data_previsao', '')),
            safe_float(tot.get('valor_total_pedido')),
            inf.get('cancelado', '') == 'S',
        ))

    n = batch_upsert('omie_nfe',
                     ['omie_id', 'data_inclusao', 'data_previsao', 'valor_total', 'cancelado'],
                     rows, 'omie_id',
                     ['valor_total', 'cancelado'])
    print(f"  Upserted {n} NF-e")
    return n, 0


def update_pagarme():
    """Fetch refunded + chargedback from Pagar.me and upsert."""
    print("[7/8] Pagar.me ...")
    rows = []
    for status in ['refunded', 'chargedback']:
        page = 1
        while True:
            url = (f"https://api.pagar.me/1/transactions?api_key={PAGARME_KEY}"
                   f"&count=1000&page={page}&status={status}")
            try:
                with urllib.request.urlopen(url, timeout=30) as resp:
                    data = json.loads(resp.read())
            except Exception as e:
                print(f"  Pagar.me error: {e}")
                break
            if not data: break
            for t in data:
                cust = t.get('customer') or {}
                rows.append((
                    t.get('tid') or t.get('id'),
                    safe_float(t.get('amount', 0)) / 100.0,
                    status,
                    (t.get('date_created') or '')[:10] or None,
                    safe_str(cust.get('name', '')),
                    safe_str(cust.get('document_number', '')),
                    safe_str(t.get('payment_method', '')),
                ))
            if len(data) < 1000: break
            page += 1
        print(f"  Fetched {len([r for r in rows if r[2]==status])} {status}")

    n = batch_upsert('pagarme_transacoes',
                     ['tid', 'amount', 'status', 'date_created',
                      'customer_name', 'customer_doc', 'payment_method'],
                     rows, 'tid',
                     ['amount', 'status'])
    print(f"  Upserted {n} Pagar.me transactions")
    return n, 0


# ═══════════════════════════════════════════════════════════════════════
#  2. GENERATE JS FILES FROM DATABASE
# ═══════════════════════════════════════════════════════════════════════

def generate_js_files():
    """Query the database and regenerate all JS data files."""
    print("\n=== Generating JS files from DB ===")
    conn = get_conn()
    cur = conn.cursor()

    # ── receitas_all_data.js ─────────────────────────────────────
    print("  receitas_all_data.js ...")
    cur.execute("""
        SELECT cliente, amount, status, created_at, mes, ano
        FROM vindi_bills WHERE status = 'paid'
        ORDER BY created_at
    """)
    rec_list = []
    for row in cur.fetchall():
        cliente, amount, status, created_at, mes, ano = row
        rec_list.append({
            'situacao': 'Pago',
            'data': safe_str(created_at),
            'mes': safe_str(mes),
            'ano': safe_str(ano),
            'cliente': safe_str(cliente),
            'categoria': 'Vindi',
            'valor': round(safe_float(amount), 2),
            'projeto': 'Clientes',
            'obs': '',
        })
    with open(os.path.join(OUTPUT_DIR, 'receitas_all_data.js'), 'w') as f:
        f.write('var data = ' + json.dumps(rec_list, ensure_ascii=True) + ';')
    print(f"    {len(rec_list)} records")

    # ── despesas_all_data.js ─────────────────────────────────────
    # Better despesas (excluding veiculos_flag) + BMA despesas
    print("  despesas_all_data.js ...")
    cur.execute("""
        SELECT omie_id, fornecedor_nome, data_emissao, data_vencimento,
               data_previsao, valor_documento, status_titulo, codigo_categoria
        FROM omie_contas_pagar
        WHERE omie_id NOT IN (SELECT omie_id FROM veiculos_flag)
        ORDER BY data_previsao
    """)
    desp_list = []
    for row in cur.fetchall():
        omie_id, forn, d_em, d_vc, d_pr, val, sit, cat = row
        primary = date_to_br(d_pr) or date_to_br(d_vc) or date_to_br(d_em)
        _, mes = parse_br_date(primary)
        y, _ = parse_br_date(primary)
        desp_list.append({
            'data': primary, 'data_emissao': date_to_br(d_em),
            'data_vencimento': date_to_br(d_vc), 'data_previsao': date_to_br(d_pr),
            'mes': mes or '', 'ano': y or '',
            'fornecedor': safe_str(forn), 'categoria': safe_str(cat),
            'valor_pago': round(safe_float(val), 2),
            'situacao': safe_str(sit), 'projeto': '', 'grupo': 'Geral',
            'fonte': 'Better',
        })

    # Add BMA records
    cur.execute("""
        SELECT omie_id, fornecedor_nome, data_emissao, data_vencimento,
               data_previsao, valor_documento, status_titulo, codigo_categoria, projeto
        FROM bma_contas_pagar
        ORDER BY data_previsao
    """)
    for row in cur.fetchall():
        omie_id, forn, d_em, d_vc, d_pr, val, sit, cat, proj = row
        primary = date_to_br(d_pr) or date_to_br(d_vc) or date_to_br(d_em)
        _, mes = parse_br_date(primary)
        y, _ = parse_br_date(primary)
        sit_clean = safe_str(sit)
        sit_map = {'LIQUIDADO': 'PAGO', 'ATRASADO': 'ATRASADO', 'A VENCER': 'A VENCER'}
        sit_clean = sit_map.get(sit_clean, sit_clean)
        desp_list.append({
            'data': primary, 'data_emissao': date_to_br(d_em),
            'data_vencimento': date_to_br(d_vc), 'data_previsao': date_to_br(d_pr),
            'mes': mes or '', 'ano': y or '',
            'fornecedor': safe_str(forn), 'categoria': safe_str(cat),
            'valor_pago': round(safe_float(val), 2),
            'situacao': sit_clean, 'projeto': safe_str(proj),
            'fonte': 'BMA',
        })

    with open(os.path.join(OUTPUT_DIR, 'despesas_all_data.js'), 'w') as f:
        f.write('var data = ' + json.dumps(desp_list, ensure_ascii=False) + ';')
    print(f"    {len(desp_list)} records (Better + BMA)")

    # ── bma_data.js ──────────────────────────────────────────────
    print("  bma_data.js ...")
    cur.execute("""
        SELECT fornecedor_nome, data_emissao, data_vencimento, data_previsao,
               valor_documento, status_titulo, codigo_categoria, projeto
        FROM bma_contas_pagar
        WHERE status_titulo != 'CANCELADO'
        ORDER BY data_previsao
    """)
    bma_all = []
    bma_desp_monthly = defaultdict(lambda: {'total': 0.0, 'count': 0})
    for row in cur.fetchall():
        forn, d_em, d_vc, d_pr, val, sit, cat, proj = row
        primary = date_to_br(d_pr) or date_to_br(d_vc) or date_to_br(d_em)
        y, mes = parse_br_date(primary)
        if not y: continue
        sit_map = {'LIQUIDADO': 'PAGO', 'ATRASADO': 'ATRASADO', 'A VENCER': 'A VENCER'}
        sit_clean = sit_map.get(safe_str(sit), safe_str(sit))
        v = safe_float(val)
        bma_all.append({
            'data': primary, 'data_emissao': date_to_br(d_em),
            'data_vencimento': date_to_br(d_vc), 'data_previsao': date_to_br(d_pr),
            'mes': mes, 'ano': y, 'fornecedor': safe_str(forn),
            'categoria': safe_str(cat), 'valor_pago': round(v, 2),
            'situacao': sit_clean, 'projeto': safe_str(proj), 'fonte': 'BMA',
        })
        bma_desp_monthly[mes]['total'] += v
        bma_desp_monthly[mes]['count'] += 1

    bma_desp_list = [{'m': m, 'total': round(d['total'], 2), 'count': d['count']}
                     for m, d in sorted(bma_desp_monthly.items(), key=lambda x: mes_sort_key(x[0]))]

    # BMA receitas (contas a receber)
    cur.execute("""
        SELECT cliente_nome, data_emissao, data_vencimento,
               valor_documento, status_titulo, codigo_categoria
        FROM bma_contas_receber
        WHERE status_titulo != 'CANCELADO'
        ORDER BY data_vencimento
    """)
    bma_rec_monthly = defaultdict(lambda: {'total': 0.0, 'count': 0})
    for row in cur.fetchall():
        cli, d_em, d_vc, val, sit, cat = row
        primary = date_to_br(d_vc) or date_to_br(d_em)
        y, mes = parse_br_date(primary)
        if not y or not mes: continue
        v = safe_float(val)
        bma_rec_monthly[mes]['total'] += v
        bma_rec_monthly[mes]['count'] += 1

    bma_rec_list = [{'m': m, 'total': round(d['total'], 2), 'count': d['count']}
                    for m, d in sorted(bma_rec_monthly.items(), key=lambda x: mes_sort_key(x[0]))]

    with open(os.path.join(OUTPUT_DIR, 'bma_data.js'), 'w', encoding='utf-8') as f:
        f.write('var bmaDesp = ' + json.dumps(bma_desp_list, ensure_ascii=False) + ';\n')
        f.write('var bmaRec = ' + json.dumps(bma_rec_list, ensure_ascii=False) + ';\n')
        f.write('var bmaAll = ' + json.dumps(bma_all, ensure_ascii=False) + ';\n')
    print(f"    bmaAll: {len(bma_all)}, bmaDesp months: {len(bma_desp_list)}, bmaRec months: {len(bma_rec_list)}")

    # ── veiculos_data.js ─────────────────────────────────────────
    print("  veiculos_data.js ...")
    cur.execute("""
        SELECT cp.fornecedor_nome, cp.data_emissao, cp.data_vencimento,
               cp.data_previsao, cp.valor_documento, cp.status_titulo, cp.codigo_categoria
        FROM omie_contas_pagar cp
        JOIN veiculos_flag v ON cp.omie_id = v.omie_id
        ORDER BY cp.data_previsao
    """)
    veic_list = []
    for row in cur.fetchall():
        forn, d_em, d_vc, d_pr, val, sit, cat = row
        primary = date_to_br(d_pr) or date_to_br(d_vc) or date_to_br(d_em)
        _, mes = parse_br_date(primary)
        y, _ = parse_br_date(primary)
        veic_list.append({
            'data': primary, 'data_emissao': date_to_br(d_em),
            'data_vencimento': date_to_br(d_vc), 'data_previsao': date_to_br(d_pr),
            'mes': mes or '', 'ano': y or '',
            'fornecedor': safe_str(forn), 'categoria': safe_str(cat),
            'valor_pago': round(safe_float(val), 2),
            'situacao': safe_str(sit), 'projeto': '', 'grupo': 'Geral',
        })
    with open(os.path.join(OUTPUT_DIR, 'veiculos_data.js'), 'w') as f:
        f.write('var veiculosData = ' + json.dumps(veic_list, ensure_ascii=False) + ';')
    print(f"    {len(veic_list)} vehicle records")

    # ── dashboard_data.js ────────────────────────────────────────
    print("  dashboard_data.js ...")

    # Receitas by month (Vindi paid)
    cur.execute("""
        SELECT mes, SUM(amount), COUNT(*) FROM vindi_bills
        WHERE status = 'paid' AND mes != '' GROUP BY mes
    """)
    receitas = {r[0]: {'total': float(r[1]), 'count': r[2]} for r in cur.fetchall()}

    # Pending by month (Vindi pending)
    cur.execute("""
        SELECT mes, SUM(amount), COUNT(*) FROM vindi_bills
        WHERE status = 'pending' AND mes != '' GROUP BY mes
    """)
    pending = {r[0]: {'total': float(r[1]), 'count': r[2]} for r in cur.fetchall()}

    # Despesas by month (Better, excl veiculos_flag, excl CANCELADO)
    cur.execute("""
        SELECT
            COALESCE(data_previsao, data_vencimento, data_emissao) as dt,
            valor_documento
        FROM omie_contas_pagar
        WHERE status_titulo != 'CANCELADO'
          AND omie_id NOT IN (SELECT omie_id FROM veiculos_flag)
    """)
    despesas = defaultdict(lambda: {'total': 0, 'count': 0})
    for row in cur.fetchall():
        dt = date_to_br(row[0])
        _, m = parse_br_date(dt)
        if m:
            despesas[m]['total'] += safe_float(row[1])
            despesas[m]['count'] += 1

    # Estornos by month (Pagar.me refunded)
    cur.execute("""
        SELECT date_created, amount FROM pagarme_transacoes WHERE status = 'refunded'
    """)
    estornos = defaultdict(lambda: {'total': 0, 'count': 0})
    for row in cur.fetchall():
        dt = date_to_br(row[0])
        if dt and len(dt) >= 7:
            mm = dt[5:7] if '-' in dt else ''
            y = dt[:4] if '-' in dt else ''
            m = MESES.get(mm, '') + '/' + y if mm and y else ''
            if m:
                estornos[m]['total'] += safe_float(row[1])
                estornos[m]['count'] += 1

    # Chargebacks
    cur.execute("""
        SELECT date_created, amount FROM pagarme_transacoes WHERE status = 'chargedback'
    """)
    chargebacks = defaultdict(lambda: {'total': 0, 'count': 0})
    for row in cur.fetchall():
        dt = date_to_br(row[0])
        if dt and len(dt) >= 7:
            mm = dt[5:7] if '-' in dt else ''
            y = dt[:4] if '-' in dt else ''
            m = MESES.get(mm, '') + '/' + y if mm and y else ''
            if m:
                chargebacks[m]['total'] += safe_float(row[1])
                chargebacks[m]['count'] += 1

    # NFS-e by month
    cur.execute("""
        SELECT data_faturamento, valor_total FROM omie_nfse
        WHERE NOT cancelada
    """)
    nfse_m = defaultdict(lambda: {'total': 0, 'count': 0})
    for row in cur.fetchall():
        dt = date_to_br(row[0])
        _, m = parse_br_date(dt)
        if m:
            nfse_m[m]['total'] += safe_float(row[1])
            nfse_m[m]['count'] += 1

    # NF-e by month
    cur.execute("""
        SELECT data_inclusao, valor_total FROM omie_nfe
        WHERE NOT cancelado
    """)
    nfe_m = defaultdict(lambda: {'total': 0, 'count': 0})
    for row in cur.fetchall():
        dt = date_to_br(row[0])
        _, m = parse_br_date(dt)
        if m:
            nfe_m[m]['total'] += safe_float(row[1])
            nfe_m[m]['count'] += 1

    # BMA despesas by month (reuse from bma_data generation)
    bma_desp_m = {d['m']: d for d in bma_desp_list}
    # BMA receitas by month
    bma_rec_m = {d['m']: d for d in bma_rec_list}

    # Combine all months
    all_months = set()
    for d in [receitas, despesas, estornos, nfse_m, nfe_m, pending]:
        all_months.update(d.keys())
    sorted_months = sorted(all_months, key=mes_sort_key)

    dash_data = []
    for m in sorted_months:
        parts = m.split('/')
        if len(parts) != 2: continue
        rec = receitas.get(m, {'total': 0, 'count': 0})
        desp = despesas.get(m, {'total': 0, 'count': 0})
        est = estornos.get(m, {'total': 0, 'count': 0})
        chb = chargebacks.get(m, {'total': 0, 'count': 0})
        pend = pending.get(m, {'total': 0, 'count': 0})
        nfse = nfse_m.get(m, {'total': 0, 'count': 0})
        nfe = nfe_m.get(m, {'total': 0, 'count': 0})
        bd = bma_desp_m.get(m, {'total': 0, 'count': 0})
        br = bma_rec_m.get(m, {'total': 0, 'count': 0})

        dash_data.append({
            'm': m, 'y': parts[1],
            'rec': round(rec['total'], 2), 'rec_n': rec['count'],
            'desp': round(desp['total'], 2), 'desp_n': desp['count'],
            'est': round(est['total'], 2), 'est_n': est['count'],
            'chb': round(chb['total'], 2), 'chb_n': chb['count'],
            'pend': round(pend['total'], 2), 'pend_n': pend['count'],
            'nfse': round(nfse['total'], 2), 'nfse_n': nfse['count'],
            'nfe': round(nfe['total'], 2), 'nfe_n': nfe['count'],
            'bma_desp': round(bd.get('total', 0), 2), 'bma_desp_n': bd.get('count', 0),
            'bma_rec': round(br.get('total', 0), 2), 'bma_rec_n': br.get('count', 0),
            'res': round(rec['total'] - desp['total'] - est['total'], 2),
        })

    with open(os.path.join(OUTPUT_DIR, 'dashboard_data.js'), 'w') as f:
        f.write('var dashData = ' + json.dumps(dash_data, ensure_ascii=True) + ';\n')
        f.write('var yearSummary = {};\n')
    print(f"    {len(dash_data)} months")

    # ── changelog.js ─────────────────────────────────────────────
    print("  changelog.js ...")
    changelog_path = os.path.join(OUTPUT_DIR, 'changelog.js')
    changelog_prev_path = os.path.join(OUTPUT_DIR, 'changelog_prev.js')
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M')

    # Read previous changelog
    prev_update_ts = ""
    try:
        with open(changelog_path, 'r') as f:
            cl_raw = f.read()
        cl_start = cl_raw.index('{')
        cl_end = cl_raw.rindex('}') + 1
        old_changelog = json.loads(cl_raw[cl_start:cl_end])
        prev_update_ts = old_changelog.get('last_update', '')
        shutil.copy2(changelog_path, changelog_prev_path)
    except Exception:
        pass

    # Read previous snapshot from DB
    cur.execute("""
        SELECT mes, rec, desp, est, bma_desp, nfse, nfe, pend
        FROM dashboard_snapshot
        ORDER BY snapshot_at DESC
        LIMIT 100
    """)
    old_snap = {}
    for row in cur.fetchall():
        m = row[0]
        if m not in old_snap:
            old_snap[m] = {
                'rec': float(row[1] or 0), 'desp': float(row[2] or 0),
                'est': float(row[3] or 0), 'nfse': float(row[5] or 0),
                'nfe': float(row[6] or 0), 'pend': float(row[7] or 0),
            }

    field_labels = {
        'rec': 'Receitas', 'desp': 'Despesas', 'est': 'Estornos',
        'nfse': 'NFS-e', 'nfe': 'NF-e', 'pend': 'Pendentes'
    }
    compare_fields = ['rec', 'desp', 'est', 'nfse', 'nfe', 'pend']
    changes, atypical = [], []

    new_by_month = {d['m']: d for d in dash_data}
    for m in sorted(set(list(old_snap.keys()) + list(new_by_month.keys())), key=mes_sort_key):
        old_row = old_snap.get(m, {})
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
                pct = round(abs(diff) / abs(old_val) * 100, 1) if old_val != 0 else 0
                if abs(diff) > 10000 or pct > 10:
                    sign = '+' if diff > 0 else ''
                    pct_str = f"{sign}{pct}%" if old_val != 0 else "novo"
                    atypical.append({
                        'month': m, 'field': field, 'diff': diff, 'pct': pct,
                        'label': f"{field_labels.get(field, field)} {m} {pct_str}"
                    })

    # Summary
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
            parts.append(f"{len(atypical)} alteracao(oes) atipica(s)")
        summary = ' | '.join(parts) if parts else 'Sem alteracoes significativas'
    else:
        summary = 'Sem alteracoes (primeira execucao ou dados identicos)'

    changelog_obj = {
        'last_update': now_str, 'previous_update': prev_update_ts,
        'changes': changes, 'atypical': atypical, 'summary': summary
    }
    with open(changelog_path, 'w') as f:
        f.write('var changelog = ' + json.dumps(changelog_obj, ensure_ascii=False) + ';\n')
    print(f"    {len(changes)} changes, {len(atypical)} atypical")

    # ── Save new dashboard_snapshot ──────────────────────────────
    print("  Saving dashboard_snapshot ...")
    snap_rows = []
    for d in dash_data:
        snap_rows.append((
            d['m'], d['rec'], d['desp'], d['est'],
            d.get('bma_desp', 0), d['nfse'], d['nfe'], d['pend'],
            datetime.now(),
        ))
    if snap_rows:
        execute_values(cur,
            """INSERT INTO dashboard_snapshot (mes, rec, desp, est, bma_desp, nfse, nfe, pend, snapshot_at)
               VALUES %s""",
            snap_rows, page_size=100)
    conn.commit()
    print(f"    Saved {len(snap_rows)} snapshot rows")

    cur.close()
    conn.close()
    return len(dash_data), len(rec_list), len(desp_list)


# ═══════════════════════════════════════════════════════════════════════
#  3. LOG
# ═══════════════════════════════════════════════════════════════════════

def log_update(fonte, new, updated, total, duration):
    for attempt in range(3):
        conn = None
        try:
            conn = get_conn()
            cur = conn.cursor()
            cur.execute(
                """INSERT INTO update_log (fonte, registros_novos, registros_atualizados,
                   total_registros, duracao_segundos, timestamp)
                   VALUES (%s, %s, %s, %s, %s, %s)""",
                (fonte, new, updated, total, duration, datetime.now())
            )
            conn.commit()
            cur.close()
            conn.close()
            break
        except Exception as e:
            print(f"  log_update attempt {attempt+1}: {e}")
            try:
                if conn: conn.close()
            except Exception:
                pass
            time.sleep(2)


# ═══════════════════════════════════════════════════════════════════════
#  Main
# ═══════════════════════════════════════════════════════════════════════
def main():
    start_time = time.time()
    print(f"=== Better Data Update (Neon) — {datetime.now().strftime('%Y-%m-%d %H:%M')} ===\n")

    def _elapsed():
        return f"[{time.time()-start_time:.0f}s elapsed]"

    # ── Phase 1: Incremental fetch + upsert ──────────────────────
    stats = {}

    stats['vindi_bills'] = update_vindi_bills()
    print(f"  {_elapsed()}")

    stats['vindi_customers'] = update_vindi_customers()
    print(f"  {_elapsed()}")

    stats['omie_cp'] = update_omie_contas_pagar()
    print(f"  {_elapsed()}")

    stats['bma'] = update_bma()
    print(f"  {_elapsed()}")

    stats['nfse'] = update_omie_nfse()
    print(f"  {_elapsed()}")

    stats['nfe'] = update_omie_nfe()
    print(f"  {_elapsed()}")

    stats['pagarme'] = update_pagarme()
    print(f"  {_elapsed()}")

    # ── Phase 2: Generate JS files from DB ───────────────────────
    n_months, n_rec, n_desp = generate_js_files()
    print(f"  {_elapsed()}")

    elapsed = time.time() - start_time
    print(f"\n=== Done in {elapsed:.0f}s — Dashboard: {n_months} months, "
          f"Receitas: {n_rec}, Despesas: {n_desp} ===")

    # ── Phase 3: Log to DB ───────────────────────────────────────
    total_new = sum(s[0] for s in stats.values())
    total_upd = sum(s[1] for s in stats.values())
    log_update('update_data', total_new, total_upd, n_rec + n_desp, elapsed)


if __name__ == '__main__':
    main()
