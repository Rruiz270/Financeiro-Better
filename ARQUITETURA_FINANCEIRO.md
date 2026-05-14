# Better Education — Arquitetura do Ecossistema Financeiro

> Documento de referencia com todas as regras, fluxos, fontes de dados e logica de negocios do sistema financeiro Better Education.
> Atualizado: Maio 2026

---

## 1. VISAO GERAL DO ECOSSISTEMA

```
                           FONTES DE DADOS (APIs)
  ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌──────────┐
  │  VINDI   │  │ PAGAR.ME │  │   OMIE   │  │ BMA OMIE │  │ PLANILHA │
  │ Receitas │  │ Estornos │  │ Despesas │  │  Folha   │  │  Vendas  │
  │ Clientes │  │Chargebck │  │  NFS-e   │  │Beneficios│  │  Alunos  │
  │Pendentes │  │          │  │  NF-e    │  │ Impostos │  │          │
  └────┬─────┘  └────┬─────┘  └────┬─────┘  └────┬─────┘  └────┬─────┘
       │              │             │              │             │
       └──────────────┴──────┬──────┴──────────────┘             │
                             │                                   │
                    ┌────────▼────────┐                ┌─────────▼────────┐
                    │   update_data   │                │  import manual   │
                    │    .py (cron)   │                │  (build_alunos)  │
                    │  2x/dia + full  │                │                  │
                    └────────┬────────┘                └─────────┬────────┘
                             │                                   │
                    ┌────────▼───────────────────────────────────▼┐
                    │         NEON POSTGRESQL (financeiro)         │
                    │                                             │
                    │  vindi_bills (7.652)    omie_nfse (8.408)  │
                    │  vindi_customers (3.473) omie_nfe (1.591)  │
                    │  omie_contas_pagar (2.669)                  │
                    │  omie_clientes (3.212)                      │
                    │  bma_contas_pagar (1.232)                   │
                    │  bma_contas_receber (100)                   │
                    │  pagarme_transacoes (200)                   │
                    │  vendas_planilha (3.328)                    │
                    │  veiculos_flag (64)                         │
                    │  dashboard_snapshot (33)                    │
                    │  update_log                                 │
                    └────────┬────────────────────────────────────┘
                             │
                    ┌────────▼────────┐
                    │  Geracao de JS  │
                    │  (SQL queries)  │
                    └────────┬────────┘
                             │
          ┌──────────────────┼──────────────────────┐
          │                  │                       │
   ┌──────▼──────┐   ┌──────▼──────┐        ┌──────▼──────┐
   │   Vercel    │   │   Vercel    │        │   GitHub    │
   │ Financeiro  │   │   LP-i10   │        │   Actions   │
   │  -Better    │   │ (i10 live) │        │   (cron)    │
   └─────────────┘   └─────────────┘        └─────────────┘
```

---

## 2. FONTES DE DADOS E APIs

### 2.1 Vindi (Receitas + Clientes)
- **URL**: `https://app.vindi.com.br/api/v1/`
- **Auth**: Basic (base64 de KEY:)
- **Endpoints usados**:
  - `bills?status:paid` → Receitas (faturas pagas)
  - `bills?status:pending` → Pendentes (inadimplencia)
  - `customers` → Base de clientes
- **Regra incremental**: busca bills com `created_at >= max_date_no_banco`
- **Cache**: customers skip se <7 dias
- **Registros**: ~7.652 bills, 3.473 customers

### 2.2 Pagar.me (Estornos)
- **URL**: `https://api.pagar.me/1/transactions`
- **Auth**: api_key no query string
- **Status buscados**: `refunded` (183) + `chargedback` (17)
- **Regra**: full fetch sempre (volume baixo)
- **Dedup**: 1 duplicata removida (Susy Santos R$240 — mesmo CPF+valor em refund e chargeback)

### 2.3 Omie Better (Despesas + Notas)
- **URL**: `https://app.omie.com.br/api/v1/`
- **Auth**: app_key + app_secret no body JSON
- **APP_KEY**: 4340156993172
- **Endpoints**:
  - `financas/contapagar/` ListarContasPagar → Contas a Pagar
  - `geral/clientes/` ListarClientes → Fornecedores
  - `servicos/os/` ListarOS → NFS-e (Notas Servico)
  - `produtos/pedido/` ListarPedidos → NF-e (Notas Produto)
- **Regra incremental**: ultimas N paginas (CP=10, NFS-e=5, NF-e=3)
- **Cache**: clientes skip se <7 dias
- **Rate limit**: delay 3s entre paginas, retry 3x com backoff

### 2.4 BMA Omie (Folha/Beneficios/Impostos)
- **Mesma API Omie**, credenciais diferentes
- **APP_KEY**: 4602985397010
- **Endpoints**: contapagar (1.232), contareceber (100), clientes (86)
- **Regra incremental**: CP ultimas 5 paginas, CR full (2 paginas), clientes cache 7 dias
- **Relacao com Better**: BMA e o backoffice/RH da Better. Emite NFs mensais (~R$200-350k) para a Better, mas essas NFs NAO estao lancadas no Omie Better (flag "pending NF")

### 2.5 Planilha de Vendas (Import manual)
- **Arquivo**: `Vendas para Emissao de NF (5).xlsx`
- **Aba "Vendas"**: 3.328 registros com CPF, nome, email, produto, valor, renovacao, cancelamento
- **Aba "Alunos"**: 3.569 registros com inicio/fim, nivel
- **Aba "Cancelamentos"**: 1.928 registros
- **Import**: via `build_alunos_v2.py` (manual)

---

## 3. BANCO DE DADOS NEON POSTGRESQL

### 3.1 Conexao
```
Host: ep-snowy-shadow-a4hoyxtl-pooler.us-east-1.aws.neon.tech
Database: financeiro
User: neondb_owner
SSL: required
Mesmo projeto Neon do Alumni (sem custo extra)
```

### 3.2 Tabelas

| Tabela | Registros | Chave Unica | Descricao |
|--------|-----------|-------------|-----------|
| vindi_bills | 7.652 | vindi_id | Faturas pagas Vindi |
| vindi_customers | 3.473 | vindi_id | Clientes Vindi |
| omie_contas_pagar | 2.669 | omie_id | Despesas Better |
| omie_clientes | 3.212 | omie_cod | Fornecedores Better |
| omie_nfse | 8.408 | omie_id | NFS-e (servico) |
| omie_nfe | 1.591 | omie_id | NF-e (produto) |
| bma_contas_pagar | 1.232 | omie_id | Despesas BMA |
| bma_contas_receber | 100 | omie_id | Recebiveis BMA |
| pagarme_transacoes | 200 | tid | Estornos + chargebacks |
| vendas_planilha | 3.328 | serial id | Vendas (import manual) |
| veiculos_flag | 64 | omie_id | Flag de despesas veiculos |
| dashboard_snapshot | 33 | serial id | Snapshot para changelog |
| update_log | - | serial id | Log de atualizacoes |

### 3.3 Regras de Upsert
- `ON CONFLICT (chave) DO UPDATE SET campos_que_mudam`
- Batch de 200 registros com reconexao automatica (Neon pooler dropa SSL em ops longas)
- 3 retries por batch com delay de 3s

---

## 4. REGRAS DE NEGOCIO

### 4.1 Despesas Better (Coluna "Despesas" no dashboard)
- **Fonte**: `omie_contas_pagar` WHERE fonte='Better'
- **Exclusoes**:
  - `status_titulo = 'CANCELADO'` → excluido
  - `omie_id IN veiculos_flag` → movido para pagina confidencial
  - Flash App → removido (ja esta na BMA)
- **Data de referencia**: `data_previsao` (data do pagamento), com fallback para `data_vencimento`, depois `data_emissao`
- **Motivo**: data_emissao e quando o documento foi criado, data_previsao e quando foi efetivamente pago. Ex: Learning Lab emitido em marco mas pago em abril

### 4.2 BMA (Coluna "BMA" no dashboard)
- **Fonte**: `bma_contas_pagar` WHERE status != 'CANCELADO'
- **Status**: SOMENTE CONSULTA, nao soma no resultado
- **Motivo**: BMA emite NFs para Better (~R$200-350k/mes), mas essas NFs nao estao lancadas no Omie Better. Quando forem lancadas, BMA passa a ser so drill-down
- **Flag**: "NF Pendente" em vermelho no dashboard

### 4.3 Veiculos (Pagina confidencial)
- **Fornecedores flagged**: Grand Point, Porto Seguro Capitalizacao, IPVA, PSMAXX, Sem Parar, Licenciamento, Visual Inspecao, Senatran
- **Categorias flagged**: 2.07.02, 2.06.93, 2.04.90, 2.11.96, 2.11.99, 2.04.95
- **Total**: R$ 2.55M (64 registros)
- **Senha**: Better@Frota2026
- **Regra**: removidos das despesas operacionais, visivel apenas na pagina confidencial

### 4.4 Estornos (Coluna "Estornos" no dashboard)
- **Fontes**: Pagar.me refunded (183) + Pagar.me chargedback (16)
- **Dedup**: 1 removida (mesmo CPF + mesmo valor + <60 dias entre refund e chargeback)
- **Total**: 199 registros, R$ 532k
- **Calculo resultado**: Receitas - Despesas - Estornos

### 4.5 Receitas (Coluna "Receitas" no dashboard)
- **Fonte**: `vindi_bills` WHERE status='paid'
- **Agrupamento**: por `created_at` (mes/ano)
- **Recorrencia vs Compra**: faturas mensais Vindi sao parcelas do mesmo contrato (recorrencia), NAO compras separadas. Compra real = entrada na planilha de vendas

### 4.6 NFS-e e NF-e (Regra 50/50)
- **NFS-e**: Notas de Servico (ISS ~2% Cajamar)
- **NF-e**: Notas de Produto (material didatico, ICMS diferido = isento)
- **Regra**: vendas CPF devem ser 50% servico + 50% produto
- **Excecoes**: B2B (CNPJ) e Aulas Particulares = 100% servico OK
- **Status atual**: ratio 1.88x (deveria ser 1.0x), ISS pago a mais ~R$ 46k
- **Recomendacao**: corrigir split em todas as novas vendas CPF

### 4.7 Pendentes (Coluna "Pendentes" no dashboard)
- **Fonte**: Vindi bills status=pending (fetch sempre fresh, nao armazena)
- **Significado**: faturas geradas aguardando pagamento do aluno
- **Tipos**: boletos nao pagos, PIX nao confirmado, cartao pendente
- **Inadimplencia**: pendentes >90 dias = inadimplencia real

### 4.8 Professores PJ
- **Categorias**: 2.16.xx (prestadores PJ), 2.09.01 (instrutores)
- **Portal Alumni**: 18 ativos, 7 inativos (ainda recebendo ~R$61k em 2026)
- **Alerta**: 7 inativos com 0 aulas mas pagamentos ativos

### 4.9 Alunos (Base unificada)
- **Fontes**: Planilha vendas (3.328) + Vindi bills (2.852 clientes)
- **Total unificado**: 3.060 alunos
- **Match**: 2.638 cruzados (86%)
- **Classificacao**:
  - Status: Ativo / Cancelado / Expirado / Ativo (Vindi) / Inativo
  - Modalidade: Community (2.249) / Flow (518) / Private (173) / Espanhol (68) / In-Company (35) / Imersao (17)
- **Renovacao real**: CPF com 2+ entradas na planilha de vendas (401 alunos)

---

## 5. CRON E ATUALIZACAO

### 5.1 Schedule

| Cron | Horario BRT | Modo | Duracao |
|------|-------------|------|---------|
| `0 16 * * *` | 13h diario | Incremental | ~3-5min |
| `0 22 * * *` | 19h diario | Incremental | ~3-5min |
| `0 14 1,11,21 * *` | 11h dias 1/11/21 | Full Refresh | ~40min |
| Manual (workflow_dispatch) | Qualquer hora | Escolha | Variavel |

### 5.2 Modo Incremental
- Vindi bills: so novos (MAX created_at)
- Omie CP: ultimas 10 paginas
- NFS-e: ultimas 5 paginas
- NF-e: ultimas 3 paginas
- BMA CP: ultimas 5 paginas
- Clientes: skip se <7 dias (carrega do banco)

### 5.3 Modo Full Refresh
- Ignora cache e recent_pages
- Puxa TUDO de todas as APIs
- Compara com snapshot anterior
- Detecta alteracoes retroativas

### 5.4 Fluxo pos-fetch
1. Upsert dados no Neon PostgreSQL
2. Gera JS files via SQL queries
3. Gera changelog (compara com snapshot anterior)
4. Salva novo snapshot
5. Git commit + push (auto-deploy Vercel)

---

## 6. ARQUIVOS JS GERADOS

| Arquivo | Conteudo | Gerado por |
|---------|----------|------------|
| dashboard_data.js | dashData: agregacao mensal | update_data.py (cron) |
| despesas_all_data.js | data: Better + BMA despesas | update_data.py (cron) |
| bma_data.js | bmaDesp, bmaRec, bmaAll | update_data.py (cron) |
| receitas_all_data.js | data: Vindi bills | update_data.py (cron) |
| veiculos_data.js | veiculosData: despesas veiculos | update_data.py (cron) |
| changelog.js | changelog: diferencas entre updates | update_data.py (cron) |
| alunos_data.js | alunosData, alunosStats | build_alunos_v2.py (manual) |
| receitas_customers.js | customerData, renewalStats | generate_customers.py (manual) |
| dashboard_flags.js | dashFlags: despesas atipicas | estatico |
| dashboard_diffs.js | dashDiffs: ERP vs API | estatico |
| dashboard_divergences.js | dashDivergences: conciliacao | estatico |

---

## 7. PAGINAS DO DASHBOARD

| Pagina | URL | Acento | Conteudo |
|--------|-----|--------|----------|
| Landing | index.html | Azul escuro | Hub central com tiles |
| Dashboard | dashboard.html | Azul | Tabela mensal, grafico, changelog |
| Receitas | receitas.html | Verde | Bills Vindi, clicavel, renovacoes |
| Despesas | despesas.html | Vermelho | Better + BMA com badges |
| BMA | bma.html | Roxo | Folha/beneficios/impostos |
| Fiscal | fiscal.html | Laranja | 50/50, DRE simplificado |
| Contabilidade | contabilidade.html | Teal | DRE completo, balancete, auditoria |
| Alunos | alunos.html | Ciano | Base 3.060, importacao portal |
| Veiculos | veiculos.html | Vermelho | Confidencial (senha) |
| Notas | notas.html | Amber | Produto vs Servico |
| Vendas | vendas.html | Cyan | Comissao vs Plataformas |
| Turnaround | turnaround.html | Vermelho | Reestruturacao |
| BizPlan | bizplan.html | Verde | Projecao 12 meses |

### 7.1 Domínios
- **Vercel direto**: financeiro-better.vercel.app
- **Dominio i10**: institutoi10.com.br/better-financeiro (com senha Better@2702026)
- **Deploy**: automatico via git push (GitHub → Vercel)
- **Sync**: LP-i10 repo espelha os arquivos de better-financeiro/

---

## 8. FINDINGS DE AUDITORIA (Big 4 Style)

| # | Finding | Severidade | Valor | Recomendacao |
|---|---------|-----------|-------|--------------|
| 1 | NFs BMA nao lancadas no Omie Better | ALTA | R$ 3.73M | Lancar todas NFs mensais |
| 2 | Desequilibrio fiscal 50/50 | ALTA | R$ 46k ISS a mais | Corrigir split novas vendas |
| 3 | Veiculos como despesa operacional | MEDIA | R$ 2.55M | Classificar como ativo/investimento |
| 4 | Inadimplencia Vindi | MEDIA | R$ 894k pendentes | Regua de cobranca automatica |
| 5 | Professores inativos pagos | MEDIA | R$ 61k/2026 | Encerrar contratos |

---

## 9. GITHUB SECRETS

| Secret | Uso |
|--------|-----|
| DATABASE_URL | Neon PostgreSQL |
| VINDI_KEY | API Vindi |
| PAGARME_KEY | API Pagar.me |
| OMIE_APP_KEY | API Omie Better |
| OMIE_APP_SECRET | API Omie Better |
| BMA_KEY | API Omie BMA |
| BMA_SECRET | API Omie BMA |

---

## 10. SCRIPTS DE MANUTENCAO

| Script | Funcao | Quando usar |
|--------|--------|-------------|
| update_data.py | Cron principal (incremental/full) | Automatico 2x/dia |
| migrate_to_neon.py | Migracao inicial para o banco | Uma vez (ja executado) |
| build_alunos_v2.py | Gera base de alunos | Quando atualizar planilha |
| generate_customers.py | Cruza Vindi x planilha | Quando atualizar planilha |
| fetch_bma.py | Fetch manual dados BMA | Descontinuado (integrado no cron) |
| bma_analise.py | Gera Excel analise BMA | Sob demanda |

---

## 11. FORMULAS DO DASHBOARD

```
RECEITA BRUTA        = SUM(vindi_bills.amount) WHERE status=paid
ESTORNOS             = SUM(pagarme.amount) WHERE refunded + chargedback - duplicatas
RECEITA LIQUIDA      = RECEITA BRUTA - ESTORNOS
DESPESAS BETTER      = SUM(omie_cp.valor) WHERE NOT cancelado AND NOT veiculo
BMA (consulta)       = SUM(bma_cp.valor) WHERE NOT cancelado
RESULTADO            = RECEITA LIQUIDA - DESPESAS BETTER
NFS-e                = SUM(omie_nfse.valor) WHERE NOT cancelada
NF-e                 = SUM(omie_nfe.valor) WHERE NOT cancelado
PENDENTES            = SUM(vindi_pending.amount) — fetch fresh, nao armazena
```

---

*Documento gerado automaticamente. Manter atualizado conforme evolucao do sistema.*
