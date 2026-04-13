import pandas as pd
import json
import os
import sys
import time
from datetime import datetime

# ═══════════════════════════════════════════════════════════
# CONFIGURAÇÃO DO ROBÔ DE APONTAMENTOS
# ═══════════════════════════════════════════════════════════
PLANILHA      = "Check_Apontamento_x_Inventario_v2.xlsx"
SAIDA_HTML    = "dash_apontamentos.html"  # Nome único para não conflitar!
INTERVALO_SEG = 5                         # verifica alteracoes a cada 5 segundos
ABRIR_BROWSER = True                      # abre no PC assim que gerar
# ═══════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════
# LEITURA DOS DADOS
# ═══════════════════════════════════════════════════════════
def carregar_dados(caminho):
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    print(f"  📂 Lendo planilha: {caminho}")
    df_ap = pd.read_excel(caminho, sheet_name='Apontamentos')
    df_inv = pd.read_excel(caminho, sheet_name='Inventario')
    df_listas = pd.read_excel(caminho, sheet_name='_Listas')

    df_ap['DATA'] = pd.to_datetime(df_ap['DATA'], errors='coerce')
    # Melhoria para o Outline: manter a data como string formatada
    df_ap['DATA_STR'] = df_ap['DATA'].dt.strftime('%d/%m/%Y').fillna("Data N/D")
    df_ap['ANO'] = df_ap['DATA'].dt.year.fillna(0).astype(int)

    print(f"     ✅ Apontamentos: {len(df_ap):,} linhas")
    print(f"     ✅ Inventário:   {len(df_inv):,} linhas")

    # Talhões únicos do Inventário
    talhoes = df_inv.groupby(['FAZENDA', 'ZONA', 'TALHAO']).agg(
        AREA=('AREA_HA', 'sum'),
        VARIEDADE=('VARIED', 'first')
    ).reset_index()

    # Apontamentos agregados por Fazenda/Zona/Talhão/CodOper/Ano
    agg = df_ap.groupby(['DE_UPNIVEL1', 'ZONA', 'TALHAO', 'COD_OPER', 'ANO']).agg(
        SOMA_APONTADA=('AREA', 'sum'),
        QTD_REGISTROS=('Chavesig', 'count')
    ).reset_index()
    agg.columns = ['FAZENDA', 'ZONA', 'TALHAO', 'COD_OPER', 'ANO',
                   'SOMA_APONTADA', 'QTD_REGISTROS']

    # INCLUSÃO: Detalhamento para o Outline (necessário para ver as datas dos registros)
    detalhes = df_ap[['DE_UPNIVEL1', 'ZONA', 'TALHAO', 'COD_OPER', 'ANO', 'DATA_STR', 'AREA']].copy()

    # Lista de operações (merge _Listas + dados reais)
    ops_listas = df_listas[['COD_OPER', 'OPERACAO']].dropna(subset=['COD_OPER']).drop_duplicates()
    ops_dados = df_ap[['COD_OPER', 'OPERACAO']].drop_duplicates()
    ops = pd.concat([ops_listas, ops_dados]).drop_duplicates(
        subset=['COD_OPER']).sort_values('COD_OPER')

    anos = sorted([int(a) for a in df_ap['ANO'].unique() if a > 0])
    fazendas = sorted(df_inv['FAZENDA'].unique().tolist())

    # Arredondar
    talhoes['AREA'] = talhoes['AREA'].round(2)
    agg['SOMA_APONTADA'] = agg['SOMA_APONTADA'].round(2)

    return {
        'talhoes': talhoes.to_dict('records'),
        'apontamentos': agg.to_dict('records'),
        'detalhes': detalhes.to_dict('records'), # Inclusão da base para o Outline
        'operacoes': ops.to_dict('records'),
        'anos': anos,
        'fazendas': fazendas
    }

# ═══════════════════════════════════════════════════════════
# TEMPLATE HTML — TEMA RAÍZEN CORPORATIVO (FUNDO CINZA + VERDE)
# ═══════════════════════════════════════════════════════════
def gerar_html(data, data_geracao):
    js_data = json.dumps(data, ensure_ascii=False, separators=(',', ':'),
                         default=lambda x: int(x) if hasattr(x, 'item') else str(x)).replace("</", "<\\/")

    html = f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Raízen — Painel de Auditoria Agrícola</title>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
:root {{
  /* Raízen Palette - Oficial e Corporativa */
  --rz-purple-main: #692096;
  --rz-purple-light: #8b31bc;
  --rz-purple-glow: rgba(105, 32, 150, 0.1);

  --rz-green-main: #00A84E; /* Verde Oficial Raízen */
  --rz-green-dark: #00843D;
  --rz-green-glow: rgba(0, 168, 78, 0.15);

  /* Tema Executivo (Fundo Cinza e Cards Brancos) */
  --bg: #e2e5e9; /* Cinza mais preenchido para contraste */
  --surface: #ffffff;
  --surface2: #f8f9fa;
  --surface3: #e9ecef;
  --border: #ced4da;
  --border2: #adb5bd;

  /* Textos Escuros */
  --text: #18181b;
  --text2: #3f3f46;
  --text3: #6c757d;

  /* Status Alerts */
  --amber: #d97706; --amber-bg: rgba(245,158,11,.15); --amber-border: rgba(245,158,11,.4);
  --red: #dc2626; --red-bg: rgba(239,68,68,.1); --red-border: rgba(239,68,68,.3);

  --radius: 10px; --radius-lg: 14px;
}}
body{{font-family:'Plus Jakarta Sans',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden}}
body::before{{content:'';position:fixed;inset:0;opacity:.025;pointer-events:none;
background:url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='.8' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)'/%3E%3C/svg%3E");z-index:9999}}
body::after{{content:'';position:fixed;top:-40%;right:-20%;width:80vw;height:80vw;
background:radial-gradient(circle, rgba(0, 168, 78, 0.08) 0%, transparent 60%);pointer-events:none;z-index:0}}

.container{{max-width:1480px;margin:0 auto;padding:24px 32px;position:relative;z-index:1}}

/* Header */
.header{{display:flex;align-items:center;justify-content:space-between;margin-bottom:28px;padding-bottom:20px;border-bottom:1px solid var(--border)}}
.header-left{{display:flex;align-items:center;gap:16px}}
.logo-mark{{width:46px;height:46px;background:linear-gradient(135deg, var(--rz-green-main), var(--rz-purple-main));border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;box-shadow:0 4px 20px var(--rz-green-glow),0 0 40px var(--rz-purple-glow)}}
.header h1{{font-size:22px;font-weight:800;letter-spacing:-.4px}}
.header h1 span{{background:linear-gradient(135deg, var(--rz-green-main), var(--rz-purple-main));-webkit-background-clip:text;-webkit-text-fill-color:transparent}}
.header-sub{{font-size:11px;color:var(--text3);font-weight:600;letter-spacing:1.2px;text-transform:uppercase;margin-top:3px}}
.header-right{{display:flex;align-items:center;gap:12px}}
.header-badge{{font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--text2);background:var(--surface);padding:8px 14px;border-radius:8px;border:1px solid var(--border)}}

/* Filters */
.filters{{display:flex;gap:14px;margin-bottom:24px;flex-wrap:wrap;align-items:flex-end}}
.filter-group{{display:flex;flex-direction:column;gap:6px;flex:1;min-width:160px;}}
.filter-group label{{font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:1px}}
.filter-group select,.filter-group input{{
background:var(--surface);border:1px solid var(--border);color:var(--text);
padding:10px 14px;border-radius:var(--radius);font-family:'Plus Jakarta Sans',sans-serif;
font-size:13px;width:100%;cursor:pointer;transition:all .2s;outline:none;
appearance:none;-webkit-appearance:none;
background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%238a789c' fill='none' stroke-width='1.5' stroke-linecap='round'/%3E%3C/svg%3E");
background-repeat:no-repeat;background-position:right 12px center;padding-right:36px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);}}
.filter-group select:focus,.filter-group input:focus{{border-color:var(--rz-green-main);box-shadow:0 0 0 3px var(--rz-green-glow)}}
.filter-group input{{background-image:none;padding-right:14px;cursor:text;}}

/* KPIs */
.kpis{{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin-bottom:24px}}
.kpi{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-lg);padding:18px 20px;position:relative;overflow:hidden;transition:all .3s ease; cursor: pointer; box-shadow: 0 2px 8px rgba(0,0,0,0.04);}}
.kpi:hover{{border-color:var(--border2);transform:translateY(-3px);box-shadow:0 8px 24px rgba(0,0,0,.08)}}
.kpi.active{{border-color:var(--rz-green-main);transform:translateY(-2px);box-shadow:0 8px 24px rgba(0, 168, 78, .15); outline: 2px solid var(--rz-green-main); outline-offset: -1px;}}
.kpi::before{{content:'';position:absolute;top:0;left:0;right:0;height:4px}}
.kpi.total::before{{background:linear-gradient(90deg, var(--rz-purple-main), var(--rz-purple-light))}}
.kpi.done::before{{background:var(--rz-green-main); box-shadow: 0 0 10px var(--rz-green-glow)}}
.kpi.partial::before{{background:var(--amber)}}
.kpi.alert::before{{background:var(--red)}}
.kpi.pending::before{{background:var(--text3)}}
.kpi-value{{font-family:'JetBrains Mono',monospace;font-size:30px;font-weight:700;margin-bottom:4px}}
.kpi.total .kpi-value{{color:var(--rz-purple-main)}}
.kpi.done .kpi-value{{color:var(--rz-green-main); text-shadow: 0 0 15px rgba(0, 168, 78, 0.2);}}
.kpi.partial .kpi-value{{color:var(--amber)}}
.kpi.alert .kpi-value{{color:var(--red)}}
.kpi.pending .kpi-value{{color:var(--text3)}}
.kpi-label{{font-size:12px;color:var(--text2);font-weight:600}}
.kpi-pct{{position:absolute;top:16px;right:18px;font-family:'JetBrains Mono',monospace;font-size:11px;padding:3px 8px;border-radius:6px; font-weight: 700;}}
.kpi.done .kpi-pct{{color:var(--surface);background:var(--rz-green-main)}}
.kpi.partial .kpi-pct{{color:var(--amber);background:var(--amber-bg)}}

/* Area summary */
.area-summary{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:24px}}
.area-card{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-lg);padding:16px 20px;transition:all .3s; box-shadow: 0 2px 8px rgba(0,0,0,0.04);}}
.area-card:hover{{border-color:var(--border2)}}
.area-card-label{{font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:.8px;font-weight:700;margin-bottom:6px}}
.area-card-value{{font-family:'JetBrains Mono',monospace;font-size:20px;font-weight:700}}
.area-card-unit{{font-size:12px;color:var(--text3);font-weight:400;margin-left:4px}}

/* Progress */
.progress-section{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-lg);padding:18px 22px;margin-bottom:24px;display:flex;align-items:center;gap:20px; box-shadow: 0 2px 8px rgba(0,0,0,0.04);}}
.progress-label{{font-size:13px;color:var(--text2);white-space:nowrap;font-weight:600}}
.progress-bar{{flex:1;height:14px;background:var(--surface3);border-radius:99px;overflow:hidden; box-shadow: inset 0 2px 4px rgba(0,0,0,0.1)}}
.progress-fill{{height:100%;border-radius:99px;transition:width .8s cubic-bezier(.4,0,.2,1);position:relative; box-shadow: 0 0 10px rgba(0, 168, 78, 0.4);}}
.progress-fill::after{{content:'';position:absolute;inset:0;background:linear-gradient(90deg,transparent,rgba(255,255,255,.2));border-radius:99px}}
.progress-pct{{font-family:'JetBrains Mono',monospace;font-size:16px;font-weight:700;min-width:56px;text-align:right}}

/* Table */
.table-container{{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-lg);overflow:hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.05);}}
.table-header{{display:flex;justify-content:space-between;align-items:center;padding:16px 22px;border-bottom:1px solid var(--border); background:var(--surface)}}
.table-title{{font-size:14px;font-weight:700;display:flex;align-items:center;gap:8px}}
.table-count{{font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--text2);background:var(--surface3);padding:3px 10px;border-radius:6px; border: 1px solid var(--border2)}}
.table-scroll{{overflow-x:auto;max-height:540px;overflow-y:auto}}
.table-scroll::-webkit-scrollbar{{width:6px;height:6px}}
.table-scroll::-webkit-scrollbar-track{{background:transparent}}
.table-scroll::-webkit-scrollbar-thumb{{background:var(--border2);border-radius:99px}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
thead{{position:sticky;top:0;z-index:10}}
th{{background:var(--surface2);color:var(--text2);font-weight:700;font-size:10px;text-transform:uppercase;letter-spacing:.7px;padding:14px 16px;text-align:left;white-space:nowrap;border-bottom:1px solid var(--border);cursor:pointer;user-select:none;transition:color .2s}}
th:hover{{color:var(--rz-green-main)}}
th.sorted-asc::after{{content:' ▲';font-size:8px;color:var(--rz-green-main)}}
th.sorted-desc::after{{content:' ▼';font-size:8px;color:var(--rz-green-main)}}
td{{padding:12px 16px;border-bottom:1px solid var(--surface3);white-space:nowrap;transition:background .15s}}
tr:nth-child(even) td {{ background: var(--surface2); }}
tr:hover td{{background:var(--rz-green-glow)}}
.mono{{font-family:'JetBrains Mono',monospace;font-size:12px}}
.right{{text-align:right}}

/* Badges */
.badge{{display:inline-flex;align-items:center;gap:5px;padding:5px 10px;border-radius:6px;font-size:11px;font-weight:700;white-space:nowrap; letter-spacing: 0.3px;}}
.badge.concluido{{color:var(--surface);background:var(--rz-green-main);border:1px solid var(--rz-green-dark); box-shadow: 0 0 8px var(--rz-green-glow);}}
.badge.parcial{{color:var(--amber);background:var(--amber-bg);border:1px solid var(--amber-border)}}
.badge.atencao{{color:var(--surface);background:var(--red);border:1px solid var(--red-border); box-shadow: 0 0 8px var(--red-bg);}}
.badge.nao-iniciado{{color:var(--text2);background:var(--surface3);border:1px solid var(--border2)}}
.badge.sem-area{{color:var(--text3);background:transparent;border:1px dashed var(--border)}}

/* Mini progress */
.mini-bar{{width:56px;height:4px;background:var(--surface3);border-radius:99px;overflow:hidden;display:inline-block;vertical-align:middle;margin-right:8px}}
.mini-bar-fill{{height:100%;border-radius:99px; box-shadow: 0 0 5px rgba(0, 168, 78, 0.5);}}

/* Pagination */
.pagination{{display:flex;justify-content:space-between;align-items:center;padding:14px 22px;border-top:1px solid var(--border); background:var(--surface)}}
.pagination-info{{font-size:12px;color:var(--text3)}}
.pagination-btns{{display:flex;gap:6px}}
.pagination-btns button{{background:var(--surface);border:1px solid var(--border);color:var(--text2);padding:6px 12px;border-radius:6px;font-size:12px;cursor:pointer;font-family:'Plus Jakarta Sans',sans-serif;transition:all .2s; font-weight: 600; box-shadow: 0 1px 2px rgba(0,0,0,0.05);}}
.pagination-btns button:hover:not(:disabled){{background:var(--surface3);color:var(--text);border-color:var(--rz-green-main)}}
.pagination-btns button:disabled{{opacity:.4;cursor:not-allowed;box-shadow:none}}
.pagination-btns button.active{{background:var(--rz-green-main);border-color:var(--rz-green-dark);color:#fff; box-shadow: 0 0 10px var(--rz-green-glow);}}

/* Export */
.btn-export{{background:var(--surface);border:1px solid var(--border);color:var(--text2);padding:10px 18px;border-radius:var(--radius);font-family:'Plus Jakarta Sans',sans-serif;font-size:11px;font-weight:700;cursor:pointer;transition:all .2s;display:flex;align-items:center;gap:6px;text-transform:uppercase;letter-spacing:.6px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);}}
.btn-export:hover{{background:var(--rz-green-main);color:#fff;border-color:var(--rz-green-dark); box-shadow: 0 0 12px var(--rz-green-glow);}}

/* INCLUSÃO: Estilos do Outline (Melhoria Solicitada) */
.row-main {{ cursor: pointer; transition: background 0.1s; }}
.row-main:hover {{ background: var(--rz-green-glow) !important; }}
.row-detail {{ display: none; background: #fafafa !important; }}
.row-detail td {{ padding: 15px 40px !important; border-left: 5px solid var(--rz-green-main); }}
.expander {{ display: inline-block; width: 16px; transition: transform 0.2s; color: var(--rz-green-main); font-weight: 800; }}
.expanded .expander {{ transform: rotate(90deg); }}
.det-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); gap: 8px; }}
.det-item {{ background: #fff; border: 1px solid #ddd; padding: 6px 10px; border-radius: 6px; font-size: 11px; }}

.footer{{text-align:center;padding:24px;color:var(--text3);font-size:11px;letter-spacing:.3px}}
.footer span{{color:var(--rz-green-main);font-weight:800}}

@media(max-width:900px){{
.kpis{{grid-template-columns:repeat(2,1fr)}}
.area-summary{{grid-template-columns:1fr}}
.filters{{flex-direction:column}}
.container{{padding:16px}}
}}
</style>
</head>
<body>
<div class="container" id="app"><div style="text-align:center;padding:80px;color:var(--text3)">Carregando dados...</div></div>

<script>
const DATA={js_data};
const GENERATED="{data_geracao}";
const PER_PAGE=50;

// Estado global dos Filtros
let S={{
    codOper:DATA.operacoes.length > 0 ? DATA.operacoes[0].COD_OPER : 789,
    ano:new Date().getFullYear(),
    search:'',
    zona:'all',
    talhao:'',
    sortCol:null,
    sortDir:'asc',
    page:0,
    statusF:'all'
}};

let focusId = null; // Rastreador de qual input está focado

// Função Toggle para o Outline (Melhoria)
function toggleRow(id) {{
    const el = document.getElementById('det-' + id);
    const main = document.getElementById('main-' + id);
    if (el.style.display === 'table-row') {{
        el.style.display = 'none';
        main.classList.remove('expanded');
    }} else {{
        el.style.display = 'table-row';
        main.classList.add('expanded');
    }}
}}

function calc(){{
  const{{codOper,ano,search,zona,talhao,statusF}}=S;
  const m={{}};
  DATA.apontamentos.forEach(a=>{{
    if(a.COD_OPER===codOper&&a.ANO===ano){{
      m[a.FAZENDA+'|'+a.ZONA+'|'+a.TALHAO]={{s:a.SOMA_APONTADA,q:a.QTD_REGISTROS}};
    }}
  }});
  
  let rows=DATA.talhoes.map(t=>{{
    const k=t.FAZENDA+'|'+t.ZONA+'|'+t.TALHAO;
    const ap=m[k]||{{s:0,q:0}};
    const d=ap.s-t.AREA,p=t.AREA>0?ap.s/t.AREA:0;
    let st;
    if(t.AREA===0)st='Sem Área';
    else if(ap.s===0)st='Não Iniciado';
    else if(Math.abs(d)<0.01)st='Concluído';
    else if(ap.s<t.AREA)st='Parcial';
    else st='Atenção';
    return{{...t,SOMA:ap.s,DIFF:d,PCT:p,STATUS:st,QTD:ap.q}};
  }});
  
  // Aplicação dos Múltiplos Filtros
  if(search){{
      const s=search.toLowerCase();
      rows=rows.filter(r=>r.FAZENDA.toLowerCase().includes(s));
  }}
  if(zona !== 'all'){{
      rows=rows.filter(r=>String(r.ZONA)===String(zona));
  }}
  if(talhao){{
      const t=talhao.toLowerCase();
      rows=rows.filter(r=>String(r.TALHAO).toLowerCase().includes(t));
  }}
  if(statusF!=='all'){{
      rows=rows.filter(r=>r.STATUS===statusF);
  }}
  
  return rows;
}}

function srt(rows){{
  if(!S.sortCol)return rows;
  const c=S.sortCol,d=S.sortDir==='asc'?1:-1;
  return[...rows].sort((a,b)=>{{let va=a[c],vb=b[c];if(typeof va==='string')return va.localeCompare(vb)*d;return((va||0)-(vb||0))*d;}});
}}

function opNm(c){{const o=DATA.operacoes.find(x=>x.COD_OPER===c);return o?o.OPERACAO:'';}}
function sc(s){{if(s==='Concluído')return'concluido';if(s==='Parcial')return'parcial';if(s==='Atenção')return'atencao';if(s==='Não Iniciado')return'nao-iniciado';return'sem-area';}}
function pc(p){{if(p>=1)return'var(--rz-green-main)';if(p>=.5)return'var(--amber)';if(p>0)return'var(--red)';return'var(--text3)';}}
function fm(n){{return n.toLocaleString('pt-BR',{{minimumFractionDigits:2,maximumFractionDigits:2}});}}
function pf(n){{return(n*100).toFixed(1)+'%';}}

function setStatusF(st) {{
    S.statusF = st;
    S.page = 0;
    render();
}}

function render(){{
  const rows=calc(),sorted=srt(rows);
  const tp=Math.ceil(sorted.length/PER_PAGE);
  if(S.page>=tp)S.page=Math.max(0,tp-1);
  const pg=sorted.slice(S.page*PER_PAGE,(S.page+1)*PER_PAGE);
  const tot=rows.length,dn=rows.filter(r=>r.STATUS==='Concluído').length;
  const pa=rows.filter(r=>r.STATUS==='Parcial').length;
  const al=rows.filter(r=>r.STATUS==='Atenção').length;
  const ni=rows.filter(r=>r.STATUS==='Não Iniciado').length;
  const gp=tot>0?dn/tot:0;
  const tA=rows.reduce((s,r)=>s+r.AREA,0),tAp=rows.reduce((s,r)=>s+r.SOMA,0);
  const aP=tA>0?tAp/tA:0;
  
  const ocs=[...new Set(DATA.apontamentos.map(a=>a.COD_OPER))].sort((a,b)=>a-b);
  // Pega todas as zonas únicas de forma limpa
  const zonas=[...new Set(DATA.talhoes.map(t=>t.ZONA))].sort((a,b)=>String(a).localeCompare(String(b), undefined, {{numeric: true}}));
  
  const si=c=>S.sortCol===c?(S.sortDir==='asc'?' sorted-asc':' sorted-desc'):'';

  document.getElementById('app').innerHTML=`
    <div class="header">
      <div class="header-left">
        <div class="logo-mark">🌾</div>
        <div>
          <h1><span>Raízen</span> — Auditoria Agrícola</h1>
          <div class="header-sub">Apontamento vs Inventário</div>
        </div>
      </div>
      <div class="header-right">
        <div class="header-badge">Gerado em: ${{GENERATED}}</div>
      </div>
    </div>

    <div class="filters">
      <div class="filter-group">
        <label>Operação</label>
        <select id="fOp">
          ${{ocs.map(c=>`<option value="${{c}}" ${{c===S.codOper?'selected':''}}>${{c}} — ${{opNm(c)}}</option>`).join('')}}
        </select>
      </div>
      <div class="filter-group">
        <label>Ano</label>
        <select id="fAno">
          ${{DATA.anos.map(a=>`<option value="${{a}}" ${{a===S.ano?'selected':''}}>${{a}}</option>`).join('')}}
        </select>
      </div>
      <div class="filter-group">
        <label>Status</label>
        <select id="fSt">
          <option value="all" ${{S.statusF==='all'?'selected':''}}>Todos</option>
          <option value="Concluído" ${{S.statusF==='Concluído'?'selected':''}}>Concluído</option>
          <option value="Parcial" ${{S.statusF==='Parcial'?'selected':''}}>Parcial</option>
          <option value="Atenção" ${{S.statusF==='Atenção'?'selected':''}}>⚠ Atenção</option>
          <option value="Não Iniciado" ${{S.statusF==='Não Iniciado'?'selected':''}}>Não Iniciado</option>
        </select>
      </div>
      <div class="filter-group">
        <label>Zona</label>
        <select id="fZn">
          <option value="all" ${{S.zona==='all'?'selected':''}}>Todas as Zonas</option>
          ${{zonas.map(z=>`<option value="${{z}}" ${{String(z)===String(S.zona)?'selected':''}}>${{z}}</option>`).join('')}}
        </select>
      </div>
      <div class="filter-group">
        <label>Buscar Talhão</label>
        <input type="text" id="fTl" placeholder="Ex: 001..." value="${{S.talhao}}">
      </div>
      <div class="filter-group">
        <label>Buscar Fazenda</label>
        <input type="text" id="fSrch" placeholder="Ex: Vale Verde..." value="${{S.search}}">
      </div>
    </div>

    <div class="kpis">
      <div class="kpi total ${{S.statusF==='all'?'active':''}}" onclick="setStatusF('all')" title="Ver Todos">
        <div class="kpi-value">${{tot.toLocaleString('pt-BR')}}</div>
        <div class="kpi-label">Total de Talhões</div>
      </div>
      <div class="kpi done ${{S.statusF==='Concluído'?'active':''}}" onclick="setStatusF('Concluído')" title="Filtrar Concluídos">
        <div class="kpi-value">${{dn.toLocaleString('pt-BR')}}</div>
        <div class="kpi-label">Concluídos</div>
        ${{tot>0?`<div class="kpi-pct">${{pf(dn/tot)}}</div>`:''}}
      </div>
      <div class="kpi partial ${{S.statusF==='Parcial'?'active':''}}" onclick="setStatusF('Parcial')" title="Filtrar Parciais">
        <div class="kpi-value">${{pa.toLocaleString('pt-BR')}}</div>
        <div class="kpi-label">Parciais</div>
        ${{tot>0?`<div class="kpi-pct">${{pf(pa/tot)}}</div>`:''}}
      </div>
      <div class="kpi alert ${{S.statusF==='Atenção'?'active':''}}" onclick="setStatusF('Atenção')" title="Filtrar Atenção">
        <div class="kpi-value">${{al.toLocaleString('pt-BR')}}</div>
        <div class="kpi-label">⚠ Atenção</div>
      </div>
      <div class="kpi pending ${{S.statusF==='Não Iniciado'?'active':''}}" onclick="setStatusF('Não Iniciado')" title="Filtrar Não Iniciados">
        <div class="kpi-value">${{ni.toLocaleString('pt-BR')}}</div>
        <div class="kpi-label">Não Iniciados</div>
      </div>
    </div>

    <div class="area-summary">
      <div class="area-card">
        <div class="area-card-label">Área Total Inventário</div>
        <div class="area-card-value">${{fm(tA)}}<span class="area-card-unit">ha</span></div>
      </div>
      <div class="area-card">
        <div class="area-card-label">Área Total Apontada</div>
        <div class="area-card-value" style="color:${{aP>=1?'var(--rz-green-main)':aP>0?'var(--amber)':'var(--text3)'}}">${{fm(tAp)}}<span class="area-card-unit">ha</span></div>
      </div>
      <div class="area-card">
        <div class="area-card-label">Execução por Área</div>
        <div class="area-card-value" style="color:${{pc(aP)}}; text-shadow: ${{aP>=1?'0 0 10px rgba(0, 168, 78, 0.3)':'none'}}">${{pf(aP)}}</div>
      </div>
    </div>

    <div class="progress-section">
      <div class="progress-label">Conclusão Geral</div>
      <div class="progress-bar">
        <div class="progress-fill" style="width:${{Math.min(gp*100,100)}}%;background:${{pc(gp)}}"></div>
      </div>
      <div class="progress-pct" style="color:${{pc(gp)}}">${{pf(gp)}}</div>
    </div>

    <div class="table-container">
      <div class="table-header">
        <div class="table-title">📋 Resultado da Auditoria <span class="table-count">${{sorted.length}} talhões</span></div>
        <button class="btn-export" onclick="exp()">⬇ Exportar CSV</button>
      </div>
      <div class="table-scroll">
        <table>
          <thead><tr>
            <th class="${{si('FAZENDA')}}" onclick="ts('FAZENDA')">Fazenda</th>
            <th class="${{si('ZONA')}}" onclick="ts('ZONA')">Zona</th>
            <th class="${{si('TALHAO')}}" onclick="ts('TALHAO')">Talhão</th>
            <th class="right${{si('AREA')}}" onclick="ts('AREA')">Área (ha)</th>
            <th class="right${{si('SOMA')}}" onclick="ts('SOMA')">Apontado (ha)</th>
            <th class="right${{si('DIFF')}}" onclick="ts('DIFF')">Diferença</th>
            <th class="right${{si('PCT')}}" onclick="ts('PCT')">% Execução</th>
            <th class="${{si('STATUS')}}" onclick="ts('STATUS')">Status</th>
            <th class="right${{si('QTD')}}" onclick="ts('QTD')">Registros</th>
            <th class="${{si('VARIEDADE')}}" onclick="ts('VARIEDADE')">Variedade</th>
          </tr></thead>
          <tbody>
            ${{pg.map(r=>{{
                const rowId = `${{r.FAZENDA}}_${{r.ZONA}}_${{r.TALHAO}}`.replace(/\\s/g,'_');
                // Filtro de detalhes para o Outline
                const dets = DATA.detalhes.filter(d => 
                    d.DE_UPNIVEL1 === r.FAZENDA && 
                    String(d.ZONA) === String(r.ZONA) && 
                    String(d.TALHAO) === String(r.TALHAO) && 
                    d.COD_OPER === S.codOper && 
                    d.ANO === S.ano
                );
                
                return `
                <tr class="row-main" id="main-${{rowId}}" onclick="toggleRow('${{rowId}}')">
                  <td><span class="expander">▶</span> ${{r.FAZENDA}}</td>
                  <td class="mono">${{r.ZONA}}</td>
                  <td class="mono">${{r.TALHAO}}</td>
                  <td class="mono right">${{fm(r.AREA)}}</td>
                  <td class="mono right">${{fm(r.SOMA)}}</td>
                  <td class="mono right" style="color:${{r.DIFF>0?'var(--red)':r.DIFF<0?'var(--amber)':'var(--text3)'}}">${{r.DIFF>0?'+':''}}${{fm(r.DIFF)}}</td>
                  <td class="right"><span class="mini-bar"><span class="mini-bar-fill" style="width:${{Math.min(r.PCT*100,100)}}%;background:${{pc(r.PCT)}}"></span></span><span class="mono">${{pf(r.PCT)}}</span></td>
                  <td><span class="badge ${{sc(r.STATUS)}}">${{r.STATUS==='Atenção'?'⚠ Excedente':r.STATUS}}</span></td>
                  <td class="mono right">${{r.QTD}}</td>
                  <td style="color:var(--text2);font-size:12px">${{r.VARIEDADE||''}}</td>
                </tr>
                <tr class="row-detail" id="det-${{rowId}}">
                  <td colspan="10">
                    <div style="margin-bottom:8px"><strong>Detalhamento dos Apontamentos (Operação ${{S.codOper}}):</strong></div>
                    <div class="det-grid">
                      ${{dets.map(d => `
                        <div class="det-item">
                          📅 ${{d.DATA_STR}}<br>
                          <b>${{fm(d.AREA)}} ha</b>
                        </div>
                      `).join('')}}
                      ${{dets.length === 0 ? '<div style="color:#999;font-style:italic">Nenhum registro encontrado para este filtro.</div>' : ''}}
                    </div>
                  </td>
                </tr>`;
            }}).join('')}}
          </tbody>
        </table>
      </div>
      <div class="pagination">
        <div class="pagination-info">Mostrando ${{S.page*PER_PAGE+1}}–${{Math.min((S.page+1)*PER_PAGE,sorted.length)}} de ${{sorted.length}}</div>
        <div class="pagination-btns">
          <button onclick="gp2(0)" ${{S.page===0?'disabled':''}}>⟪</button>
          <button onclick="gp2(${{S.page-1}})" ${{S.page===0?'disabled':''}}>◂ Anterior</button>
          <button onclick="gp2(${{S.page+1}})" ${{S.page>=tp-1?'disabled':''}}>Próximo ▸</button>
          <button onclick="gp2(${{tp-1}})" ${{S.page>=tp-1?'disabled':''}}>⟫</button>
        </div>
      </div>
    </div>

    <div class="footer">
      <span>Raízen</span> — Painel de Auditoria Agrícola · ${{DATA.talhoes.length.toLocaleString('pt-BR')}} talhões · ${{DATA.apontamentos.length.toLocaleString('pt-BR')}} apontamentos
    </div>`;

  // Listeners de Eventos
  document.getElementById('fOp').onchange=e=>{{S.codOper=+e.target.value;S.page=0;focusId=null;render();}};
  document.getElementById('fAno').onchange=e=>{{S.ano=+e.target.value;S.page=0;focusId=null;render();}};
  document.getElementById('fSt').onchange=e=>{{S.statusF=e.target.value;S.page=0;focusId=null;render();}};
  document.getElementById('fZn').onchange=e=>{{S.zona=e.target.value;S.page=0;focusId=null;render();}};
  
  // Tratamento especial para Text Inputs não perderem o cursor
  document.getElementById('fTl').oninput=e=>{{S.talhao=e.target.value;S.page=0;focusId='fTl';render();}};
  document.getElementById('fSrch').oninput=e=>{{S.search=e.target.value;S.page=0;focusId='fSrch';render();}};

  // Devolve o foco pro campo de texto ativo
  if(focusId){{
      const el = document.getElementById(focusId);
      if(el){{
          el.focus(); 
          el.selectionStart = el.selectionEnd = el.value.length;
      }}
  }}
}}

function ts(c){{if(S.sortCol===c)S.sortDir=S.sortDir==='asc'?'desc':'asc';else{{S.sortCol=c;S.sortDir='asc';}}render();}}
function gp2(p){{S.page=Math.max(0,p);render();}}
function exp(){{
  const rows=srt(calc());
  let csv='\\uFEFFFAZENDA;ZONA;TALHAO;AREA_HA;SOMA_APONTADA;DIFERENCA;PCT_EXECUCAO;STATUS;QTD_REGISTROS;VARIEDADE\\n';
  rows.forEach(r=>{{csv+=`${{r.FAZENDA}};${{r.ZONA}};${{r.TALHAO}};${{r.AREA}};${{r.SOMA}};${{r.DIFF.toFixed(2)}};${{(r.PCT*100).toFixed(1)}}%;${{r.STATUS}};${{r.QTD}};${{r.VARIEDADE||''}}\\n`;}});
  const b=new Blob([csv],{{type:'text/csv;charset=utf-8'}});
  const u=URL.createObjectURL(b);
  const a=document.createElement('a');a.href=u;a.download='auditoria_raizen.csv';a.click();URL.revokeObjectURL(u);
}}
render();
</script>
</body>
</html>'''
    return html


# ═══════════════════════════════════════════════════════════
# O MOTOBOY AUTOMÁTICO (ATUALIZAÇÃO NA NUVEM)
# ═══════════════════════════════════════════════════════════
def enviar_para_github():
    print("  [☁️] Subindo nova versão para a Nuvem (GitHub)...")
    try:
        os.system(f'git add "{SAIDA_HTML}"')
        mensagem = f'Atualizacao automatica APONTAMENTOS: {time.strftime("%d/%m/%Y %H:%M:%S")}'
        os.system(f'git commit -m "{mensagem}"')
        os.system('git push')
        print(f"  [✅] Sucesso! O Netlify já está atualizando o link {SAIDA_HTML}.")
    except Exception as e:
        print(f"  [❌] Erro ao enviar para a nuvem: {e}")


# ═══════════════════════════════════════════════════════════
# EXECUÇÃO PRINCIPAL - MODO MONITORAMENTO INFINITO
# ═══════════════════════════════════════════════════════════
def monitorar():
    pasta_script = os.path.dirname(os.path.abspath(__file__))
    os.chdir(pasta_script)
    planilha_path = os.path.join(pasta_script, PLANILHA)

    print("=" * 60)
    print("  🌾 GERADOR DE DASHBOARD — RAÍZEN AUDITORIA AGRÍCOLA")
    print("  (Modo Monitoramento Automático para Nuvem)")
    print("=" * 60)

    if not os.path.exists(planilha_path):
        print(f"  [ERRO] Arquivo não encontrado: {planilha_path}")
        print("  Coloque o arquivo Excel na mesma pasta do script.")
        sys.exit(1)

    print("  Gerando dashboard de apontamentos inicial...")
    try:
        data = carregar_dados(planilha_path)
        data_geracao = datetime.now().strftime("%d/%m/%Y %H:%M")
        
        html = gerar_html(data, data_geracao)
        with open(SAIDA_HTML, 'w', encoding='utf-8') as f:
            f.write(html)
            
        enviar_para_github()
        
        if ABRIR_BROWSER:
            os.startfile(os.path.abspath(SAIDA_HTML))
    except Exception as e:
        print(f"  [ERRO] {e}")

    ultima_mod = os.path.getmtime(planilha_path)
    print("-" * 60)

    try:
        while True:
            time.sleep(INTERVALO_SEG)
            mod_atual = os.path.getmtime(planilha_path)

            if mod_atual != ultima_mod:
                ultima_mod = mod_atual
                hora = time.strftime("%H:%M:%S")
                print()
                print(f"  [{hora}] Excel de Apontamentos salvo! Gerando novos dados...")
                time.sleep(3) # Espera o Excel liberar o arquivo
                try:
                    data = carregar_dados(planilha_path)
                    data_geracao = datetime.now().strftime("%d/%m/%Y %H:%M")
                    
                    html = gerar_html(data, data_geracao)
                    with open(SAIDA_HTML, 'w', encoding='utf-8') as f:
                        f.write(html)
                        
                    enviar_para_github()
                except Exception as e:
                    print(f"  [ERRO] {e}")
                print("-" * 60)
            else:
                print(f"\r  Monitorando {PLANILHA}... {time.strftime('%H:%M:%S')}", end="", flush=True)

    except KeyboardInterrupt:
        print("\n  Monitor de apontamentos encerrado.")

if __name__ == '__main__':
    monitorar()