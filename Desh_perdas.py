"""
╔═══════════════════════════════════════════════════════════════╗
║  GERADOR DE DASHBOARD — PERDAS DE COLHEITA RAÍZEN            ║
║  Safra 25/26                                                  ║
╠═══════════════════════════════════════════════════════════════╣
║  COMO USAR:                                                   ║
║  1. Atualize os dados na aba "Base colheita" do Excel         ║
║  2. Salve o Excel                                             ║
║  3. Rode este script no terminal do VSCode:                   ║
║     python gerar_dashboard.py                                 ║
║  4. O arquivo "dashboard_perdas_colheita.html" será gerado    ║
║     na mesma pasta do script                                  ║
╠═══════════════════════════════════════════════════════════════╣
║  CONFIGURAÇÕES EDITÁVEIS (abaixo)                             ║
╚═══════════════════════════════════════════════════════════════╝
"""

import pandas as pd
import json
import os
import sys
from pathlib import Path

# ╔═══════════════════════════════════════════════════════════╗
# ║  CONFIGURAÇÕES — EDITE AQUI SE NECESSÁRIO                ║
# ╚═══════════════════════════════════════════════════════════╝

# Nome do arquivo Excel (deve estar na mesma pasta do script)
ARQUIVO_EXCEL = "CCT_Safra_25_26_Dashboard_Raizen.xlsx"

# Nome da aba com os dados
ABA_DADOS = "Base colheita"

# Nome do HTML de saída
ARQUIVO_SAIDA = "dashboard_perdas_colheita.html"

# Meta de perdas (kg)
META = 3.50

# Margem de tolerância (10% da meta = 0.35)
TOLERANCIA = 0.35

# Colunas de perdas que serão somadas
COLUNAS_PERDAS = [
    'AGR_TOUCEIRA',
    'AGR_PALMITO',
    'CANA_INTEIRA',
    'REBOLO',
    'PEDACO',
    'ESTILHACO',
    'LASCA'
]

# Turnos válidos (exclui registros sem turno)
TURNOS_VALIDOS = ['A', 'B', 'C']


# ╔═══════════════════════════════════════════════════════════╗
# ║  PROCESSAMENTO DOS DADOS                                  ║
# ╚═══════════════════════════════════════════════════════════╝

def processar_dados(caminho_excel):
    """Lê o Excel e retorna os dados processados como dicionário."""

    print(f"📂 Lendo arquivo: {caminho_excel}")
    df = pd.read_excel(caminho_excel, sheet_name=ABA_DADOS)
    print(f"   → {len(df)} registros encontrados na aba '{ABA_DADOS}'")

    # Filtrar registros com dados de perdas válidos
    df_valid = df[df[COLUNAS_PERDAS].notna().any(axis=1)].copy()
    df_valid['TOTAL_PERDAS'] = df_valid[COLUNAS_PERDAS].fillna(0).sum(axis=1)

    # Filtrar turnos válidos
    df_valid = df_valid[df_valid['TURNO'].isin(TURNOS_VALIDOS)]
    print(f"   → {len(df_valid)} registros válidos (com perdas e turno)")

    if len(df_valid) == 0:
        print("❌ ERRO: Nenhum registro válido encontrado!")
        sys.exit(1)

    # Função auxiliar para média ponderada
    def media_fill(x):
        return x.fillna(0).mean()

    # 1. Agrupamento por Equipamento × Frente
    equip_frente = df_valid.groupby(['EQUIP', 'FRENTE']).agg(
        media=('TOTAL_PERDAS', 'mean'),
        count=('TOTAL_PERDAS', 'count'),
        agr_touceira=('AGR_TOUCEIRA', media_fill),
        agr_palmito=('AGR_PALMITO', media_fill),
        cana_inteira=('CANA_INTEIRA', media_fill),
        rebolo=('REBOLO', media_fill),
        pedaco=('PEDACO', media_fill),
        estilhaco=('ESTILHACO', media_fill),
        lasca=('LASCA', media_fill),
    ).reset_index()
    equip_frente.columns = ['equip', 'frente', 'media', 'count',
                            'agr_touceira', 'agr_palmito', 'cana_inteira',
                            'rebolo', 'pedaco', 'estilhaco', 'lasca']

    # 2. Detalhamento por Equipamento × Frente × Turno
    detail = df_valid.groupby(['EQUIP', 'FRENTE', 'TURNO']).agg(
        media=('TOTAL_PERDAS', 'mean'),
        count=('TOTAL_PERDAS', 'count'),
        agr_touceira=('AGR_TOUCEIRA', media_fill),
        agr_palmito=('AGR_PALMITO', media_fill),
        cana_inteira=('CANA_INTEIRA', media_fill),
        rebolo=('REBOLO', media_fill),
        pedaco=('PEDACO', media_fill),
        estilhaco=('ESTILHACO', media_fill),
        lasca=('LASCA', media_fill),
    ).reset_index()
    detail.columns = ['equip', 'frente', 'turno', 'media', 'count',
                       'agr_touceira', 'agr_palmito', 'cana_inteira',
                       'rebolo', 'pedaco', 'estilhaco', 'lasca']

    # 3. Resumo por Frente
    frente_sum = df_valid.groupby('FRENTE').agg(
        media=('TOTAL_PERDAS', 'mean'),
        count=('TOTAL_PERDAS', 'count')
    ).reset_index()
    frente_sum.columns = ['frente', 'media', 'count']

    # 4. Resumo por Turno
    turno_sum = df_valid.groupby('TURNO').agg(
        media=('TOTAL_PERDAS', 'mean'),
        count=('TOTAL_PERDAS', 'count')
    ).reset_index()
    turno_sum.columns = ['turno', 'media', 'count']

    # 5. Geral
    overall = {
        'media': round(float(df_valid['TOTAL_PERDAS'].mean()), 3),
        'count': int(df_valid['TOTAL_PERDAS'].count()),
        'meta': META
    }

    # 6. Composição de perdas
    loss_breakdown = {}
    for col, name in zip(COLUNAS_PERDAS, ['agr_touceira', 'agr_palmito', 'cana_inteira',
                                           'rebolo', 'pedaco', 'estilhaco', 'lasca']):
        loss_breakdown[name] = round(float(df_valid[col].fillna(0).mean()), 3)

    data = {
        'equip_frente': equip_frente.round(3).to_dict('records'),
        'detail': detail.round(3).to_dict('records'),
        'frente_summary': frente_sum.round(3).to_dict('records'),
        'turno_summary': turno_sum.round(3).to_dict('records'),
        'overall': overall,
        'loss_breakdown': loss_breakdown,
        'frentes': sorted([int(x) for x in df_valid['FRENTE'].unique()]),
        'equipamentos': sorted([int(x) for x in df_valid['EQUIP'].unique()]),
        'turnos': sorted(df_valid['TURNO'].unique().tolist())
    }

    # Converter tipos numpy
    def converter(obj):
        import numpy as np
        if isinstance(obj, (np.integer,)): return int(obj)
        if isinstance(obj, (np.floating,)): return float(obj)
        if isinstance(obj, np.ndarray): return obj.tolist()
        return obj

    return json.loads(json.dumps(data, default=converter))


# ╔═══════════════════════════════════════════════════════════╗
# ║  TEMPLATE HTML DO DASHBOARD                                ║
# ╚═══════════════════════════════════════════════════════════╝

def get_template_html():
    return '''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard Perdas de Colheita — Raízen Safra 25/26</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"><\\/script>
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,500;0,9..40,700;1,9..40,400&family=JetBrains+Mono:wght@400;600&display=swap');

  :root {
    --bg: #0b0f14;
    --surface: #131920;
    --surface-2: #1a222b;
    --surface-3: #222c38;
    --border: #2a3544;
    --text: #e8ecf1;
    --text-muted: #8494a7;
    --text-dim: #5a6b80;
    --green: #22c55e;
    --green-bg: rgba(34,197,94,0.1);
    --green-border: rgba(34,197,94,0.25);
    --red: #ef4444;
    --red-bg: rgba(239,68,68,0.1);
    --red-border: rgba(239,68,68,0.25);
    --yellow: #eab308;
    --yellow-bg: rgba(234,179,8,0.1);
    --yellow-border: rgba(234,179,8,0.25);
    --accent: #3b82f6;
    --accent-bg: rgba(59,130,246,0.1);
  }

  * { margin: 0; padding: 0; box-sizing: border-box; }

  body {
    font-family: 'DM Sans', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    line-height: 1.5;
  }

  .header {
    padding: 28px 32px 20px;
    border-bottom: 1px solid var(--border);
    background: linear-gradient(180deg, #111820 0%, var(--bg) 100%);
  }
  .header-top { display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 16px; }
  .header h1 { font-size: 22px; font-weight: 700; letter-spacing: -0.5px; }
  .header h1 span { color: var(--accent); }
  .header-sub { font-size: 13px; color: var(--text-muted); margin-top: 4px; font-weight: 300; }
  .badge-meta {
    display: inline-flex; align-items: center; gap: 6px;
    background: var(--surface-2); border: 1px solid var(--border);
    border-radius: 8px; padding: 6px 14px; font-size: 13px;
    font-family: 'JetBrains Mono', monospace;
  }
  .badge-meta .label { color: var(--text-muted); }
  .badge-update { font-size: 11px; color: var(--text-dim); margin-top: 6px; }

  .filters {
    display: flex; gap: 10px; padding: 16px 32px; flex-wrap: wrap;
    border-bottom: 1px solid var(--border);
    background: var(--surface);
  }
  .filter-btn {
    padding: 7px 18px; border-radius: 6px; border: 1px solid var(--border);
    background: transparent; color: var(--text-muted); cursor: pointer;
    font-family: 'DM Sans', sans-serif; font-size: 13px; font-weight: 500;
    transition: all .2s;
  }
  .filter-btn:hover { border-color: var(--text-dim); color: var(--text); }
  .filter-btn.active {
    background: var(--accent-bg); border-color: var(--accent);
    color: var(--accent); font-weight: 600;
  }
  .filter-label { display: flex; align-items: center; font-size: 12px; color: var(--text-dim); text-transform: uppercase; letter-spacing: 1px; font-weight: 600; margin-right: 4px; }

  .main { padding: 24px 32px 40px; max-width: 1440px; margin: 0 auto; }

  .kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 14px; margin-bottom: 24px; }
  .kpi-card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 10px; padding: 18px 20px; position: relative; overflow: hidden;
  }
  .kpi-card::before {
    content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px;
  }
  .kpi-card.status-green::before { background: var(--green); }
  .kpi-card.status-red::before { background: var(--red); }
  .kpi-card.status-yellow::before { background: var(--yellow); }
  .kpi-label { font-size: 11px; color: var(--text-dim); text-transform: uppercase; letter-spacing: 1px; font-weight: 600; }
  .kpi-value {
    font-family: 'JetBrains Mono', monospace; font-size: 28px; font-weight: 700;
    margin: 6px 0 2px;
  }
  .kpi-unit { font-size: 14px; color: var(--text-muted); font-weight: 400; }
  .kpi-status { font-size: 12px; margin-top: 4px; }

  .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 24px; }

  .card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 10px; overflow: hidden;
  }
  .card-header {
    padding: 14px 18px; border-bottom: 1px solid var(--border);
    display: flex; align-items: center; justify-content: space-between;
  }
  .card-title { font-size: 14px; font-weight: 600; }
  .card-body { padding: 16px 18px; }

  table { width: 100%; border-collapse: collapse; }
  th {
    text-align: left; font-size: 10px; text-transform: uppercase;
    letter-spacing: 1px; color: var(--text-dim); padding: 8px 10px;
    border-bottom: 1px solid var(--border); font-weight: 600;
  }
  td {
    padding: 10px 10px; border-bottom: 1px solid rgba(42,53,68,0.5);
    font-size: 13px;
  }
  tr:last-child td { border-bottom: none; }
  tr:hover td { background: rgba(59,130,246,0.03); }

  .mono { font-family: 'JetBrains Mono', monospace; font-size: 12px; }
  .text-green { color: var(--green); }
  .text-red { color: var(--red); }
  .text-yellow { color: var(--yellow); }
  .text-muted { color: var(--text-muted); }

  .status-pill {
    display: inline-flex; align-items: center; gap: 4px;
    padding: 3px 10px; border-radius: 99px; font-size: 11px; font-weight: 600;
  }
  .pill-green { background: var(--green-bg); color: var(--green); border: 1px solid var(--green-border); }
  .pill-red { background: var(--red-bg); color: var(--red); border: 1px solid var(--red-border); }
  .pill-yellow { background: var(--yellow-bg); color: var(--yellow); border: 1px solid var(--yellow-border); }

  .bar-cell { display: flex; align-items: center; gap: 8px; }
  .bar-track { flex: 1; height: 6px; background: var(--surface-3); border-radius: 3px; overflow: hidden; position: relative; }
  .bar-fill { height: 100%; border-radius: 3px; transition: width .4s ease; }
  .bar-meta { position: absolute; right: 0; top: -1px; width: 1px; height: 8px; }

  .chart-container { position: relative; height: 240px; }

  .breakdown-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr)); gap: 10px; }
  .breakdown-item {
    background: var(--surface-2); border-radius: 8px; padding: 12px 14px;
    border: 1px solid var(--border);
  }
  .breakdown-label { font-size: 10px; color: var(--text-dim); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px; }
  .breakdown-value { font-family: 'JetBrains Mono', monospace; font-size: 18px; font-weight: 600; }
  .breakdown-bar { height: 3px; border-radius: 2px; margin-top: 8px; background: var(--surface-3); overflow: hidden; }
  .breakdown-bar-fill { height: 100%; border-radius: 2px; background: var(--accent); }

  .rank-num {
    width: 24px; height: 24px; border-radius: 6px; display: inline-flex;
    align-items: center; justify-content: center; font-size: 11px; font-weight: 700;
    font-family: 'JetBrains Mono', monospace;
  }
  .rank-1 { background: var(--red-bg); color: var(--red); border: 1px solid var(--red-border); }
  .rank-2 { background: var(--red-bg); color: var(--red); border: 1px solid var(--red-border); }
  .rank-3 { background: var(--yellow-bg); color: var(--yellow); border: 1px solid var(--yellow-border); }
  .rank-default { background: var(--surface-3); color: var(--text-muted); }

  @media (max-width: 900px) {
    .grid-2 { grid-template-columns: 1fr; }
    .main { padding: 16px; }
    .header, .filters { padding-left: 16px; padding-right: 16px; }
    .kpi-row { grid-template-columns: repeat(2, 1fr); }
  }
</style>
</head>
<body>

<div class="header">
  <div class="header-top">
    <div>
      <h1>Perdas de Colheita <span>— Raízen</span></h1>
      <div class="header-sub">Safra 25/26 · CCT Dashboard de Qualidade</div>
      <div class="badge-update">Atualizado em: %%TIMESTAMP%%</div>
    </div>
    <div class="badge-meta">
      <span class="label">Meta:</span> <strong>%%META%% kg</strong>
    </div>
  </div>
</div>

<div class="filters" id="filters">
  <div class="filter-label">Frente</div>
  <button class="filter-btn active" data-frente="all" onclick="setFilter('all')">Todas</button>
</div>

<div class="main">
  <div class="kpi-row" id="kpiRow"></div>

  <div class="card" style="margin-bottom:24px">
    <div class="card-header">
      <div class="card-title">Composição das Perdas (Média kg)</div>
    </div>
    <div class="card-body">
      <div class="breakdown-grid" id="breakdownGrid"></div>
    </div>
  </div>

  <div class="grid-2">
    <div class="card">
      <div class="card-header"><div class="card-title">Média de Perdas por Frente</div></div>
      <div class="card-body"><div class="chart-container"><canvas id="chartFrente"></canvas></div></div>
    </div>
    <div class="card">
      <div class="card-header"><div class="card-title">Média de Perdas por Turno</div></div>
      <div class="card-body"><div class="chart-container"><canvas id="chartTurno"></canvas></div></div>
    </div>
  </div>

  <div class="card" style="margin-bottom:24px">
    <div class="card-header">
      <div class="card-title">🏆 Ranking de Equipamentos por Perda</div>
      <div style="font-size:11px;color:var(--text-dim)">Ordenado por maior média de perdas</div>
    </div>
    <div class="card-body" style="padding:0;overflow-x:auto">
      <table id="rankingTable">
        <thead>
          <tr>
            <th>#</th><th>Equipamento</th><th>Frente</th><th>Aval.</th>
            <th>Média Perdas</th><th>Visual</th><th>Status</th>
            <th>Touc.</th><th>Palm.</th><th>Inteira</th><th>Rebolo</th>
            <th>Pedaço</th><th>Estil.</th><th>Lasca</th>
          </tr>
        </thead>
        <tbody id="rankingBody"></tbody>
      </table>
    </div>
  </div>

  <div class="card">
    <div class="card-header">
      <div class="card-title">📋 Detalhamento por Equipamento × Turno</div>
      <div style="font-size:11px;color:var(--text-dim)">Cada combinação de equipamento, frente e turno</div>
    </div>
    <div class="card-body" style="padding:0;overflow-x:auto">
      <table id="detailTable">
        <thead>
          <tr>
            <th>#</th><th>Equip.</th><th>Frente</th><th>Turno</th><th>Aval.</th>
            <th>Média Perdas</th><th>Visual</th><th>Status</th>
          </tr>
        </thead>
        <tbody id="detailBody"></tbody>
      </table>
    </div>
  </div>
</div>

<script>
const DATA = %%DATA_JSON%%;
const META = %%META_JS%%;
const TOLERANCE = %%TOLERANCE_JS%%;

let currentFilter = 'all';
let chartFrenteInst = null;
let chartTurnoInst = null;

function getStatus(val) {
  if (val > META) return 'red';
  if (val >= META - TOLERANCE) return 'yellow';
  return 'green';
}
function getStatusLabel(val) {
  const s = getStatus(val);
  if (s === 'red') return '⚠️ Acima';
  if (s === 'yellow') return '🟡 Atenção';
  return '✅ Dentro';
}
function statusPill(val) {
  const s = getStatus(val);
  return `<span class="status-pill pill-${s}">${getStatusLabel(val)}</span>`;
}
function fmt(v) { return v.toFixed(2).replace('.', ','); }
function rankClass(i) {
  if (i <= 2) return 'rank-1';
  if (i <= 4) return 'rank-2';
  if (i <= 6) return 'rank-3';
  return 'rank-default';
}

function filterData(arr) {
  if (currentFilter === 'all') return arr;
  return arr.filter(r => String(r.frente) === currentFilter);
}

function aggregate(records) {
  const total = records.reduce((s, r) => s + r.count, 0);
  if (total === 0) return null;
  const fields = ['media', 'agr_touceira', 'agr_palmito', 'cana_inteira', 'rebolo', 'pedaco', 'estilhaco', 'lasca'];
  const agg = { count: total };
  fields.forEach(f => {
    agg[f] = records.reduce((s, r) => s + (r[f] || 0) * r.count, 0) / total;
  });
  return agg;
}

function renderKPIs() {
  const filtered = filterData(DATA.equip_frente);
  const agg = aggregate(filtered);
  if (!agg) return;
  const status = getStatus(agg.media);
  const aboveMeta = filtered.filter(r => r.media > META).length;
  const attn = filtered.filter(r => r.media >= META - TOLERANCE && r.media <= META).length;

  document.getElementById('kpiRow').innerHTML = `
    <div class="kpi-card status-${status}">
      <div class="kpi-label">Média Geral de Perdas</div>
      <div class="kpi-value text-${status}">${fmt(agg.media)} <span class="kpi-unit">kg</span></div>
      <div class="kpi-status">${statusPill(agg.media)}</div>
    </div>
    <div class="kpi-card status-green">
      <div class="kpi-label">Total de Avaliações</div>
      <div class="kpi-value">${agg.count.toLocaleString('pt-BR')}</div>
      <div class="kpi-status text-muted">registros válidos</div>
    </div>
    <div class="kpi-card ${aboveMeta > 0 ? 'status-red' : 'status-green'}">
      <div class="kpi-label">Equip. Acima da Meta</div>
      <div class="kpi-value ${aboveMeta > 0 ? 'text-red' : 'text-green'}">${aboveMeta}</div>
      <div class="kpi-status text-muted">de ${filtered.length} combinações equip×frente</div>
    </div>
    <div class="kpi-card ${attn > 0 ? 'status-yellow' : 'status-green'}">
      <div class="kpi-label">Equip. em Atenção</div>
      <div class="kpi-value ${attn > 0 ? 'text-yellow' : 'text-green'}">${attn}</div>
      <div class="kpi-status text-muted">entre ${fmt(META - TOLERANCE)} e ${fmt(META)} kg</div>
    </div>`;
}

function renderBreakdown() {
  const filtered = filterData(DATA.equip_frente);
  const agg = aggregate(filtered);
  if (!agg) return;
  const items = [
    { key: 'pedaco', label: 'Pedaço' },
    { key: 'agr_touceira', label: 'Agr. Touceira' },
    { key: 'lasca', label: 'Lasca' },
    { key: 'cana_inteira', label: 'Cana Inteira' },
    { key: 'rebolo', label: 'Rebolo' },
    { key: 'estilhaco', label: 'Estilhaço' },
    { key: 'agr_palmito', label: 'Agr. Palmito' },
  ].sort((a, b) => (agg[b.key] || 0) - (agg[a.key] || 0));
  const maxVal = Math.max(...items.map(i => agg[i.key] || 0));
  document.getElementById('breakdownGrid').innerHTML = items.map(item => {
    const val = agg[item.key] || 0;
    const pct = maxVal > 0 ? (val / maxVal * 100) : 0;
    return `<div class="breakdown-item">
        <div class="breakdown-label">${item.label}</div>
        <div class="breakdown-value">${fmt(val)}</div>
        <div class="breakdown-bar"><div class="breakdown-bar-fill" style="width:${pct}%"></div></div>
      </div>`;
  }).join('');
}

function makeBarChart(canvasId, labels, values, existingChart) {
  if (existingChart) existingChart.destroy();
  const ctx = document.getElementById(canvasId).getContext('2d');
  return new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: values.map(v => {
          const s = getStatus(v);
          return s === 'red' ? 'rgba(239,68,68,0.7)' : s === 'yellow' ? 'rgba(234,179,8,0.7)' : 'rgba(34,197,94,0.7)';
        }),
        borderColor: values.map(v => {
          const s = getStatus(v);
          return s === 'red' ? '#ef4444' : s === 'yellow' ? '#eab308' : '#22c55e';
        }),
        borderWidth: 1, borderRadius: 4,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: { beginAtZero: true, grid: { color: 'rgba(42,53,68,0.5)' }, ticks: { color: '#8494a7', font: { family: 'JetBrains Mono', size: 11 } } },
        x: { grid: { display: false }, ticks: { color: '#8494a7', font: { family: 'DM Sans', size: 12 } } }
      }
    },
    plugins: [{
      afterDraw(chart) {
        const ctx2 = chart.ctx;
        const yScale = chart.scales.y;
        const xScale = chart.scales.x;
        const y = yScale.getPixelForValue(META);
        ctx2.save();
        ctx2.strokeStyle = '#ef4444'; ctx2.lineWidth = 1.5; ctx2.setLineDash([6, 4]);
        ctx2.beginPath(); ctx2.moveTo(xScale.left, y); ctx2.lineTo(xScale.right, y); ctx2.stroke();
        ctx2.fillStyle = '#ef4444'; ctx2.font = '10px JetBrains Mono';
        ctx2.fillText('META ' + META.toFixed(2).replace('.', ','), xScale.right - 70, y - 5);
        ctx2.restore();
      }
    }]
  });
}

function renderCharts() {
  const frenteData = DATA.frente_summary
    .filter(r => currentFilter === 'all' || String(r.frente) === currentFilter)
    .sort((a, b) => b.media - a.media);
  chartFrenteInst = makeBarChart('chartFrente',
    frenteData.map(r => 'Frente ' + r.frente),
    frenteData.map(r => r.media), chartFrenteInst);

  let turnoDisplay = DATA.turno_summary;
  if (currentFilter !== 'all') {
    const filtered = DATA.detail.filter(r => String(r.frente) === currentFilter);
    const grouped = {};
    filtered.forEach(r => {
      if (!grouped[r.turno]) grouped[r.turno] = { turno: r.turno, totalW: 0, totalC: 0 };
      grouped[r.turno].totalW += r.media * r.count;
      grouped[r.turno].totalC += r.count;
    });
    turnoDisplay = Object.values(grouped).map(g => ({ turno: g.turno, media: g.totalW / g.totalC }));
  }
  turnoDisplay.sort((a, b) => a.turno.localeCompare(b.turno));
  chartTurnoInst = makeBarChart('chartTurno',
    turnoDisplay.map(r => 'Turno ' + r.turno),
    turnoDisplay.map(r => r.media), chartTurnoInst);
}

function renderRanking() {
  const filtered = filterData(DATA.equip_frente).sort((a, b) => b.media - a.media);
  const maxMedia = Math.max(...filtered.map(r => r.media), META);
  document.getElementById('rankingBody').innerHTML = filtered.map((r, i) => {
    const s = getStatus(r.media);
    const barPct = (r.media / (maxMedia * 1.1)) * 100;
    return `<tr>
      <td><span class="rank-num ${rankClass(i+1)}">${i+1}</span></td>
      <td class="mono" style="font-weight:600">${r.equip}</td>
      <td>${r.frente}</td>
      <td class="text-muted">${r.count}</td>
      <td class="mono text-${s}" style="font-weight:600">${fmt(r.media)} kg</td>
      <td style="min-width:140px"><div class="bar-cell"><div class="bar-track"><div class="bar-fill" style="width:${barPct}%;background:${s==='red'?'var(--red)':s==='yellow'?'var(--yellow)':'var(--green)'}"></div></div></div></td>
      <td>${statusPill(r.media)}</td>
      <td class="mono text-muted">${fmt(r.agr_touceira||0)}</td>
      <td class="mono text-muted">${fmt(r.agr_palmito||0)}</td>
      <td class="mono text-muted">${fmt(r.cana_inteira||0)}</td>
      <td class="mono text-muted">${fmt(r.rebolo||0)}</td>
      <td class="mono text-muted">${fmt(r.pedaco||0)}</td>
      <td class="mono text-muted">${fmt(r.estilhaco||0)}</td>
      <td class="mono text-muted">${fmt(r.lasca||0)}</td>
    </tr>`;
  }).join('');
}

function renderDetail() {
  const filtered = filterData(DATA.detail).sort((a, b) => b.media - a.media);
  const maxMedia = Math.max(...filtered.map(r => r.media), META);
  document.getElementById('detailBody').innerHTML = filtered.map((r, i) => {
    const s = getStatus(r.media);
    const barPct = (r.media / (maxMedia * 1.1)) * 100;
    return `<tr>
      <td><span class="rank-num ${rankClass(i+1)}">${i+1}</span></td>
      <td class="mono" style="font-weight:600">${r.equip}</td>
      <td>${r.frente}</td>
      <td style="font-weight:600">Turno ${r.turno}</td>
      <td class="text-muted">${r.count}</td>
      <td class="mono text-${s}" style="font-weight:600">${fmt(r.media)} kg</td>
      <td style="min-width:120px"><div class="bar-cell"><div class="bar-track"><div class="bar-fill" style="width:${barPct}%;background:${s==='red'?'var(--red)':s==='yellow'?'var(--yellow)':'var(--green)'}"></div></div></div></td>
      <td>${statusPill(r.media)}</td>
    </tr>`;
  }).join('');
}

function setFilter(frente) {
  currentFilter = frente;
  document.querySelectorAll('.filter-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.frente === String(frente));
  });
  renderAll();
}

function renderAll() {
  renderKPIs(); renderBreakdown(); renderCharts(); renderRanking(); renderDetail();
}

const filtersEl = document.getElementById('filters');
DATA.frentes.forEach(f => {
  const btn = document.createElement('button');
  btn.className = 'filter-btn';
  btn.dataset.frente = String(f);
  btn.textContent = 'Frente ' + f;
  btn.onclick = () => setFilter(String(f));
  filtersEl.appendChild(btn);
});

renderAll();
</script>
</body>
</html>'''


# ╔═══════════════════════════════════════════════════════════╗
# ║  GERAÇÃO DO HTML FINAL                                    ║
# ╚═══════════════════════════════════════════════════════════╝

def gerar_dashboard(dados, caminho_saida):
    """Injeta os dados no template HTML e salva o arquivo."""
    from datetime import datetime

    html = get_template_html()

    # Substituir placeholders
    timestamp = datetime.now().strftime("%d/%m/%Y às %H:%M")
    html = html.replace('%%TIMESTAMP%%', timestamp)
    html = html.replace('%%META%%', f"{META:.2f}".replace('.', ','))
    html = html.replace('%%META_JS%%', str(META))
    html = html.replace('%%TOLERANCE_JS%%', str(TOLERANCIA))
    html = html.replace('%%DATA_JSON%%', json.dumps(dados))

    with open(caminho_saida, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"✅ Dashboard gerado com sucesso!")
    print(f"   → Arquivo: {caminho_saida}")
    print(f"   → Tamanho: {os.path.getsize(caminho_saida) / 1024:.0f} KB")
    print(f"   → Timestamp: {timestamp}")


# ╔═══════════════════════════════════════════════════════════╗
# ║  EXECUÇÃO PRINCIPAL                                        ║
# ╚═══════════════════════════════════════════════════════════╝

if __name__ == '__main__':
    # Determinar pasta do script
    pasta_script = Path(__file__).parent.resolve()

    # Caminhos
    caminho_excel = pasta_script / ARQUIVO_EXCEL
    caminho_saida = pasta_script / ARQUIVO_SAIDA

    # Verificar se o Excel existe
    if not caminho_excel.exists():
        print(f"❌ ERRO: Arquivo não encontrado: {caminho_excel}")
        print(f"   Certifique-se de que o arquivo '{ARQUIVO_EXCEL}' está na mesma pasta do script.")
        sys.exit(1)

    print("=" * 60)
    print("  GERADOR DE DASHBOARD — PERDAS DE COLHEITA RAÍZEN")
    print("=" * 60)
    print()

    # Processar
    dados = processar_dados(str(caminho_excel))

    # Resumo rápido
    print()
    print(f"📊 Resumo dos dados:")
    print(f"   → Média geral: {dados['overall']['media']:.2f} kg (meta: {META:.2f} kg)")
    print(f"   → Avaliações: {dados['overall']['count']}")
    print(f"   → Frentes: {dados['frentes']}")
    print(f"   → Equipamentos: {len(dados['equipamentos'])}")
    print()

    # Gerar
    gerar_dashboard(dados, str(caminho_saida))

    print()
    print("💡 Para visualizar, abra o arquivo HTML no navegador.")
    print("   Ou no VSCode: botão direito → 'Open with Live Server'")
    print()
