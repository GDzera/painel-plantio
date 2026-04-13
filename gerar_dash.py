"""
╔══════════════════════════════════════════════════════════════╗
║  GERADOR DE DASHBOARD — PLANTIO 18M RAÍZEN                   ║
║  Lê a planilha Excel, gera o HTML e atualiza na Nuvem        ║
╚══════════════════════════════════════════════════════════════╝
"""

import pandas as pd
import json
import os
import sys
import time
import re
import unicodedata
from pathlib import Path
from datetime import datetime

try:
    from openpyxl import load_workbook
except ImportError:
    print("Instale: pip install openpyxl pandas")
    sys.exit(1)

# ═══════════════════════════════════════════════════════════
#  CONFIGURACAO — altere apenas estas linhas se necessario
# ═══════════════════════════════════════════════════════════
ARQUIVOS_SAFRA = [
    ("Acompanhamento_Plantio sf25'26.xlsm", "25/26"),
    ("Acompanhamento_Plantio s26'27.xlsm",  "26/27"),
]
ARQUIVO_HTML  = "index.html"
INTERVALO_SEG = 5
ABRIR_BROWSER = True
# ═══════════════════════════════════════════════════════════

def safe_int(val):
    if pd.isnull(val): return 0
    if isinstance(val, (int, float)): return int(val)
    s = str(val).strip()
    if not s or s.lower() == 'nan': return 0
    nums = re.findall(r'\d+', s)
    if nums: return int(nums[0])
    return 0

def limpar_nome(v):
    if pd.isnull(v): return ""
    s = str(v).strip()
    if s.lower() == 'nan' or "►" in s or "botões" in s.lower() or s.upper().startswith("TOTAL"):
        return ""
    s = re.sub(r'[^\w\s\.\-\/\(\)]', '', s).strip()
    return s

def padronizar_chave(n):
    s = unicodedata.normalize('NFKD', str(n)).encode('ASCII', 'ignore').decode('utf-8')
    return re.sub(r'[^a-z0-9]', '', s.lower())

def ler_dados(arquivo_excel):
    wb = load_workbook(arquivo_excel, data_only=True)
    meta = float(wb["Dashboard"]["B6"].value)
    xl = pd.ExcelFile(arquivo_excel)
    xl_sheets = xl.sheet_names
    
    # 1. BASE DE PLANTIO
    df = pd.read_excel(xl, sheet_name="Base Plantio 2526")
    df["Data_plantio"] = pd.to_datetime(df["Data_plantio"], errors='coerce', dayfirst=True)
    
    if df["Area_ha"].dtype == object:
        def fix_num(x):
            s = str(x).strip()
            if s.lower() == 'nan' or s == '': return 0.0
            if ',' in s and '.' in s: s = s.replace('.', '')
            s = s.replace(',', '.')
            try: return float(s)
            except: return 0.0
        df["Area_ha"] = df["Area_ha"].apply(fix_num)
    df["Area_ha"] = pd.to_numeric(df["Area_ha"], errors='coerce').fillna(0.0)
    df["Variedade"] = df["Variedade"].astype(str).str.strip()
    df["Ciclo de plantio"] = df["Ciclo de plantio"].astype(str).str.strip()

    # 2. FAZENDAS
    df_fazendas = pd.DataFrame()
    aba_fazendas = None
    try:
        for nome in xl_sheets:
            nl = padronizar_chave(nome)
            if "controle" in nl and "fazenda" in nl:
                aba_fazendas = nome
                break
        if aba_fazendas:
            df_temp_faz = pd.read_excel(xl, sheet_name=aba_fazendas, header=None)
            idx_cab_faz = 0
            for i, row in df_temp_faz.iterrows():
                valores = row.dropna().astype(str).str.lower().str.replace('\n', ' ').str.strip().tolist()
                if any("cod" in v or "cód" in v for v in valores) and any("fazenda" in v for v in valores):
                    idx_cab_faz = i
                    break
            df_fazendas = pd.read_excel(xl, sheet_name=aba_fazendas, header=idx_cab_faz)
        else:
            print("  [AVISO] Aba 'Controle Fazendas' não encontrada.")
    except Exception as e:
        print(f"  [AVISO] Erro fazendas: {e}")

    # 3. METAS DIÁRIAS
    erro_metas = ""
    try:
        aba_metas = None
        for nome in xl_sheets:
            nome_limpo = padronizar_chave(nome)
            if "metadiaria" in nome_limpo or "metasdiarias" in nome_limpo:
                aba_metas = nome
                break
        if not aba_metas:
            raise ValueError("Aba 'Meta diaria' não encontrada.")

        df_temp = pd.read_excel(xl, sheet_name=aba_metas, header=None)
        
        idx_cabecalho = -1
        for i, row in df_temp.iterrows():
            row_str = " ".join([str(x).lower() for x in row.dropna()])
            if "frente 140" in row_str and "data" in row_str:
                idx_cabecalho = i
                break
        
        if idx_cabecalho == -1:
            raise ValueError("Cabeçalho diário com 'Frente 140' não encontrado na aba Meta diaria.")

        header_row = df_temp.iloc[idx_cabecalho]
        col_map = {}
        for ci, val in enumerate(header_row):
            if pd.isnull(val): continue
            vl = str(val).strip().lower()
            if vl == 'data': col_map['Data'] = ci
            elif vl == 'dia': col_map['Dia'] = ci
            elif 'dia semana' in vl or 'dia_semana' in vl: col_map['DiaSemana'] = ci
            elif 'frente' in vl: col_map['Frente140'] = ci
            elif vl == 'total': col_map['Total'] = ci
            elif 'meta real' in vl: col_map['MetaReal'] = ci
            elif 'meta corrigida' in vl or 'meta wd' in vl: col_map['MetaCorrigida'] = ci
            elif 'nova meta' in vl: col_map['NovaMeta'] = ci
            elif 'desvio de meta' in vl: col_map['DesvioMeta'] = ci
            elif 'desvio dia' in vl: col_map['DesvioDia'] = ci

        print(f"  [META] Cabeçalho na linha {idx_cabecalho}. Colunas: {col_map}")

        df_dados = df_temp.iloc[idx_cabecalho+1:].copy()
        
        df_metas = pd.DataFrame()
        df_metas['Data'] = pd.to_datetime(df_dados.iloc[:, col_map.get('Data', 1)], errors='coerce', dayfirst=True)
        df_metas['Dia'] = df_dados.iloc[:, col_map.get('Dia', 2)].astype(str).str.replace('.0', '', regex=False).str.strip()
        df_metas['DiaSemana'] = df_dados.iloc[:, col_map.get('DiaSemana', 3)].astype(str).str.strip()
        
        col_real = col_map.get('Total', col_map.get('Frente140', 6))
        df_metas['Realizado'] = pd.to_numeric(df_dados.iloc[:, col_real].astype(str).str.replace(',', '.'), errors='coerce')
        df_metas['MetaReal'] = pd.to_numeric(df_dados.iloc[:, col_map.get('MetaReal', 7)].astype(str).str.replace(',', '.'), errors='coerce')
        df_metas['MetaCorrigida'] = pd.to_numeric(df_dados.iloc[:, col_map.get('MetaCorrigida', 8)].astype(str).str.replace(',', '.'), errors='coerce')
        df_metas['NovaMeta'] = pd.to_numeric(df_dados.iloc[:, col_map.get('NovaMeta', 9)].astype(str).str.replace(',', '.'), errors='coerce')
        
        if 'DesvioMeta' in col_map:
            df_metas['DesvioMeta'] = df_dados.iloc[:, col_map['DesvioMeta']].apply(
                lambda x: None if str(x).strip().lower() in ('false','nan','') else pd.to_numeric(str(x).replace(',','.'), errors='coerce')
            )
        else:
            df_metas['DesvioMeta'] = None
            
        if 'DesvioDia' in col_map:
            df_metas['DesvioDia'] = pd.to_numeric(df_dados.iloc[:, col_map['DesvioDia']].astype(str).str.replace(',', '.'), errors='coerce')
        else:
            df_metas['DesvioDia'] = None
        
        df_metas = df_metas.dropna(subset=["Data"]).copy()
        df_metas = df_metas[df_metas['Dia'].str.match(r'^\d+$', na=False)].copy()
        
        print(f"  [META] {len(df_metas)} dias carregados.")

    except Exception as e:
        erro_metas = str(e).replace('"', "'")
        print(f"  [AVISO] Metas: {e}")
        df_metas = pd.DataFrame(columns=["Data", "Dia", "DiaSemana", "Realizado", "MetaReal", "MetaCorrigida", "NovaMeta", "DesvioMeta", "DesvioDia"])

    # 4. STATUS INSUMOS
    df_insumos = pd.DataFrame()
    aba_insumos = None
    try:
        for nome in xl_sheets:
            nl = padronizar_chave(nome)
            if "statusinsumo" in nl or "statusdeinsumo" in nl:
                aba_insumos = nome
                break
        if aba_insumos:
            df_temp_ins = pd.read_excel(xl, sheet_name=aba_insumos, header=None)
            idx_cab_ins = 0
            for i, row in df_temp_ins.iterrows():
                valores = row.dropna().astype(str).str.lower().str.replace('\n', ' ').str.strip().tolist()
                if any("fazenda" in v for v in valores) and any("status" in v for v in valores):
                    idx_cab_ins = i
                    break
            df_insumos = pd.read_excel(xl, sheet_name=aba_insumos, header=idx_cab_ins)
            df_insumos = df_insumos.loc[:, ~df_insumos.columns.str.contains('^Unnamed')]
            df_insumos.columns = df_insumos.columns.astype(str).str.strip().str.replace('\n', ' ')
            if "Fazenda" in df_insumos.columns:
                df_insumos = df_insumos.dropna(subset=["Fazenda"])
            df_insumos = df_insumos.fillna("")
        else:
            print("  [AVISO] Aba 'Status Insumos' não encontrada.")
    except Exception as e:
        print(f"  [AVISO] Erro insumos: {e}")

    # 5. SALDO DE INSUMOS
    df_saldo = pd.DataFrame()
    try:
        if "Saldo Insumo" in xl_sheets:
            df_saldo = pd.read_excel(xl, sheet_name="Saldo Insumo")
            df_saldo.columns = df_saldo.columns.astype(str).str.strip()
            if "Data" in df_saldo.columns:
                df_saldo["Data"] = pd.to_datetime(df_saldo["Data"], errors='coerce').dt.strftime("%Y-%m-%d")
            df_saldo = df_saldo.fillna("")
            print(f"  [SALDO] {len(df_saldo)} registros de saldo insumo carregados.")
        else:
            print("  [AVISO] Aba 'Saldo Insumo' não encontrada.")
    except Exception as e:
        print(f"  [AVISO] Erro saldo insumo: {e}")

    return df, meta, df_metas, erro_metas, df_fazendas, df_insumos, df_saldo

def injetar_json(obj):
    return json.dumps(obj, separators=(",", ":"), ensure_ascii=False).replace("</", "<\\/")

def preparar_dados(df, df_metas, df_fazendas, df_insumos, df_saldo):
    hier = {}
    
    if not df_fazendas.empty:
        df_fazendas.columns = df_fazendas.columns.astype(str).str.strip().str.replace('\n', ' ')
        c_cod, c_faz = None, None
        for c in df_fazendas.columns:
            cl = str(c).lower().strip()
            if "cod" in cl or "cód" in cl:
                c_cod = c
            elif "fazenda" in cl or "nome" in cl or "prop" in cl:
                if not c_faz and "cod" not in cl and "cód" not in cl:
                    c_faz = c
        if c_cod and c_faz:
            for _, r in df_fazendas.iterrows():
                cod = safe_int(r.get(c_cod))
                if cod == 0: continue
                nome = limpar_nome(r.get(c_faz))
                if not nome: continue
                if cod not in hier:
                    hier[cod] = {"n": nome, "z": {}}

    registros = []
    for _, r in df.iterrows():
        cod = safe_int(r.get("Codigo"))
        if cod == 0: continue
        nome = limpar_nome(r.get("Fazenda", ""))
        zona = safe_int(r.get("Zona"))
        talh = safe_int(r.get("Talhao"))
        var  = str(r.get("Variedade", "")).strip()
        area = float(r["Area_ha"]) if pd.notnull(r.get("Area_ha")) else 0.0
        safra = str(r.get("_safra", "")) if "_safra" in df.columns else ""
        data_valida = pd.notnull(r.get("Data_plantio")) and r["Data_plantio"].year > 2000
        data = r["Data_plantio"].strftime("%Y-%m-%d") if data_valida else ""
        if data_valida:
            registros.append({
                "d": data, "a": area, "v": var, "f": nome, "c": cod,
                "z": zona, "t": talh, "ci": str(r.get("Ciclo de plantio", "")),
                "s": safra
            })
        if cod not in hier:
            hier[cod] = {"n": nome if nome else f"Fazenda {cod}", "z": {}}
        if nome and (not hier[cod]["n"] or hier[cod]["n"].startswith("Fazenda ")):
            hier[cod]["n"] = nome
        if zona not in hier[cod]["z"]:
            hier[cod]["z"][zona] = {}
        talh_key = f"{talh}_{safra}" if safra else str(talh)
        if talh_key not in hier[cod]["z"][zona]:
            hier[cod]["z"][zona][talh_key] = {"a": 0.0, "v": [], "d": data, "t": talh, "s": safra}
        t = hier[cod]["z"][zona][talh_key]
        t["a"] += area
        if var and var.lower() != 'nan' and var not in t["v"]: t["v"].append(var)
        if data and not t["d"]: t["d"] = data

    registros_metas = []
    safra_label = df["_safra"].iloc[0] if "_safra" in df.columns and len(df) > 0 else ""
    for _, r in df_metas.iterrows():
        registros_metas.append({
            "d": r["Data"].strftime("%Y-%m-%d"),
            "dia": str(r["Dia"]),
            "sem": str(r.get("DiaSemana", "")),
            "r": r["Realizado"] if pd.notnull(r["Realizado"]) else None,
            "mr": r["MetaReal"] if pd.notnull(r["MetaReal"]) else None,
            "mc": r["MetaCorrigida"] if pd.notnull(r["MetaCorrigida"]) else None,
            "n": r["NovaMeta"] if pd.notnull(r["NovaMeta"]) else None,
            "dm": r["DesvioMeta"] if pd.notnull(r.get("DesvioMeta")) else None,
            "dd": r["DesvioDia"] if pd.notnull(r.get("DesvioDia")) else None,
            "s": safra_label,
        })

    registros_insumos = df_insumos.to_dict('records') if not df_insumos.empty else []
    registros_saldo = []
    if not df_saldo.empty:
        registros_saldo = df_saldo.to_dict('records')

    return (injetar_json(registros), injetar_json(hier), injetar_json(registros_metas),
            injetar_json(registros_insumos), injetar_json(registros_saldo))

def ler_todas_safras():
    """Lê todas as planilhas de safra configuradas e mescla os dados"""
    all_df = []
    all_metas_df = []
    all_fazendas = []
    all_insumos = []
    all_saldo = []
    metas_por_safra = {}
    erros_metas = {}
    safras_encontradas = []
    
    for arquivo, safra_label in ARQUIVOS_SAFRA:
        if not os.path.exists(arquivo):
            print(f"  [AVISO] {arquivo} não encontrado, pulando safra {safra_label}")
            continue
        
        print(f"  Lendo safra {safra_label}: {arquivo}")
        try:
            df, meta, df_metas, erro_metas, df_fazendas, df_insumos, df_saldo = ler_dados(arquivo)
            df["_safra"] = safra_label
            df_metas["_safra"] = safra_label
            
            all_df.append(df)
            all_metas_df.append(df_metas)
            if not df_fazendas.empty: all_fazendas.append(df_fazendas)
            if not df_insumos.empty: all_insumos.append(df_insumos)
            if not df_saldo.empty: all_saldo.append(df_saldo)
            
            metas_por_safra[safra_label] = meta
            erros_metas[safra_label] = erro_metas
            safras_encontradas.append(safra_label)
            print(f"    Meta {safra_label}: {int(meta)} ha | {len(df)} registros | {len(df_metas)} dias de meta")
        except Exception as e:
            print(f"  [ERRO] Safra {safra_label}: {e}")
    
    if not all_df:
        print("  [ERRO] Nenhuma planilha de safra encontrada!")
        sys.exit(1)
    
    merged_df = pd.concat(all_df, ignore_index=True)
    # Metas diárias: usar apenas do último arquivo que tenha dados (mais recente)
    merged_metas = pd.DataFrame()
    for mdf in reversed(all_metas_df):
        if len(mdf) > 0:
            merged_metas = mdf
            break
    merged_fazendas = pd.concat(all_fazendas, ignore_index=True) if all_fazendas else pd.DataFrame()
    merged_insumos = pd.concat(all_insumos, ignore_index=True) if all_insumos else pd.DataFrame()
    merged_saldo = pd.concat(all_saldo, ignore_index=True) if all_saldo else pd.DataFrame()
    
    # Erro de metas: combinar
    erro_str = "; ".join(f"{s}: {e}" for s, e in erros_metas.items() if e)
    
    return (merged_df, metas_por_safra, merged_metas, erro_str, 
            merged_fazendas, merged_insumos, merged_saldo, safras_encontradas)


def gerar_html(df, metas_por_safra, df_metas, erro_metas, df_fazendas, df_insumos, df_saldo, safras):
    reg_json, hier_json, metas_json, insumos_json, saldo_json = preparar_dados(df, df_metas, df_fazendas, df_insumos, df_saldo)
    gerado_em = datetime.now().strftime("%d/%m/%Y %H:%M")
    
    # Use first safra's meta as default, inject all as JSON
    meta_default = list(metas_por_safra.values())[0] if metas_por_safra else 0
    metas_safra_json = json.dumps(metas_por_safra, separators=(",", ":"))
    safras_json = json.dumps(safras)
    
    datas_validas = df.loc[df["Data_plantio"].dt.year > 2000, "Data_plantio"]
    todas_datas = pd.concat([datas_validas, df_metas["Data"]]).dropna() if not df_metas.empty else datas_validas.dropna()
    if not todas_datas.empty:
        dt_min = todas_datas.min().strftime("%Y-%m-%d")
        dt_max = todas_datas.max().strftime("%Y-%m-%d")
    else:
        dt_min, dt_max = "2025-04-01", "2026-04-30"

    L = []
    w = L.append

    w("<!DOCTYPE html>")
    w('<html lang="pt-BR">')
    w("<head>")
    w('<meta charset="UTF-8">')
    w('<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=5">')
    w('<meta name="apple-mobile-web-app-capable" content="yes">')
    w("<title>Plantio 18M | Safra 25/26</title>")
    w('<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>')
    w('<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0"></script>')
    
    w("<style>")
    w("*{box-sizing:border-box;margin:0;padding:0}")
    w("body{font-family:Arial,sans-serif;font-size:13px;background:#ECEFF1;color:#212121}")
    w(".tabs{display:flex;background:#1B5E20;overflow-x:auto}")
    w(".tab{padding:13px 26px;color:rgba(255,255,255,.6);cursor:pointer;font-size:13px;font-weight:bold;border-bottom:3px solid transparent;white-space:nowrap}")
    w(".tab.ativo{color:#fff;border-bottom-color:#A5D6A7}.tab:hover{color:#fff}")
    w(".pg{display:none}.pg.ativo{display:block}")
    w(".wrap{max-width:1400px;margin:0 auto;padding:14px}")
    w(".hdr{background:#1B5E20;color:#fff;padding:13px 20px;font-size:15px;font-weight:bold}")
    w(".filtro{background:#E3F2FD;border:1px solid #90CAF9;border-radius:8px;padding:12px 14px;margin-bottom:12px;display:flex;flex-wrap:wrap;gap:10px;align-items:center}")
    w(".filtro strong{color:#1565C0}")
    w(".filtro label{color:#1565C0;font-size:12px;display:flex;align-items:center;gap:5px}")
    w(".filtro input[type=date]{border:2px solid #1565C0;border-radius:4px;padding:5px 8px;font-size:13px;color:#1565C0;background:#fff}")
    w(".filtro select{border:2px solid #2E7D32;border-radius:4px;padding:5px 8px;font-size:12px;color:#1B5E20;background:#fff}")
    w(".btn{border:none;border-radius:4px;padding:8px 16px;cursor:pointer;font-size:12px;font-weight:bold;color:#fff}")
    w(".btn-az{background:#1565C0}.btn-ci{background:#546E7A}")
    w(".info{color:#546E7A;font-size:11px}")
    w(".kpis{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:12px}")
    w(".kpi{background:#fff;border-radius:8px;padding:14px 16px;border-top:4px solid #43A047;flex:1 1 140px;min-width:120px}")
    w(".kpi.az{border-top-color:#1565C0}.kpi.la{border-top-color:#E65100}.kpi.ro{border-top-color:#6A1B9A}.kpi.am{border-top-color:#F57F17}")
    w(".kpi-lb{font-size:10px;color:#888;margin-bottom:5px;text-transform:uppercase}")
    w(".kpi-vl{font-size:21px;font-weight:bold;color:#1B5E20}")
    w(".kpi-sub{font-size:10px;color:#999;margin-top:2px}")
    w(".kpi.az .kpi-vl{color:#1565C0}.kpi.la .kpi-vl{color:#E65100}.kpi.ro .kpi-vl{color:#6A1B9A}.kpi.am .kpi-vl{color:#F57F17}")
    w(".grid{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:10px}")
    w(".card{background:#fff;border-radius:8px;padding:12px;overflow:auto;flex:1 1 300px;min-width:260px}")
    w(".card-full{background:#fff;border-radius:8px;padding:12px;margin-bottom:10px;overflow-x:auto}")
    w("h2{background:#2E7D32;color:#fff;padding:7px 12px;font-size:12px;border-radius:4px;margin-bottom:8px}")
    w("table{width:100%;border-collapse:collapse;font-size:12px}")
    w("th{background:#2E7D32;color:#fff;padding:9px 8px;text-align:left;white-space:nowrap;position:sticky;top:0;z-index:20;box-shadow:0 2px 4px rgba(0,0,0,0.1)}")
    w("td{padding:8px;border-bottom:1px solid #ECEFF1}")
    w("tr:hover td{background:#F1F8E9}")
    w(".num{text-align:right}")
    w(".viv td{background:#FFF9C4 !important;color:#E65100 !important}")
    w(".bw{background:#E8F5E9;border-radius:4px;height:7px;overflow:hidden;min-width:40px}")
    w(".bb{background:#43A047;height:7px}")
    w(".nlctrl{display:flex;gap:6px;margin-bottom:10px;flex-wrap:wrap}")
    w(".nlctrl button{background:#ECEFF1;border:1px solid #B0BEC5;border-radius:4px;padding:6px 12px;cursor:pointer;font-size:12px}")
    w(".nlctrl button.ativo{background:#2E7D32;color:#fff}")
    w(".faz td{background:#C8E6C9;font-weight:bold;color:#1B5E20}")
    w(".faz:hover td{background:#A5D6A7}")
    w(".zon td{background:#E8EAF6;color:#283593}")
    w(".zon:hover td{background:#C5CAE9}")
    w(".tal td{background:#FAFAFA;color:#546E7A}")
    w(".tal.viv td{background:#FFF9C4;color:#E65100}")
    w(".i0{padding-left:8px}.i1{padding-left:28px}.i2{padding-left:50px}")
    w(".ec{width:26px}")
    w(".eb{background:#fff;border:1px solid #90A4AE;border-radius:3px;width:22px;height:22px;cursor:pointer;font-size:11px;padding:0;color:#546E7A;line-height:22px;text-align:center}")
    w(".busca input{border:1px solid #B0BEC5;border-radius:4px;padding:7px 12px;font-size:12px;width:100%;max-width:360px;margin-bottom:10px}")
    w(".badge{display:inline-block;padding:1px 7px;border-radius:10px;font-size:10px;font-weight:bold;margin-left:4px}")
    w(".b18{background:#C8E6C9;color:#1B5E20}.b12{background:#BBDEFB;color:#0D47A1}.biv{background:#F3E5F5;color:#6A1B9A}")
    w(".rodape{text-align:center;color:#78909c;font-size:11px;padding:12px 0;border-top:1px solid #cfd8dc;margin-top:15px}")
    w(".ok-cell{background-color:#43A047 !important;color:#ffffff !important;font-weight:bold;text-align:center;border-radius:4px}")
    w(".warn-cell{color:#E65100 !important;font-weight:bold}")
    w(".saldo-pos{color:#1B5E20;font-weight:bold}")
    w(".saldo-zero{color:#d32f2f;font-weight:bold}")
    w(".saldo-kpi{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:12px}")
    w(".saldo-kpi .kpi{flex:1 1 180px}")
    w("@media(max-width:600px){.tab{padding:11px 14px;font-size:12px}.kpi-vl{font-size:17px}}")
    w("</style></head><body>")

    w('<div class="tabs">')
    w('  <div class="tab ativo" onclick="aba(\'dash\',this)">Dashboard</div>')
    w('  <div class="tab" onclick="aba(\'faz\',this)">Controle por Fazenda</div>')
    w('  <div class="tab" onclick="aba(\'ins\',this)">Status Insumos</div>')
    w('  <div class="tab" onclick="aba(\'saldo\',this)">Saldo Insumos</div>')
    w("</div>")

    # DASHBOARD
    w('<div class="pg ativo" id="pg-dash">')
    w('<div class="hdr">Acompanhamento Plantio 18M | Safra 25/26</div>')
    w('<div class="wrap">')
    w('<div class="filtro">')
    w('  <strong>Safra:</strong>')
    w('  <select id="selSafra" onchange="aplicarRefresh()" style="border:2px solid #6A1B9A;color:#6A1B9A;border-radius:4px;padding:5px 8px;font-size:12px;font-weight:bold;background:#fff">')
    for s in safras:
        sel = ' selected' if s == safras[0] else ''
        w(f'    <option value="{s}"{sel}>Safra {s}</option>')
    w('  </select>')
    w('  <strong>Periodo:</strong>')
    w(f'  <label>De <input type="date" id="dtI" value="{dt_min}"></label>')
    w(f'  <label>Ate <input type="date" id="dtF" value="{dt_max}"></label>')
    w('  <label>Ciclo:<select id="selC"><option value="">Todos</option><option value="18 Meses">18 Meses</option><option value="12 Meses">12 Meses</option><option value="Inverno">Inverno</option></select></label>')
    w('  <button class="btn btn-az" onclick="aplicar()">Aplicar</button>')
    w('  <button class="btn btn-ci" onclick="verTudo()">Ver tudo</button>')
    w('  <span class="info" id="inf"></span>')
    w('  <span class="info">Gerado ' + gerado_em + '</span>')
    w('</div>')
    w('<div class="kpis">')
    w('  <div class="kpi az"><div class="kpi-lb">Meta</div><div class="kpi-vl" id="kM">-</div></div>')
    w('  <div class="kpi"><div class="kpi-lb">Plantado no periodo</div><div class="kpi-vl" id="kP">-</div><div class="kpi-sub" id="kPsub"></div></div>')
    w('  <div class="kpi la"><div class="kpi-lb">Saldo pendente</div><div class="kpi-vl" id="kS">-</div></div>')
    w('  <div class="kpi ro"><div class="kpi-lb">Progresso</div><div class="kpi-vl" id="kG">-</div></div>')
    w('  <div class="kpi am"><div class="kpi-lb">18 Meses</div><div class="kpi-vl" id="k18">-</div></div>')
    w('  <div class="kpi am"><div class="kpi-lb">12 Meses</div><div class="kpi-vl" id="k12">-</div></div>')
    w('</div>')
    w('<div class="grid">')
    w('  <div class="card"><h2>Evolucao mensal (ha)</h2><table id="tM"><tr><td>Carregando...</td></tr></table></div>')
    w('  <div class="card"><h2>Top variedades</h2><table id="tV"><tr><td>Carregando...</td></tr></table></div>')
    w('</div>')
    w('<div class="card-full"><h2>Acompanhamento Diário - Meta vs Realizado</h2>')
    w('  <div id="erroMetas" style="display:none;color:#d32f2f;background:#ffebee;padding:15px;border-radius:8px;margin-bottom:10px;border:1px solid #ef5350;font-size:13px"></div>')
    w('  <div id="containerMetas" style="position:relative;height:320px;width:100%"><canvas id="graficoMetas"></canvas></div>')
    w('  <div style="margin-top:10px;max-height:250px;overflow-y:auto"><table id="tMetaDia"><tr><td>Carregando...</td></tr></table></div>')
    w('</div>')
    w('<div class="card-full"><h2>Plantio por ciclo</h2><table id="tC"><tr><td>Carregando...</td></tr></table></div>')
    w('<div class="rodape">Painel de Acompanhamento | Safra 25/26</div>')
    w('</div></div>')

    # FAZENDAS
    w('<div class="pg" id="pg-faz">')
    w('<div class="hdr">Controle por Fazenda | Fazenda &gt; Zona &gt; Talhao</div>')
    w('<div class="wrap">')
    w('<div class="busca"><input type="text" id="busca" placeholder="Buscar fazenda..." oninput="filtrarFaz()"></div>')
    w('<div class="nlctrl"><button id="b0" class="ativo" onclick="nivel(0)">So fazendas</button><button id="b1" onclick="nivel(1)">+ Zonas</button><button id="b2" onclick="nivel(2)">Tudo</button></div>')
    w('<div class="info" id="faz-info" style="margin-bottom:8px"></div>')
    w('<div style="overflow-x:auto"><table id="tD"><tr><th class="ec"></th><th>Fazenda / Zona / Talhao</th><th>Area (ha)</th><th>Variedades</th><th>Zonas/Talh.</th><th>Status</th></tr></table></div>')
    w('<div class="rodape">Gerado em ' + gerado_em + '</div>')
    w('</div></div>')

    # INSUMOS
    w('<div class="pg" id="pg-ins">')
    w('<div class="hdr">Status de Insumos por Talhão</div>')
    w('<div class="wrap">')
    w('<div class="busca"><input type="text" id="buscaIns" placeholder="Buscar por fazenda, talhão, ou status..." oninput="renderInsumos()"></div>')
    w('<div class="card-full" style="padding:0;max-height:70vh;overflow-y:auto;border:1px solid #ced4da"><table id="tIns"><tr><td>Carregando dados de insumos...</td></tr></table></div>')
    w('<div class="rodape">Gerado em ' + gerado_em + '</div>')
    w('</div></div>')

    # SALDO
    w('<div class="pg" id="pg-saldo">')
    w('<div class="hdr">Saldo de Insumos — Requisições e Movimentações</div>')
    w('<div class="wrap">')
    w('<div class="saldo-kpi" id="saldoKpis"></div>')
    w('<div class="busca"><input type="text" id="buscaSaldo" placeholder="Buscar insumo, requisição, material..." oninput="renderSaldo()"></div>')
    w('<div class="card-full" style="padding:0;max-height:70vh;overflow-y:auto;border:1px solid #ced4da"><table id="tSaldo"><tr><td>Carregando saldo de insumos...</td></tr></table></div>')
    w('<div class="rodape">Gerado em ' + gerado_em + '</div>')
    w('</div></div>')

    # JSON DATA
    w(f'<script type="application/json" id="data-reg">{reg_json}</script>')
    w(f'<script type="application/json" id="data-hier">{hier_json}</script>')
    w(f'<script type="application/json" id="data-metas">{metas_json}</script>')
    w(f'<script type="application/json" id="data-insumos">{insumos_json}</script>')
    w(f'<script type="application/json" id="data-saldo">{saldo_json}</script>')

    w("<script>")
    w("var META=" + str(int(meta_default)) + ";")
    w("var METAS_SAFRA=" + metas_safra_json + ";")
    w("var SAFRAS=" + safras_json + ";")
    w(f'var DT_MIN="{dt_min}";')
    w(f'var DT_MAX="{dt_max}";')
    w(f'var ERRO_METAS="{erro_metas}";')
    w("var REG=[],HIER={},METAS=[],INSUMOS=[],SALDO=[],chartMetas=null;")

    # ══════════════════════════════════════════════════════════
    # JAVASCRIPT COMPLETO
    # ══════════════════════════════════════════════════════════
    w(r"""
function aba(id,el){
  var pgs=document.getElementsByClassName("pg");
  for(var i=0;i<pgs.length;i++)pgs[i].className="pg";
  document.getElementById("pg-"+id).className="pg ativo";
  var tabs=document.getElementsByClassName("tab");
  for(var i=0;i<tabs.length;i++)tabs[i].className="tab";
  el.className="tab ativo";
}
function fmt(n,d){
  if(isNaN(n)||n===null)return"0";
  d=d===undefined?2:d;
  var s=Math.abs(n).toFixed(d),p=s.split("."),i=p[0],r="",c=0;
  for(var j=i.length-1;j>=0;j--){if(c>0&&c%3===0)r="."+r;r=i[j]+r;c++;}
  return(n<0?"-":"")+r+(d>0?","+p[1]:"");
}
function inArr(a,v){for(var i=0;i<a.length;i++)if(a[i]===v)return true;return false;}
function up(el,tag){while(el){if(el.tagName&&el.tagName.toUpperCase()===tag.toUpperCase())return el;el=el.parentElement;}return null;}

function renderGraficoMetas(dadosCompletos) {
  if (typeof Chart === 'undefined') return;
  var ctx = document.getElementById('graficoMetas');
  if (!ctx) return;
  ctx = ctx.getContext('2d');
  if (chartMetas) chartMetas.destroy();

  var labels=[], rData=[], mrData=[], mcData=[], nData=[];
  for (var i=0; i<dadosCompletos.length; i++) {
    var m = dadosCompletos[i];
    if (!m.dia || m.dia==='nan') continue;
    var dtParts = m.d.split("-");
    labels.push([dtParts[2]+"/"+dtParts[1], m.sem]);
    rData.push(m.r!==null && m.r>0 ? m.r : null);
    mrData.push(m.mr!==null ? m.mr : null);
    mcData.push(m.mc!==null ? m.mc : null);
    nData.push(m.n!==null ? m.n : null);
  }

  if (typeof ChartDataLabels !== 'undefined') Chart.register(ChartDataLabels);

  // Fechamento robusto para evitar que o escopo se perca na renderização do Chart.js
  var mostrarSoNoFinal = function(context) {
    var d = context.dataset.data;
    var lastIndex = -1;
    for (var j = d.length - 1; j >= 0; j--) {
      if (d[j] !== null && d[j] !== undefined && d[j] !== 0 && d[j] !== '') {
        lastIndex = j;
        break;
      }
    }
    return context.dataIndex === lastIndex;
  };

  chartMetas = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        {
          label: 'Realizado (ha)', data: rData, backgroundColor: '#43A047', order: 4, maxBarThickness: 40,
          datalabels: {
            anchor:'end', align:'top', color:'#1B5E20', font:{weight:'bold',size:10},
            display: function(ctx) { return ctx.dataset.data[ctx.dataIndex] !== null && ctx.dataset.data[ctx.dataIndex] > 0; }
          }
        },
        {
          label: 'Meta Real', data: mrData, type:'line', borderColor:'#1565C0', backgroundColor:'#1565C0',
          borderWidth:2, pointRadius:0, fill:false, order:3, spanGaps:true,
          datalabels: {
            anchor:'end', align:'right', color:'#1565C0', font:{weight:'bold',size:11},
            display: mostrarSoNoFinal
          }
        },
        {
          label: 'Nova Meta', data: nData, type:'line', borderColor:'#d32f2f', backgroundColor:'#d32f2f',
          borderWidth:2, borderDash:[5,5], pointRadius:0, fill:false, order:2, spanGaps:true,
          datalabels: {
            anchor:'end', align:'right', color:'#d32f2f', font:{weight:'bold',size:11},
            display: mostrarSoNoFinal
          }
        },
        {
          label: 'Meta Corrigida', data: mcData, type:'line', borderColor:'#FBC02D', backgroundColor:'#FBC02D',
          borderWidth:2, pointRadius:0, fill:false, order:1, spanGaps:true,
          datalabels: {
            anchor:'end', align:'right', color:'#F57F17', font:{weight:'bold',size:11},
            display: mostrarSoNoFinal
          }
        }
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false, layout:{padding:{top:25,right:45}},
      scales: {
        y:{beginAtZero:true,grid:{color:'#e0e0e0'}},
        x:{grid:{display:false},ticks:{font:{size:10}}}
      },
      plugins: {
        legend:{position:'bottom',labels:{usePointStyle:true,boxWidth:8}},
        datalabels:{formatter:function(v){if(!v)return'';return fmt(v,2);}}
      }
    }
  });
}

function renderTabelaMetas(dados) {
  var h="<tr><th>Data</th><th>Dia</th><th>Dia Sem.</th><th class='num'>Realizado</th><th class='num'>Meta Real</th><th class='num'>Meta Corrigida</th><th class='num'>Nova Meta</th><th class='num'>Desvio</th></tr>";
  for (var i=0; i<dados.length; i++) {
    var m=dados[i];
    if (!m.dia || m.dia==='nan') continue;
    var dtParts=m.d.split("-");
    var dataFmt=dtParts[2]+"/"+dtParts[1]+"/"+dtParts[0];
    var realizado=m.r!==null?fmt(m.r,2):"-";
    var metaReal=m.mr!==null?fmt(m.mr,2):"-";
    var metaCorr=m.mc!==null?fmt(m.mc,2):"-";
    var novaMeta=m.n!==null?fmt(m.n,2):"-";
    var desvio=m.dm!==null?fmt(m.dm,2):"-";
    var cls="";
    if(m.r!==null && m.r>0 && m.mr!==null) {
      cls = m.r >= m.mr ? "style='background:#E8F5E9'" : "style='background:#FFF3E0'";
    }
    h+="<tr "+cls+"><td>"+dataFmt+"</td><td>"+m.dia+"</td><td>"+m.sem+"</td>";
    h+="<td class='num'>"+realizado+"</td><td class='num'>"+metaReal+"</td>";
    h+="<td class='num'>"+metaCorr+"</td><td class='num'>"+novaMeta+"</td>";
    h+="<td class='num'>"+desvio+"</td></tr>";
  }
  document.getElementById("tMetaDia").innerHTML=h;
}

function aplicar(){
  var di=document.getElementById("dtI").value;
  var df=document.getElementById("dtF").value;
  var ci=document.getElementById("selC").value;
  var selSafra=document.getElementById("selSafra");
  var safra=selSafra?selSafra.value:"";
  if(safra && METAS_SAFRA[safra]) META=METAS_SAFRA[safra];
  if(!REG) return;
  if(!di||!df){alert("Preencha as duas datas.");return;}
  if(di>df){alert("Data inicio maior que data fim.");return;}
  var dados=[],p18=0,p12=0,piv=0;
  for(var i=0;i<REG.length;i++){
    var r=REG[i];
    if(r.d>=di&&r.d<=df&&(!ci||r.ci===ci)&&(!safra||r.s===safra))dados.push(r);
  }
  for(var i=0;i<dados.length;i++){
    if(dados[i].ci==="18 Meses")p18+=dados[i].a;
    else if(dados[i].ci==="12 Meses")p12+=dados[i].a;
    else piv+=dados[i].a;
  }
  var pl=p18+p12+piv,sl=META-pl,pg=META>0?pl/META*100:0;
  document.getElementById("kM").textContent=fmt(META,0)+" ha";
  document.getElementById("kP").textContent=fmt(pl,1)+" ha";
  document.getElementById("kPsub").textContent=ci?"Ciclo: "+ci:"Todos os ciclos";
  document.getElementById("kS").textContent=fmt(sl,1)+" ha";
  document.getElementById("kG").textContent=fmt(pg,1)+"%";
  document.getElementById("k18").textContent=fmt(p18,1)+" ha";
  document.getElementById("k12").textContent=fmt(p12,1)+" ha";
  rMeses(dados,pl);rVars(dados,pl);rCiclos(dados);
  var divErro=document.getElementById("erroMetas");
  var divCanvas=document.getElementById("containerMetas");
  if(ERRO_METAS){
    divErro.style.display="block";divErro.innerHTML="<b>Erro:</b> "+ERRO_METAS;divCanvas.style.display="none";
  } else if(METAS&&METAS.length>0){
    divErro.style.display="none";divCanvas.style.display="block";
    renderGraficoMetas(METAS);renderTabelaMetas(METAS);
  }
}

function verTudo(){document.getElementById("dtI").value=DT_MIN;document.getElementById("dtF").value=DT_MAX;document.getElementById("selC").value="";aplicar();}

function aplicarRefresh(){aplicar();renderFaz();nivel(0);}

function rMeses(dados,pl){
  var m={},keys=[];
  for(var i=0;i<dados.length;i++){var k=dados[i].d.substring(0,7);if(!m[k]){m[k]=0;keys.push(k);}m[k]+=dados[i].a;}
  keys.sort();var ac=0;
  var h="<tr><th>Mes</th><th>Area (ha)</th><th>Acumulado</th><th>%</th></tr>";
  for(var i=0;i<keys.length;i++){var k=keys[i],a=m[k];ac+=a;var pct=pl>0?a/pl*100:0,pts=k.split("-"),larg=Math.min(pct,100).toFixed(0);
    h+="<tr><td>"+pts[1]+"/"+pts[0]+"</td><td class='num'>"+fmt(a)+"</td><td class='num'>"+fmt(ac)+"</td>";
    h+="<td><div class='bw'><div class='bb' style='width:"+larg+"%'></div></div> "+fmt(pct,1)+"%</td></tr>";}
  document.getElementById("tM").innerHTML=h;
}
function rVars(dados,pl){
  var v={},vk=[];
  for(var i=0;i<dados.length;i++){var k=dados[i].v;if(!v[k]){v[k]=0;vk.push(k);}v[k]+=dados[i].a;}
  vk.sort(function(a,b){return v[b]-v[a];});
  var h="<tr><th>Variedade</th><th>Area (ha)</th><th>%</th></tr>";
  var lim=Math.min(vk.length,12);
  for(var i=0;i<lim;i++){var k=vk[i],a=v[k],pct=pl>0?a/pl*100:0,iv=(k==="VIVEIRO");
    var larg=Math.min(pct,100).toFixed(0),cor=iv?"#F9A825":"#43A047";
    h+="<tr"+(iv?" class='viv'":"")+"><td>"+k+(iv?" !":"")+"</td><td class='num'>"+fmt(a)+"</td>";
    h+="<td><div class='bw'><div class='bb' style='width:"+larg+"%;background:"+cor+"'></div></div> "+fmt(pct,1)+"%</td></tr>";}
  document.getElementById("tV").innerHTML=h;
}
function rCiclos(dados){
  var tot={};for(var i=0;i<dados.length;i++){if(!tot[dados[i].ci])tot[dados[i].ci]=0;tot[dados[i].ci]+=dados[i].a;}
  var total=0;for(var k in tot)total+=tot[k];
  var cls={"18 Meses":"b18","12 Meses":"b12","Inverno":"biv"};
  var h="<tr><th>Ciclo</th><th>Area (ha)</th><th>%</th><th>Progresso</th></tr>";
  var gps=["18 Meses","12 Meses","Inverno"];
  for(var i=0;i<gps.length;i++){var ci=gps[i],a=tot[ci]||0,pct=total>0?a/total*100:0;
    h+="<tr><td><span class='badge "+cls[ci]+"'>"+ci+"</span></td><td class='num'>"+fmt(a)+"</td><td class='num'>"+fmt(pct,1)+"%</td>";
    h+="<td style='min-width:80px'><div class='bw'><div class='bb' style='width:"+Math.min(pct,100).toFixed(0)+"%'></div></div></td></tr>";}
  h+="<tr><td><b>TOTAL</b></td><td class='num'><b>"+fmt(total)+"</b></td><td class='num'>100%</td><td></td></tr>";
  document.getElementById("tC").innerHTML=h;
}

function renderFaz(){
  if(!HIER)return;
  var selSafra=document.getElementById("selSafra");
  var safra=selSafra?selSafra.value:"";
  var cods=[];for(var k in HIER){if(HIER.hasOwnProperty(k))cods.push(k);}
  cods.sort(function(a,b){return HIER[a].n.localeCompare(HIER[b].n);});
  var h="";
  for(var ci=0;ci<cods.length;ci++){
    var cod=cods[ci],faz=HIER[cod],af=0,vf=[],zks=[],nZ=0,nT=0;
    var fazTemDados=false;
    for(var zk in faz.z){if(!faz.z.hasOwnProperty(zk))continue;
      var zonaTemDados=false;
      for(var tk in faz.z[zk]){if(!faz.z[zk].hasOwnProperty(tk))continue;var t=faz.z[zk][tk];
        if(safra&&t.s&&t.s!==safra)continue;
        zonaTemDados=true;nT++;af+=t.a;
        for(var vi=0;vi<t.v.length;vi++)if(!inArr(vf,t.v[vi]))vf.push(t.v[vi]);}
      if(zonaTemDados){nZ++;zks.push(parseInt(zk));fazTemDados=true;}
    }
    if(!fazTemDados)continue;
    var iv=inArr(vf,"VIVEIRO");
    var vm=vf.slice(0,3).join(" / ")+(vf.length>3?" +"+(vf.length-3):"");
    h+="<tr class='faz"+(iv?" viv":"")+"' data-nivel='0' data-nome='"+faz.n.toLowerCase()+"'>";
    h+="<td class='ec'><button class='eb' data-exp='0' onclick='tog(this)'>&#9654;</button></td>";
    h+="<td class='i0'>&#127968; "+faz.n+"</td><td class='num'>"+fmt(af)+" ha</td>";
    h+="<td style='font-size:11px'>"+vm+"</td><td class='num' style='font-size:11px'>"+nZ+"z / "+nT+"t</td>";
    h+="<td style='font-size:11px'>"+(iv?"! VIV":"OK")+"</td></tr>";
    zks.sort(function(a,b){return a-b;});
    for(var zi=0;zi<zks.length;zi++){
      var zn=zks[zi],zh=faz.z[zn],az=0,vz=[],tks=[];
      for(var tk in zh){if(!zh.hasOwnProperty(tk))continue;var t=zh[tk];
        if(safra&&t.s&&t.s!==safra)continue;
        tks.push(t.t||parseInt(tk));az+=t.a;
        for(var vi=0;vi<t.v.length;vi++)if(!inArr(vz,t.v[vi]))vz.push(t.v[vi]);}
      if(tks.length===0)continue;
      var vzm=vz.slice(0,3).join(" / ")+(vz.length>3?" +"+(vz.length-3):"");
      h+="<tr class='zon' data-nivel='1' data-nome='"+faz.n.toLowerCase()+"' style='display:none'>";
      h+="<td class='ec'><button class='eb' data-exp='0' onclick='tog(this)'>&#9654;</button></td>";
      h+="<td class='i1'>&#128205; Zona "+zn+"</td><td class='num'>"+fmt(az)+" ha</td>";
      h+="<td style='font-size:11px'>"+vzm+"</td><td class='num' style='font-size:11px'>"+tks.length+"t</td><td></td></tr>";
      tks.sort(function(a,b){return a-b;});
      for(var ti=0;ti<tks.length;ti++){
        var tn=tks[ti];
        // Find the talhao entry
        var tEntry=null;
        for(var ttk in zh){if(zh[ttk].t===tn||parseInt(ttk)===tn){if(!safra||!zh[ttk].s||zh[ttk].s===safra){tEntry=zh[ttk];break;}}}
        if(!tEntry)continue;
        var ivt=inArr(tEntry.v,"VIVEIRO");
        var dt=tEntry.d?tEntry.d.split("-").reverse().join("/"):"";
        h+="<tr class='tal"+(ivt?" viv":"")+"' data-nivel='2' data-nome='"+faz.n.toLowerCase()+"' style='display:none'>";
        h+="<td class='ec'></td><td class='i2'>&#8627; Talhao "+tn+"</td><td class='num'>"+fmt(tEntry.a)+" ha</td>";
        h+="<td style='font-size:11px'>"+tEntry.v.join(" / ")+(ivt?" !":"")+"</td>";
        h+="<td class='num' style='font-size:11px'>"+dt+"</td><td style='font-size:11px'>"+(ivt?"! VIV":"")+"</td></tr>";
      }
    }
  }
  document.getElementById("tD").innerHTML=
    "<tr><th class='ec'></th><th>Fazenda / Zona / Talhao</th><th>Area (ha)</th><th>Variedades</th><th>Zonas/Talh.</th><th>Status</th></tr>"+h;
}
function filtrarFaz(){
  var q=document.getElementById("busca").value.toLowerCase().trim();
  var rows=document.getElementById("tD").getElementsByTagName("tr");
  for(var i=1;i<rows.length;i++){
    var nm=rows[i].getAttribute("data-nome")||"";
    var nv=parseInt(rows[i].getAttribute("data-nivel"));
    if(isNaN(nv))continue;
    if(!q){rows[i].style.display=nv===0?"":"none";}
    else{rows[i].style.display=nm.indexOf(q)>-1?"":"none";}
  }
}
function tog(btn){
  var tr=up(btn,"tr");if(!tr)return;
  var nv=parseInt(tr.getAttribute("data-nivel"));
  var rows=tr.parentElement.getElementsByTagName("tr");
  var isExp=btn.getAttribute("data-exp")==="1";
  if(isExp){btn.setAttribute("data-exp","0");btn.innerHTML="&#9654;";}
  else{btn.setAttribute("data-exp","1");btn.innerHTML="&#9660;";}
  var found=false;
  for(var i=0;i<rows.length;i++){
    if(rows[i]===tr){found=true;continue;}
    if(!found)continue;
    var n=parseInt(rows[i].getAttribute("data-nivel"));
    if(isNaN(n)||n<=nv)break;
    if(isExp){rows[i].style.display="none";var b=rows[i].getElementsByTagName("button")[0];if(b){b.setAttribute("data-exp","0");b.innerHTML="&#9654;";}}
    else if(n===nv+1){rows[i].style.display="";}
  }
}
function nivel(nv){
  var ids=["b0","b1","b2"];
  for(var i=0;i<ids.length;i++){var el=document.getElementById(ids[i]);if(el)el.className=i===nv?"ativo":"";}
  var rows=document.getElementById("tD").getElementsByTagName("tr");
  for(var i=0;i<rows.length;i++){
    var n=parseInt(rows[i].getAttribute("data-nivel"));if(isNaN(n))continue;
    rows[i].style.display=n<=nv?"":"none";
    var btn=rows[i].getElementsByTagName("button")[0];
    if(btn){if(n<nv){btn.setAttribute("data-exp","1");btn.innerHTML="&#9660;";}else{btn.setAttribute("data-exp","0");btn.innerHTML="&#9654;";}}
  }
}

function renderInsumos(){
  if(!INSUMOS||INSUMOS.length===0){document.getElementById("tIns").innerHTML="<tr><td style='padding:20px;text-align:center'>Nenhuma informação de insumos.</td></tr>";return;}
  var q=document.getElementById("buscaIns").value.toLowerCase().trim();
  var cols=Object.keys(INSUMOS[0]);
  var h="<thead><tr>";
  for(var i=0;i<cols.length;i++)h+="<th>"+cols[i]+"</th>";
  h+="</tr></thead><tbody>";
  for(var r=0;r<INSUMOS.length;r++){
    var row=INSUMOS[r],match=!q;
    if(q){for(var i=0;i<cols.length;i++){if(String(row[cols[i]]).toLowerCase().indexOf(q)>-1){match=true;break;}}}
    if(!match)continue;
    h+="<tr>";
    for(var i=0;i<cols.length;i++){
      var val=row[cols[i]],sv=String(val).trim(),cl=sv.toLowerCase();
      var cls="";
      if(sv==="✓"||cl==="ok")cls=" class='ok-cell'";
      else if(cl.indexOf("sem insumo")>-1||cl.indexOf("⚠")>-1)cls=" class='warn-cell'";
      else if(sv==="-")cls=" style='text-align:center;color:#B0BEC5'";
      if(typeof val==='number'&&!Number.isInteger(val))val=fmt(val,2);
      h+="<td"+cls+">"+val+"</td>";
    }
    h+="</tr>";
  }
  h+="</tbody>";
  document.getElementById("tIns").innerHTML=h;
}

function renderSaldo(){
  if(!SALDO||SALDO.length===0){
    document.getElementById("tSaldo").innerHTML="<tr><td style='padding:20px;text-align:center'>Nenhum dado de saldo de insumo.</td></tr>";
    document.getElementById("saldoKpis").innerHTML="";return;
  }
  var totReq=0, totUti=0, totDev=0, totSaldo=0, nItens=0;
  for(var i=0;i<SALDO.length;i++){
    var s=SALDO[i];
    totReq+=parseFloat(s["Qtde Requisitada"])||0;
    totUti+=parseFloat(s["Qtde Utilizada"])||0;
    totDev+=parseFloat(s["Qtde Devolvida"])||0;
    totSaldo+=parseFloat(s["Saldo"])||0;
    nItens++;
  }
  var kh="";
  kh+="<div class='kpi'><div class='kpi-lb'>Total Requisitado</div><div class='kpi-vl'>"+fmt(totReq,1)+"</div></div>";
  kh+="<div class='kpi az'><div class='kpi-lb'>Total Utilizado</div><div class='kpi-vl'>"+fmt(totUti,1)+"</div></div>";
  kh+="<div class='kpi la'><div class='kpi-lb'>Total Devolvido</div><div class='kpi-vl'>"+fmt(totDev,1)+"</div></div>";
  kh+="<div class='kpi ro'><div class='kpi-lb'>Saldo Disponível</div><div class='kpi-vl'>"+fmt(totSaldo,1)+"</div></div>";
  kh+="<div class='kpi am'><div class='kpi-lb'>Linhas</div><div class='kpi-vl'>"+nItens+"</div></div>";
  document.getElementById("saldoKpis").innerHTML=kh;
  var q=document.getElementById("buscaSaldo").value.toLowerCase().trim();
  var cols=Object.keys(SALDO[0]);
  var h="<thead><tr>";
  for(var i=0;i<cols.length;i++)h+="<th>"+cols[i]+"</th>";
  h+="</tr></thead><tbody>";
  for(var r=0;r<SALDO.length;r++){
    var row=SALDO[r],match=!q;
    if(q){for(var i=0;i<cols.length;i++){if(String(row[cols[i]]).toLowerCase().indexOf(q)>-1){match=true;break;}}}
    if(!match)continue;
    h+="<tr>";
    for(var i=0;i<cols.length;i++){
      var val=row[cols[i]],sv=String(val).trim();
      var cls="";
      if(cols[i]==="Saldo"){var sn=parseFloat(sv);if(!isNaN(sn)){cls=sn>0?" class='saldo-pos'":" class='saldo-zero'";val=fmt(sn,1);}}
      else if(typeof val==='number'&&!Number.isInteger(val)){val=fmt(val,2);}
      if(cols[i]==="Data"&&sv.length===10){var dp=sv.split("-");val=dp[2]+"/"+dp[1]+"/"+dp[0];}
      h+="<td"+cls+">"+val+"</td>";
    }
    h+="</tr>";
  }
  h+="</tbody>";
  document.getElementById("tSaldo").innerHTML=h;
}

window.addEventListener('DOMContentLoaded', function() {
  try {
    REG=JSON.parse(document.getElementById('data-reg').textContent||'[]');
    HIER=JSON.parse(document.getElementById('data-hier').textContent||'{}');
    METAS=JSON.parse(document.getElementById('data-metas').textContent||'[]');
    INSUMOS=JSON.parse(document.getElementById('data-insumos').textContent||'[]');
    SALDO=JSON.parse(document.getElementById('data-saldo').textContent||'[]');
    aplicar();
    renderFaz();
    nivel(0);
    renderInsumos();
    renderSaldo();
  } catch(e) {
    var d=document.getElementById("kM");if(d)d.textContent="Erro: "+e.message;
    console.error(e);
  }
});
""")

    w("</script></body></html>")

    # ── MONTAR HTML FINAL COM PROTEÇÃO ──
    html_final = "\n".join(L)

    try:
        from auth_dashboard import proteger_html, deve_proteger
        if deve_proteger():
            html_final = proteger_html(html_final)
    except ImportError:
        pass

    Path(ARQUIVO_HTML).write_text(html_final, encoding="utf-8")
    tam = round(Path(ARQUIVO_HTML).stat().st_size / 1024, 1)
    metas_str = " + ".join(f"{s}:{int(m)}" for s, m in metas_por_safra.items())
    print(f"  [OK] {ARQUIVO_HTML} ({tam} KB) | Metas: {metas_str} ha | Dashboard Pronto!")


# ═══════════════════════════════════════════════════════════
#  MOTOBOY AUTOMÁTICO
# ═══════════════════════════════════════════════════════════
def enviar_para_github():
    print("  [☁️] Subindo nova versão para a Nuvem (GitHub)...")
    try:
        os.system(f'git add "{ARQUIVO_HTML}"')
        mensagem = f'Atualizacao automatica: {time.strftime("%d/%m/%Y %H:%M:%S")}'
        os.system(f'git commit -m "{mensagem}"')
        os.system('git push')
        print("  [✅] Sucesso! O GitHub Pages já está atualizando o link.")
    except Exception as e:
        print(f"  [❌] Erro ao enviar para a nuvem: {e}")

def get_mod_times():
    """Retorna dict com timestamps de modificação dos arquivos de safra"""
    mods = {}
    for arquivo, safra in ARQUIVOS_SAFRA:
        if os.path.exists(arquivo):
            mods[arquivo] = os.path.getmtime(arquivo)
    return mods

def monitorar():
    pasta = os.path.dirname(os.path.abspath(__file__))
    os.chdir(pasta)
    print()
    print("=" * 55)
    print("  DASHBOARD PLANTIO MULTI-SAFRA | AUTOMATIZADO")
    print("=" * 55)
    print()

    # Verificar quais arquivos existem
    encontrados = []
    for arquivo, safra in ARQUIVOS_SAFRA:
        if os.path.exists(arquivo):
            print(f"  [OK] Safra {safra}: {arquivo}")
            encontrados.append(safra)
        else:
            print(f"  [--] Safra {safra}: {arquivo} (não encontrado)")

    if not encontrados:
        print(f"\n  [ERRO] Nenhuma planilha de safra encontrada!")
        print(f"  Verifique os nomes em ARQUIVOS_SAFRA no início do script.")
        sys.exit(1)

    print(f"\n  Safras ativas: {', '.join(encontrados)}")
    print("  Gerando dashboard inicial...")

    try:
        df, metas_safra, df_metas, err, df_faz, df_ins, df_sal, safras = ler_todas_safras()
        gerar_html(df, metas_safra, df_metas, err, df_faz, df_ins, df_sal, safras)
        enviar_para_github()
        if ABRIR_BROWSER:
            try:
                os.startfile(os.path.abspath(ARQUIVO_HTML))
            except:
                pass
    except Exception as e:
        print(f"  [ERRO] {e}")
        import traceback; traceback.print_exc()

    ultima_mods = get_mod_times()
    print("-" * 55)

    try:
        while True:
            time.sleep(INTERVALO_SEG)
            mods_atuais = get_mod_times()

            if mods_atuais != ultima_mods:
                ultima_mods = mods_atuais
                hora = time.strftime("%H:%M:%S")
                print(f"\n  [{hora}] Excel salvo! Gerando novos gráficos...")
                time.sleep(3)
                try:
                    df, metas_safra, df_metas, err, df_faz, df_ins, df_sal, safras = ler_todas_safras()
                    gerar_html(df, metas_safra, df_metas, err, df_faz, df_ins, df_sal, safras)
                    enviar_para_github()
                except Exception as e:
                    print(f"  [ERRO] {e}")
                print("-" * 55)
            else:
                print(f"\r  Monitorando {len(mods_atuais)} planilha(s)... {time.strftime('%H:%M:%S')}", end="", flush=True)
    except KeyboardInterrupt:
        print("\n  Monitor encerrado.")


if __name__ == "__main__":
    try:
        from auth_dashboard import gerenciar_usuarios
        if gerenciar_usuarios():
            exit()
    except ImportError:
        pass
    monitorar()