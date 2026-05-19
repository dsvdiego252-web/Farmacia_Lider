"""
atualizar_polo.py
=================
Lê a planilha Resumo_Farmacia_Lider.xlsx e atualiza os dados
embutidos no HTML do dashboard, preservando todo o layout/lógica intactos.

Uso:
    python atualizar_polo.py

Requer (instale uma vez):
    pip install pandas openpyxl gitpython

Configuração (edite as variáveis abaixo):
"""


# ── REGRAS PIS/COFINS ────────────────────────────────────────────────────────
import re as _re

def _gerar_pc_rules(planilha_ncm):
    """Lê planilha NCM e gera PC_RULES_SP, PC_META_SP, PC_RULES_PERF, PC_META_PERF."""
    def _clean_ncm(v):
        return _re.sub(r"[^\d]","",str(v)).strip()
    def _fmt_cst(v):
        try:
            s = str(v).strip()
            if s in ["-","","nan","NaN"]: return ""
            return str(int(float(s.replace(",",".")))).zfill(2)
        except: return ""
    def _proc_aba(df_pc, col_ncm, col_regime, col_cst_ent, col_cst_sai,
                  col_aliq_pis, col_aliq_cof, col_nat, col_nat_desc, col_base, col_obs, fonte):
        rules, meta = {}, {}
        for _, row in df_pc.iterrows():
            try:
                ncm = _clean_ncm(row[col_ncm])
                if not ncm or len(ncm) < 4: continue
                cst_ent = _fmt_cst(row[col_cst_ent]) if pd.notna(row[col_cst_ent]) else ""
                cst_sai = _fmt_cst(row[col_cst_sai]) if pd.notna(row[col_cst_sai]) else ""
                if not cst_ent and not cst_sai: continue
                rules[ncm] = {"cst_ent": cst_ent, "cst_sai": cst_sai, "fonte": fonte}
                meta[ncm] = {
                    "regime": str(row[col_regime]).strip() if pd.notna(row[col_regime]) else "",
                    "cst_ent": cst_ent, "cst_sai": cst_sai,
                    "aliq_pis": str(row[col_aliq_pis]).strip() if pd.notna(row[col_aliq_pis]) else "",
                    "aliq_cofins": str(row[col_aliq_cof]).strip() if pd.notna(row[col_aliq_cof]) else "",
                    "nat_receita": str(row[col_nat]).strip() if pd.notna(row[col_nat]) else "",
                    "nat_desc": str(row[col_nat_desc]).strip() if pd.notna(row[col_nat_desc]) else "",
                    "base_legal": str(row[col_base]).strip() if pd.notna(row[col_base]) else "",
                    "obs": str(row[col_obs]).strip() if pd.notna(row[col_obs]) else "",
                    "fonte": fonte
                }
            except: continue
        return rules, meta
    try:
        # PIS-COFINS SP
        df_sp = pd.read_excel(planilha_ncm, sheet_name="PIS-COFINS SP")
        df_sp.columns = ["ncm","desc","ipi","regime","cst_ent","cst_sai","aliq_pis","aliq_cofins","nat_receita","nat_desc","base_legal","obs"]
        rules_sp, meta_sp = _proc_aba(df_sp,"ncm","regime","cst_ent","cst_sai",
            "aliq_pis","aliq_cofins","nat_receita","nat_desc","base_legal","obs","PIS-COFINS SP")
        # PIS-COFINS Perfumaria
        df_pf = pd.read_excel(planilha_ncm, sheet_name="PIS-COFINS Perfumaria")
        rules_pf, meta_pf = _proc_aba(df_pf,
            "NCM / SH","Regime\nPIS/COFINS","CST Entrada\n(Compra)","CST Saída\n(Venda)",
            "Alíq. PIS\n(Saída)","Alíq. COFINS\n(Saída)","Nat. Receita\n(EFD)",
            "Descrição Natureza Receita","Base Legal",
            "Observações (Farmácia – Lucro Real)","PIS-COFINS Perfumaria")
        print(f"✅ Regras PIS/COFINS: {len(rules_sp)} SP + {len(rules_pf)} Perfumaria")
        return rules_sp, meta_sp, rules_pf, meta_pf
    except Exception as e:
        print(f"⚠️  Erro ao carregar planilha NCM: {e}")
        return None, None, None, None


# ── CONFIGURAÇÃO ──────────────────────────────────────────────────────────────
PLANILHA   = "Resumo_Farmacia_Lider.xlsx"   # planilha atualizada
HTML_INPUT = "index.html"                           # HTML atual (template)
HTML_OUTPUT = "index.html"                          # HTML gerado (pode ser igual)
GIT_PUSH   = True                                  # False = só gera o HTML
COMMIT_MSG = "auto: atualiza dados Farmacia Lider"
# ─────────────────────────────────────────────────────────────────────────────

import json, re, sys
from collections import defaultdict
from datetime import datetime
import pandas as pd

print("📊 Lendo planilha...")
df_s = pd.read_excel(PLANILHA, sheet_name="Saidas")
df_e = pd.read_excel(PLANILHA, sheet_name="Entradas")

# Normalizar datas
df_s["data_ent"] = pd.to_datetime(df_s["data_ent"])
df_e["data_ent"] = pd.to_datetime(df_e["data_ent"])
df_s["Periodo"]  = pd.to_datetime(df_s["Periodo"]).dt.to_period("M").astype(str)
df_e["Periodo"]  = pd.to_datetime(df_e["Periodo"]).dt.to_period("M").astype(str)
df_s["ncm_produto"] = df_s["ncm_produto"].astype(str).str.strip()
df_e["ncm_produto"] = df_e["ncm_produto"].astype(str).str.strip()
df_s["produto"] = df_s["produto"].astype(str).str.strip()
df_e["produto"] = df_e["produto"].astype(str).str.strip()

# ── HELPERS ───────────────────────────────────────────────────────────────────

def dominant(grp, col):
    """Retorna o valor dominante (pela maior soma de valor_liq) de uma coluna."""
    agg = grp.groupby(col)["valor_liq"].sum()
    return str(agg.idxmax()) if not agg.empty else ""

def fmt_cst(val, width=2):
    """Converte CST float/string para formato correto: '04', '001', etc."""
    try:
        v = str(int(float(val)))
        return v.zfill(width)
    except:
        return str(val).strip()

def fmt_date(ts):
    return ts.strftime("%Y-%m-%d") if pd.notna(ts) else ""

def sparse_encode(arr):
    """Compacta séries diárias em pares [índice, valor].
    O HTML expande automaticamente no navegador pela função exp().
    Isso reduz muito o tamanho do arquivo de dados."""
    out = []
    for i, v in enumerate(arr):
        try:
            rv = round(float(v or 0), 2)
        except Exception:
            rv = 0
        if rv != 0:
            out.append([i, rv])
    return out

def build_daily_map(df, date_index):
    """date_index: lista ordenada de datas únicas → posição.
    Chave: (produto, cfop) para preservar CFOPs diferentes do mesmo produto."""
    di = {d: i for i, d in enumerate(date_index)}
    out = defaultdict(lambda: defaultdict(lambda: {
        "liq": 0.0, "qty": 0.0, "icms": 0.0, "base": 0.0,
        "pis": 0.0, "cofins": 0.0
    }))
    for _, row in df.iterrows():
        d = row["data_ent"].date()
        if d not in di:
            continue
        idx = di[d]
        key = (str(row["produto"]), int(float(row["cfop"] or 0)))
        out[key][idx]["liq"]    += float(row["valor_liq"] or 0)
        out[key][idx]["qty"]    += float(row["quantidade"] or 0)
        out[key][idx]["icms"]   += float(row["valor_icms"] or 0)
        out[key][idx]["base"]   += float(row["base_icms"] or 0)
        out[key][idx]["pis"]    += float(row["valor_pis"] or 0)
        out[key][idx]["cofins"] += float(row["valor_cofins"] or 0)
    return out

def build_prods(df, date_index):
    """Constrói lista PRODS agrupando por produto+cfop para preservar todos os CFOPs."""
    n = len(date_index)
    daily_map = build_daily_map(df, date_index)

    prods_out = []
    for (prod_id, cfop_val), grp in df.groupby(["produto", "cfop"]):
        liq      = round(float(grp["valor_liq"].sum()), 2)
        qty      = round(float(grp["quantidade"].sum()), 2)
        base     = round(float(grp["base_icms"].sum()), 2)
        icms     = round(float(grp["valor_icms"].sum()), 2)
        base_pis = round(float(grp["base_pis"].sum()), 2)
        val_pis  = round(float(grp["valor_pis"].sum()), 2)
        val_cof  = round(float(grp["valor_cofins"].sum()), 2)
        aliq_icms  = float(dominant(grp, "aliq_icms") or 0)
        # Alíquota dominante: pegar a de maior valor (excluindo zero)
        _pis_nz = grp[grp["aliq_pis"] > 0]
        aliq_pis = float(_pis_nz["aliq_pis"].mode()[0]) if not _pis_nz.empty else float(dominant(grp, "aliq_pis") or 0)
        _cof_nz = grp[grp["aliq_cofins"] > 0]
        aliq_cofins = float(_cof_nz["aliq_cofins"].mode()[0]) if not _cof_nz.empty else float(dominant(grp, "aliq_cofins") or 0)
        cst_icms  = dominant(grp, "cst_icms")
        cst_pis   = dominant(grp, "cst_pis")
        cst_cofins= dominant(grp, "cst_cofins")
        cfop      = int(float(cfop_val or 0))
        ncm       = str(grp["ncm_produto"].mode()[0]) if not grp["ncm_produto"].empty else ""
        cod       = str(grp["produto"].iloc[0])
        nome      = str(grp["descricao_produto"].iloc[0]) if "descricao_produto" in grp.columns else str(prod_id)

        # daily sparse
        key = (str(prod_id), cfop)
        dm = daily_map[key]
        def sp(field, dm=dm):
            arr = [dm[i][field] for i in range(n)]
            return sparse_encode(arr)

        prods_out.append({
            "p": nome,
            "liq": liq, "qty": qty, "base": base, "icms": icms,
            "aliq": aliq_icms,
            "cst": fmt_cst(cst_icms, 3),
            "cfop": cfop, "ncm": ncm, "cod": cod,
            "base_pis": base_pis, "val_pis": val_pis, "val_cofins": val_cof,
            "aliq_pis": aliq_pis, "aliq_cofins": aliq_cofins,
            "cst_pis": fmt_cst(cst_pis, 2),
            "cst_cofins": fmt_cst(cst_cofins, 2),
            "daily":      sp("liq"),
            "dailyq":     sp("qty"),
            "dailyicms":  sp("icms"),
            "dailybase":  sp("base"),
            "dailypis":   sp("pis"),
            "dailycofins":sp("cofins"),
        })

    prods_out.sort(key=lambda x: -x["liq"])
    return prods_out

def build_gdaily(df, date_index):
    """Agrega valores diários globais."""
    n = len(date_index)
    di = {d: i for i, d in enumerate(date_index)}
    liq_arr    = [0.0] * n
    qty_arr    = [0.0] * n
    icms_arr   = [0.0] * n
    base_arr   = [0.0] * n
    pis_arr    = [0.0] * n
    cofins_arr = [0.0] * n
    for _, row in df.iterrows():
        d = row["data_ent"].date()
        if d not in di:
            continue
        i = di[d]
        liq_arr[i]    += float(row["valor_liq"] or 0)
        qty_arr[i]    += float(row["quantidade"] or 0)
        icms_arr[i]   += float(row["valor_icms"] or 0)
        base_arr[i]   += float(row["base_icms"] or 0)
        pis_arr[i]    += float(row["valor_pis"] or 0)
        cofins_arr[i] += float(row["valor_cofins"] or 0)
    return {
        "liq":    [round(v,2) for v in liq_arr],
        "qty":    [round(v,2) for v in qty_arr],
        "icms":   [round(v,2) for v in icms_arr],
        "base":   [round(v,2) for v in base_arr],
        "pis":    [round(v,2) for v in pis_arr],
        "cofins": [round(v,2) for v in cofins_arr],
    }

def build_detail_s(df):
    """DETAIL_S: [data, doc, produto, descricao, cfop, qty, valor_liq]"""
    rows = []
    for _, r in df.iterrows():
        rows.append([
            fmt_date(r["data_ent"]),
            str(r["entrada"]),
            str(r["produto"]),
            str(r["descricao_produto"]),
            int(r["cfop"]),
            round(float(r["quantidade"] or 0), 2),
            round(float(r["valor_liq"] or 0), 2),
        ])
    return rows

def build_detail_i(df):
    """DETAIL_I: [..., base_icms, aliq_icms, valor_icms, cst_icms, ncm]"""
    rows = []
    for _, r in df.iterrows():
        rows.append([
            fmt_date(r["data_ent"]),
            str(r["entrada"]),
            str(r["produto"]),
            str(r["descricao_produto"]),
            int(r["cfop"]),
            round(float(r["quantidade"] or 0), 2),
            round(float(r["valor_liq"] or 0), 2),
            round(float(r["base_icms"] or 0), 2),
            round(float(r["aliq_icms"] or 0), 2),
            round(float(r["valor_icms"] or 0), 2),
            fmt_cst(r["cst_icms"], 3),
            str(r["ncm_produto"]),
        ])
    return rows

def build_detail_ps(df):
    """DETAIL_PS: [..., base_pis, aliq_pis, val_pis, base_cofins, aliq_cofins, val_cofins, cst_pis, cst_cofins, ncm]"""
    rows = []
    for _, r in df.iterrows():
        rows.append([
            fmt_date(r["data_ent"]),
            str(r["entrada"]),
            str(r["produto"]),
            str(r["descricao_produto"]),
            int(r["cfop"]),
            round(float(r["quantidade"] or 0), 2),
            round(float(r["valor_liq"] or 0), 2),
            round(float(r["base_pis"] or 0), 2),
            round(float(r["aliq_pis"] or 0), 2),
            round(float(r["valor_pis"] or 0), 2),
            round(float(r["base_cofins"] or 0), 2),
            round(float(r["aliq_cofins"] or 0), 2),
            round(float(r["valor_cofins"] or 0), 2),
            fmt_cst(r["cst_pis"], 2),
            fmt_cst(r["cst_cofins"], 2),
            str(r["ncm_produto"]),
        ])
    return rows

def build_detail_e(df):
    """DETAIL_E: entradas com dados ICMS"""
    rows = []
    for _, r in df.iterrows():
        rows.append([
            fmt_date(r["data_ent"]),
            str(r["entrada"]),
            str(r["produto"]),
            str(r["descricao_produto"]),
            int(r["cfop"]),
            round(float(r["quantidade"] or 0), 2),
            round(float(r["valor_liq"] or 0), 2),
            round(float(r["base_icms"] or 0), 2),
            round(float(r["aliq_icms"] or 0), 2),
            round(float(r["valor_icms"] or 0), 2),
            fmt_cst(r["cst_icms"], 3),
            str(r["ncm_produto"]),
        ])
    return rows

def build_detail_pe(df):
    """DETAIL_PE: entradas PIS/COFINS"""
    rows = []
    for _, r in df.iterrows():
        rows.append([
            fmt_date(r["data_ent"]),
            str(r["entrada"]),
            str(r["produto"]),
            str(r["descricao_produto"]),
            int(r["cfop"]),
            round(float(r["quantidade"] or 0), 2),
            round(float(r["valor_liq"] or 0), 2),
            round(float(r["base_pis"] or 0), 2),
            round(float(r["aliq_pis"] or 0), 2),
            round(float(r["valor_pis"] or 0), 2),
            round(float(r["base_cofins"] or 0), 2),
            round(float(r["aliq_cofins"] or 0), 2),
            round(float(r["valor_cofins"] or 0), 2),
            fmt_cst(r["cst_pis"], 2),
            fmt_cst(r["cst_cofins"], 2),
            str(r["ncm_produto"]),
        ])
    return rows

def build_detail_en(df):
    """DETAIL_EN: entradas simplificado (sem dados ICMS/PIS)"""
    rows = []
    for _, r in df.iterrows():
        rows.append([
            fmt_date(r["data_ent"]),
            str(r["entrada"]),
            str(r["produto"]),
            str(r["descricao_produto"]),
            int(r["cfop"]),
            round(float(r["quantidade"] or 0), 2),
            round(float(r["valor_liq"] or 0), 2),
        ])
    return rows

# ── PROCESSAR SAIDAS ──────────────────────────────────────────────────────────
print("⚙️  Processando saídas...")
dates_s  = sorted(df_s["data_ent"].dt.date.unique())
dates_s_str = [d.strftime("%Y-%m-%d") for d in dates_s]
dl_s     = [d.strftime("%d/%m") for d in dates_s]
periods_s = sorted(df_s["Periodo"].unique())
cfops_s  = sorted(df_s["cfop"].dropna().astype(int).unique().tolist())
aliquotas_s = sorted(df_s["aliq_icms"].dropna().astype(float).unique().tolist())
aliquotas_s = [int(a) if a == int(a) else a for a in aliquotas_s]
ncms_s   = sorted(df_s["ncm_produto"].dropna().unique().tolist())

PRODS     = build_prods(df_s, dates_s)
GDAILY    = build_gdaily(df_s, dates_s)
DETAIL_S  = build_detail_s(df_s)
DETAIL_I  = build_detail_i(df_s)
DETAIL_PS = build_detail_ps(df_s)
N_DATES   = len(dates_s)

# ── PROCESSAR ENTRADAS ────────────────────────────────────────────────────────
print("⚙️  Processando entradas...")
dates_e  = sorted(df_e["data_ent"].dt.date.unique())
dates_e_str = [d.strftime("%Y-%m-%d") for d in dates_e]
dl_e     = [d.strftime("%d/%m") for d in dates_e]
periods_e = sorted(df_e["Periodo"].unique())
cfops_en = sorted(df_e["cfop"].dropna().astype(int).unique().tolist())
aliquotas_e = sorted(df_e["aliq_icms"].dropna().astype(float).unique().tolist())
aliquotas_e = [int(a) if a == int(a) else a for a in aliquotas_e]
ncms_e   = sorted(df_e["ncm_produto"].dropna().unique().tolist())

PRODS_E   = build_prods(df_e, dates_e)
GDAILY_E  = build_gdaily(df_e, dates_e)
PRODS_EN  = PRODS_E   # mesmos produtos, mesmo índice
GDAILY_EN = GDAILY_E
DETAIL_E  = build_detail_e(df_e)
DETAIL_PE = build_detail_pe(df_e)
DETAIL_EN = build_detail_en(df_e)
N_DATES_E = len(dates_e)

# ── SERIALIZAR ────────────────────────────────────────────────────────────────
def j(obj):
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))

def jset(lst):
    return f'new Set({j(lst)})'

new_lines = {
    # Saidas
    "PRODS_EN":   f"const PRODS_EN={j(PRODS_EN)};",
    "GDAILY_EN":  f"const GDAILY_EN={j(GDAILY_EN)};",
    "DATES_EN":   f"const DATES_EN={j(dates_e_str)};",
    "DL_EN":      f"const DL_EN={j(dl_e)};",
    "PERIODS_EN": f"const PERIODS_EN={j(periods_e)};",
    "CFOPS_EN":   f"const CFOPS_EN={j(cfops_en)};",
    "PRODS":      f"const PRODS={j(PRODS)};",
    "GDAILY":     f"const GDAILY={j(GDAILY)};",
    "ALIQUOTAS":  f"const ALIQUOTAS={j(aliquotas_s)};",
    "ALL_NCMS":   f"const ALL_NCMS={jset(ncms_s)};",
    "DATES":      f"const DATES={j(dates_s_str)};",
    "DL":         f"const DL={j(dl_s)};",
    "DETAIL_S":   f"const DETAIL_S={j(DETAIL_S)};",
    "PERIODS":    f"const PERIODS={j(periods_s)};",
    "PRODS_E":    f"const PRODS_E={j(PRODS_E)};",
    "GDAILY_E":   f"const GDAILY_E={j(GDAILY_E)};",
    "ALIQUOTAS_E":f"const ALIQUOTAS_E={j(aliquotas_e)};",
    "ALL_NCMS_E": f"const ALL_NCMS_E={jset(ncms_e)};",
    "DATES_E":    f"const DATES_E={j(dates_e_str)};",
    "DL_E":       f"const DL_E={j(dl_e)};",
    "PERIODS_E":  f"const PERIODS_E={j(periods_e)};",
    "DETAIL_I":   f"const DETAIL_I={j(DETAIL_I)};",
    "DETAIL_E":   f"const DETAIL_E={j(DETAIL_E)};",
    "DETAIL_EN":  f"const DETAIL_EN={j(DETAIL_EN)};",
    "DATES_PS":   f"const DATES_PS={j(dates_s_str)};",
    "PERIODS_PS": f"const PERIODS_PS={j(periods_s)};",
    "DETAIL_PS":  f"const DETAIL_PS={j(DETAIL_PS)};",
    # PRODS_PS/PRODS_PE são aliases necessários para compatibilidade com funções antigas do HTML.
    # A matriz corrigida usa DETAIL_PS/DETAIL_PE, mas manter os aliases evita erro se alguma rotina antiga chamar esses nomes.
    "PRODS_PS":   "const PRODS_PS=PRODS;",
    "DATES_PE":   f"const DATES_PE={j(dates_e_str)};",
    "PERIODS_PE": f"const PERIODS_PE={j(periods_e)};",
    "DETAIL_PE":  f"const DETAIL_PE={j(DETAIL_PE)};",
    # Mesmo motivo do PRODS_PS: compatibilidade com funções antigas do HTML.
    "PRODS_PE":   "const PRODS_PE=PRODS_E;",
    "N_DATES":    f"const N_DATES={N_DATES},N_DATES_E={N_DATES_E};",
}


# ── GERAR ARQUIVO DE DADOS SEPARADO ──────────────────────────────────────────
print("🧩 Gerando dados separados do HTML...")

import os as _os
from pathlib import Path as _Path

DADOS_DIR = _Path("dados")
DADOS_DIR.mkdir(exist_ok=True)
DADOS_OUTPUT = DADOS_DIR / "dashboard_dados.js"

RULE_VARS = ["ICMS_RULES", "PC_RULES_SP", "PC_META_SP", "PC_RULES_PERF", "PC_META_PERF"]

# Ordem estável dos dados no arquivo externo
DATA_ORDER = [
    "PRODS_EN","GDAILY_EN","DATES_EN","DL_EN","PERIODS_EN","CFOPS_EN",
    "PRODS","GDAILY","ALIQUOTAS","ALL_NCMS","DATES","DL","DETAIL_S","PERIODS",
    "PRODS_E","GDAILY_E","ALIQUOTAS_E","ALL_NCMS_E","DATES_E","DL_E","PERIODS_E",
    "DETAIL_I","DETAIL_E","DETAIL_EN",
    "PRODS_PS","DATES_PS","PERIODS_PS","DETAIL_PS",
    "PRODS_PE","DATES_PE","PERIODS_PE","DETAIL_PE",
    "N_DATES"
]

def _find_const_statement(txt, varname):
    """Retorna a linha/declaração const VAR=...; mesmo se for objeto grande."""
    m = re.search(r"const\s+" + re.escape(varname) + r"\s*=", txt)
    if not m:
        return None
    i = m.start()
    # Para objetos/arrays grandes, procura o ; final respeitando strings e chaves.
    depth_curly = depth_square = depth_par = 0
    in_str = False
    quote = ""
    esc = False
    for j in range(m.end(), len(txt)):
        c = txt[j]
        if in_str:
            if esc:
                esc = False
            elif c == "\\":
                esc = True
            elif c == quote:
                in_str = False
        else:
            if c in ("'", '"', "`"):
                in_str = True
                quote = c
            elif c == "{": depth_curly += 1
            elif c == "}": depth_curly -= 1
            elif c == "[": depth_square += 1
            elif c == "]": depth_square -= 1
            elif c == "(": depth_par += 1
            elif c == ")": depth_par -= 1
            elif c == ";" and depth_curly <= 0 and depth_square <= 0 and depth_par <= 0:
                return txt[i:j+1]
    return None

def _load_existing_rules():
    txt = ""
    if DADOS_OUTPUT.exists():
        txt = DADOS_OUTPUT.read_text(encoding="utf-8", errors="ignore")
    elif _Path(HTML_INPUT).exists():
        txt = _Path(HTML_INPUT).read_text(encoding="utf-8", errors="ignore")
    out = {}
    for v in RULE_VARS:
        stmt = _find_const_statement(txt, v)
        if stmt:
            out[v] = stmt
    return out

rules_lines = _load_existing_rules()

# ── ATUALIZAR REGRAS PIS/COFINS, SE EXISTIR PLANILHA NCM ─────────────────────
NCM_PLANILHA = "ncm_med_v8.xlsx"
if _os.path.exists(NCM_PLANILHA):
    print("📋 Atualizando regras PIS/COFINS...")
    _rules, _meta, _rules_pf, _meta_pf = _gerar_pc_rules(NCM_PLANILHA)
    if _rules and _meta:
        _rules_js = json.dumps(_rules, ensure_ascii=False, separators=(",",":"))
        _meta_js  = json.dumps(_meta,  ensure_ascii=False, separators=(",",":"))
        rules_lines["PC_RULES_SP"] = f"const PC_RULES_SP={_rules_js};"
        rules_lines["PC_META_SP"]  = f"const PC_META_SP={_meta_js};"
    if _rules_pf and _meta_pf:
        _rules_pf_js = json.dumps(_rules_pf, ensure_ascii=False, separators=(",",":"))
        _meta_pf_js  = json.dumps(_meta_pf,  ensure_ascii=False, separators=(",",":"))
        rules_lines["PC_RULES_PERF"] = f"const PC_RULES_PERF={_rules_pf_js};"
        rules_lines["PC_META_PERF"]  = f"const PC_META_PERF={_meta_pf_js};"
else:
    print(f"ℹ️  {NCM_PLANILHA} não encontrada - regras PIS/COFINS mantidas do template/dados")

missing_rules = [v for v in RULE_VARS if v not in rules_lines]
if missing_rules:
    print(f"⚠️  Atenção: regras não encontradas: {missing_rules}")
    print("   O dashboard pode funcionar parcialmente, mas a conferência fiscal pode ficar incompleta.")

_dados = []
_dados.append("// Dashboard Polo e Polo - dados separados")
_dados.append("// Gerado automaticamente pelo atualizar_polo.py")
_dados.append("// Não edite manualmente.\n")
for k in DATA_ORDER:
    if k in new_lines:
        _dados.append(new_lines[k])
    else:
        print(f"⚠️  Variável de dados não encontrada e foi ignorada: {k}")
_dados.append("\n// Regras fiscais")
for k in RULE_VARS:
    if k in rules_lines:
        _dados.append(rules_lines[k])

DADOS_OUTPUT.write_text("\n".join(_dados) + "\n", encoding="utf-8")
print(f"💾 Dados salvos em: {DADOS_OUTPUT}")

# ── GERAR resumo_dados.js (arquivo leve para carregamento rápido) ─────────────
print("📋 Gerando resumo_dados.js...")
import json as _json2
from datetime import datetime as _dt2

_periods_r = sorted(set(
    list(df_s["Periodo"].unique()) + list(df_e["Periodo"].unique())
))
_by_period_r = {}
for _p in _periods_r:
    _ds = df_s[df_s["Periodo"] == _p]
    _de = df_e[df_e["Periodo"] == _p]
    _icms_deb  = float(_ds["valor_icms"].sum())
    _icms_cred = float(_de["valor_icms"].sum())
    _pis_deb   = float(_ds["valor_pis"].sum())
    _pis_cred  = float(_de["valor_pis"].sum())
    _cof_deb   = float(_ds["valor_cofins"].sum())
    _cof_cred  = float(_de["valor_cofins"].sum())
    _by_period_r[_p] = {
        "vendas":          round(float(_ds["valor_liq"].sum()), 2),
        "compras":         round(float(_de["valor_liq"].sum()), 2),
        "icms":            round(max(0, _icms_deb - _icms_cred), 2),
        "icms_debito":     round(_icms_deb, 2),
        "icms_credito":    round(_icms_cred, 2),
        "pis_recolher":    round(max(0, _pis_deb - _pis_cred), 2),
        "pis_debito":      round(_pis_deb, 2),
        "pis_credito":     round(_pis_cred, 2),
        "cofins_recolher": round(max(0, _cof_deb - _cof_cred), 2),
        "cofins_debito":   round(_cof_deb, 2),
        "cofins_credito":  round(_cof_cred, 2),
    }
_mes_map_r = {"01":"Jan","02":"Fev","03":"Mar","04":"Abr","05":"Mai","06":"Jun",
              "07":"Jul","08":"Ago","09":"Set","10":"Out","11":"Nov","12":"Dez"}
_labels_r = [_mes_map_r.get(_p.split("-")[1], _p) for _p in _periods_r]
_all_r = {k: round(sum(_by_period_r[_p].get(k, 0) for _p in _periods_r), 2)
          for k in ["vendas","compras","icms","icms_debito","icms_credito",
                    "pis_recolher","pis_debito","pis_credito",
                    "cofins_recolher","cofins_debito","cofins_credito"]}
_resumo_r = {
    "periods":     _periods_r,
    "labels":      _labels_r,
    "all":         _all_r,
    "by_period":   _by_period_r,
    "generated_at": _dt2.now().strftime("%Y-%m-%dT%H:%M:%S")
}
RESUMO_OUTPUT = DADOS_DIR / "resumo_dados.js"
_resumo_js = "// Resumo inicial leve - gerado automaticamente pelo atualizar_polo.py\n"
_resumo_js += "var RESUMO_DATA=" + _json2.dumps(_resumo_r, ensure_ascii=False, separators=(",",":")) + ";\n"
RESUMO_OUTPUT.write_text(_resumo_js, encoding="utf-8")
print(f"💾 Resumo salvo em: {RESUMO_OUTPUT}")



# ── GERAR RT_DATA (Reforma Tributária) ──────────────────────────────────────
print("⚖️  Gerando dados da Reforma Tributária...")
TIPI_PLANILHA = "TIPI_cClassTrib_IBS_CBS_2026_v5.xlsx"
_rt_data = []
if _os.path.exists(TIPI_PLANILHA):
    try:
        _df_tipi = pd.read_excel(TIPI_PLANILHA, sheet_name="TIPI x cClassTrib", dtype={"NCM": str})
        _df_tipi["NCM"] = _df_tipi["NCM"].astype(str).str.strip().str.zfill(8)
        _df_tipi["Alíq. IBS (%)"] = pd.to_numeric(_df_tipi["Alíq. IBS (%)"], errors="coerce").fillna(0)
        _df_tipi["Alíq. CBS (%)"] = pd.to_numeric(_df_tipi["Alíq. CBS (%)"], errors="coerce").fillna(0)
        _df_tipi["Redução"] = _df_tipi["Redução"].astype(str).str.strip()
        for _, row in _df_tipi.iterrows():
            _rt_data.append([
                str(row["NCM"]),
                str(row.get("Descrição TIPI","")).strip()[:80],
                str(row.get("Produto / Situação","")).strip()[:80],
                str(row.get("cClassTrib","")).strip(),
                int(float(row.get("CST", 0) or 0)),
                str(row.get("Tipo de Tributação IBS/CBS","")).strip(),
                str(row.get("Redução","")).strip(),
                round(float(row["Alíq. IBS (%)"]), 4),
                round(float(row["Alíq. CBS (%)"]), 4),
                str(row.get("Base Legal","")).strip()[:120],
            ])
        print(f"✅ Reforma Tributária: {len(_rt_data)} produtos carregados")
    except Exception as _e:
        print(f"⚠️  Erro ao carregar TIPI: {_e}")
else:
    print(f"ℹ️  {TIPI_PLANILHA} não encontrada - aba Reforma Trib. ficará sem dados")

# ── GERAR NCM_BY_CONTEXT (NCMs dinâmicas por período/CST para filtros) ────────
print("🗂️  Gerando NCMs dinâmicas por contexto...")

def _build_ncm_context(df, detail_rows, col_period, col_cst, col_ncm_idx=11):
    """Gera mapa: period -> cst -> [ncms] e period -> 'all' -> [ncms]"""
    ctx = {}
    for row in detail_rows:
        # row: [data, doc, cod, produto, cfop, qty, liq, base, aliq, icms_val, cst, ncm, ...]
        try:
            period = row[0][:7]  # "2026-01"
            cst = str(row[10]) if len(row) > 10 else "all"
            ncm = str(row[col_ncm_idx]) if len(row) > col_ncm_idx else ""
            if not ncm or ncm in ("nan","None",""):
                continue
            if period not in ctx:
                ctx[period] = {}
            if "all" not in ctx[period]:
                ctx[period]["all"] = set()
            ctx[period]["all"].add(ncm)
            if cst not in ctx[period]:
                ctx[period][cst] = set()
            ctx[period][cst].add(ncm)
        except:
            continue
    # Converter sets para listas ordenadas
    return {p: {c: sorted(list(v)) for c, v in csts.items()} for p, csts in ctx.items()}

# DETAIL_I e DETAIL_E já estão em new_lines como strings JS
# Precisamos dos dados Python originais para gerar o contexto
# Usar df_s e df_e diretamente

def _build_ncm_ctx_from_df(df, cst_col, ncm_col="ncm_produto", period_col="Periodo"):
    ctx = {"all": {"all": set()}}
    for _, row in df.iterrows():
        try:
            period = str(row[period_col])
            cst = str(fmt_cst(row[cst_col], 3)) if cst_col == "cst_icms" else str(fmt_cst(row[cst_col], 2))
            ncm = str(row[ncm_col]).strip()
            if not ncm or ncm in ("nan","None",""):
                continue
            # Global
            ctx["all"]["all"].add(ncm)
            if "all" not in ctx.get(period, {}):
                if period not in ctx:
                    ctx[period] = {}
                ctx[period]["all"] = set()
            ctx[period]["all"].add(ncm)
            # Por CST
            if cst not in ctx["all"]:
                ctx["all"][cst] = set()
            ctx["all"][cst].add(ncm)
            if cst not in ctx[period]:
                ctx[period][cst] = set()
            ctx[period][cst].add(ncm)
        except:
            continue
    return {p: {c: sorted(list(v)) for c, v in csts.items()} for p, csts in ctx.items()}

# Para Saídas ICMS (cst_icms, aliq_icms, cfop)
_ncm_ctx_saidas = _build_ncm_ctx_from_df(df_s, "cst_icms")
# Para Entradas ICMS
_ncm_ctx_entradas = _build_ncm_ctx_from_df(df_e, "cst_icms")
# Para Saídas PIS/Cofins (cst_pis)
_ncm_ctx_ps = _build_ncm_ctx_from_df(df_s, "cst_pis")
# Para Entradas PIS/Cofins
_ncm_ctx_pe = _build_ncm_ctx_from_df(df_e, "cst_pis")

_ncm_ctx_js = json.dumps({
    "saidas":   _ncm_ctx_saidas,
    "entradas": _ncm_ctx_entradas,
    "ps":       _ncm_ctx_ps,
    "pe":       _ncm_ctx_pe,
}, ensure_ascii=False, separators=(",",":"))

# Adicionar RT_DATA e NCM_CTX ao arquivo de dados
_rt_js = json.dumps(_rt_data, ensure_ascii=False, separators=(",",":"))
with open(str(DADOS_OUTPUT), "a", encoding="utf-8") as _frt:
    _frt.write(f"\nconst RT_DATA={_rt_js};\n")
    _frt.write(f"const NCM_CTX={_ncm_ctx_js};\n")
print(f"✅ RT_DATA e NCM_CTX adicionados ao {DADOS_OUTPUT}")

# Atualizar DATA_ORDER para incluir no git add (já incluso via DADOS_OUTPUT)




# ── GARANTIR QUE O INDEX ESTEJA NA VERSÃO LEVE ───────────────────────────────
def _converter_index_para_dados_separados():
    p = _Path(HTML_INPUT)
    if not p.exists():
        return False
    html = p.read_text(encoding="utf-8", errors="ignore")
    if 'src="dados/dashboard_dados.js"' in html or "src='dados/dashboard_dados.js'" in html:
        return False
    # Se ainda tiver dados inline, remove as declarações grandes e injeta script externo.
    data_vars = set(DATA_ORDER + RULE_VARS)
    out = []
    inserted = False
    removed = 0
    for line in html.splitlines(True):
        m = re.match(r"^\s*const\s+([A-Z0-9_]+)\s*=", line)
        if m and m.group(1) in data_vars:
            removed += 1
            if not inserted:
                out.append('</script>\n')
                out.append('<script src="dados/dashboard_dados.js"></script>\n')
                out.append('<script>\n')
                inserted = True
            continue
        out.append(line)
    if inserted:
        p.write_text("".join(out), encoding="utf-8")
        print(f"✅ index.html convertido para versão leve ({removed} blocos de dados removidos).")
        return True
    return False

_index_alterado = _converter_index_para_dados_separados()

# ── GIT PUSH ──────────────────────────────────────────────────────────────────
if GIT_PUSH:
    try:
        import git
        repo = git.Repo(".")
        add_files = [str(DADOS_OUTPUT)]
        # Incluir resumo_dados.js
        _resumo_path = DADOS_DIR / "resumo_dados.js"
        if _resumo_path.exists():
            add_files.append(str(_resumo_path))
        # Incluir arquivos de aba
        for _aba_f in ["icms.js", "piscofins.js", "entradas.js"]:
            _aba_path = DADOS_DIR / _aba_f
            if _aba_path.exists():
                add_files.append(str(_aba_path))
        if _Path(HTML_OUTPUT).exists():
            add_files.append(HTML_OUTPUT)
        repo.index.add(add_files)
        try:
            repo.index.commit(COMMIT_MSG)
        except Exception as cex:
            print(f"ℹ️  Commit não criado ou sem alterações: {cex}")
        origin = repo.remote(name="origin")
        # Envia para main (branch do Vercel Production)
        push_info = origin.push(refspec="master:main", force=True)
        for info in push_info:
            if info.flags & info.ERROR:
                print(f"⚠️  Aviso no push: {info.summary}")
        print("🚀 Push realizado com sucesso!")
    except ImportError:
        print("⚠️  gitpython não instalado. Instale com: pip install gitpython")
        print("   Rode manualmente: git add index.html dados/dashboard_dados.js && git commit -m 'atualiza dados' && git push")
    except Exception as ex:
        print(f"⚠️  Erro no git push: {ex}")
        print("   Rode manualmente: git add index.html dados/dashboard_dados.js && git commit -m 'atualiza dados' && git push")

print("\n✅ Concluído!")
