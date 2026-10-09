#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
processar_ifood.py — Gerador de Dashboard BI para iFood (Falaw Advogados)
Uso: python processar_ifood.py "caminho/para/RELATÓRIO IFOOD.xlsx"
"""

import sys
import os
import json
import re
from datetime import datetime
from collections import Counter, defaultdict

import pandas as pd

# ─── CONFIG ────────────────────────────────────────────────────────────────────

DEFAULT_EXCEL = "RELATÓRIO IFOOD.xlsx"

CAUSA_RAIZ_TABS = [
    ("ENTREGADOR NUVEM",          "Nuvem"),
    ("ENTREGADOR OL",             "OL"),
    ("EX-FOODLOVER",              "Ex-Foodlover"),
    ("TERCEIRIZAÇÃO",             "Terceirização"),
    ("ENTREGADOR - MARKETPLACE",  "Marketplace"),
    ("RESTAURANTES - CONTRATO",   "Restaurantes"),
]

UF_TRT = {
    "MG": "TRT-3",  "RJ": "TRT-1",  "RS": "TRT-4",  "BA": "TRT-5",
    "PE": "TRT-6",  "CE": "TRT-7",  "PA": "TRT-8",  "AP": "TRT-8",
    "PR": "TRT-9",  "DF": "TRT-10", "TO": "TRT-10", "AM": "TRT-11",
    "RR": "TRT-11", "SC": "TRT-12", "PB": "TRT-13", "RO": "TRT-14",
    "AC": "TRT-14", "MA": "TRT-16", "ES": "TRT-17", "GO": "TRT-18",
    "AL": "TRT-19", "SE": "TRT-20", "RN": "TRT-21", "PI": "TRT-22",
    "MT": "TRT-23", "MS": "TRT-24",
}

# SP interior cities (TRT-15)
SP_TRT15_KEYWORDS = [
    "campinas", "jundiaí", "jundiai", "sorocaba", "ribeirão preto", "ribeirao preto",
    "são josé dos campos", "sao jose dos campos", "santo andré", "santo andre",
    "piracicaba", "limeira", "bauru", "araraquara", "franca", "americana",
    "marília", "marilia", "presidente prudente", "são carlos", "sao carlos",
    "taubaté", "taubate", "araçatuba", "aracatuba", "botucatu", "catanduva",
    "guaratinguetá", "guaratingueta", "indaiatuba", "itapeva", "jaboticabal",
    "jaú", "jau", "mogi das cruzes", "ourinhos", "pindamonhangaba",
    "são bernardo", "sao bernardo", "osasco", "guarulhos", "mauá", "maua",
    "carapicuíba", "carapicuiba", "barueri", "praia grande", "santos",
    "são vicente", "sao vicente", "trt-15", "trt15",
]

FAVORAVEL_VALUES   = {"Improcedente", "Extinto sem Resolução"}
DESFAVORAVEL_VALUES = {"Procedente", "Parc. Procedente"}
NEUTRO_VALUES       = {"Acordo", "Sem Sentença Expressa"}

DECISAO_FAVORAVEL_VALUES   = {"Acórdão Favorável", "Sentença Improcedente", "Decisão Favorável TST"}
DECISAO_DESFAVORAVEL_VALUES = {"Acórdão Desfavorável", "Sentença Desfavorável", "Acórdão Anula Sentença",
                               "Decisão Desfavorável TST", "Acórdão Desfavorável TST"}


# ─── HELPERS ───────────────────────────────────────────────────────────────────

def safe_str(v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    return str(v).strip()

def safe_int(v, default=0):
    try:
        return int(v)
    except Exception:
        return default

def fmt_br(n):
    """Format number in Brazilian style."""
    try:
        n = int(n)
        return f"{n:,}".replace(",", ".")
    except Exception:
        return str(n)

def fmt_pct(num, den):
    if den == 0:
        return "0,0%"
    return f"{100*num/den:.1f}%".replace(".", ",")

def json_safe(obj):
    """Recursively convert numpy/pandas types to Python native."""
    if isinstance(obj, dict):
        return {k: json_safe(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [json_safe(i) for i in obj]
    if hasattr(obj, "item"):  # numpy scalar
        return obj.item()
    if isinstance(obj, float) and obj != obj:  # NaN
        return None
    return obj

def filter_nonzero(labels, data):
    """Remove entries where data value is 0 or None."""
    pairs = [(l, d) for l, d in zip(labels, data) if d and int(d) > 0]
    if not pairs:
        return [], []
    ls, ds = zip(*pairs)
    return list(ls), list(ds)

def get_trt(uf, vara=""):
    uf = safe_str(uf).upper()
    vara_lower = safe_str(vara).lower()
    if uf == "SP":
        for kw in SP_TRT15_KEYWORDS:
            if kw in vara_lower:
                return "TRT-15"
        return "TRT-2"
    return UF_TRT.get(uf, f"TRT-{uf}")


# ─── DATA LOADING ──────────────────────────────────────────────────────────────

def load_excel(path):
    print(f"[INFO] Lendo arquivo: {path}")
    xl = pd.ExcelFile(path, engine="openpyxl")
    sheets = xl.sheet_names
    print(f"[INFO] Abas encontradas: {sheets}")

    dfs = {}
    for sheet in sheets:
        try:
            df = xl.parse(sheet, header=0)
            dfs[sheet] = df
            print(f"[INFO] Aba '{sheet}': {df.shape[0]} linhas, {df.shape[1]} colunas")
        except Exception as e:
            print(f"[WARN] Erro ao ler aba '{sheet}': {e}")
            dfs[sheet] = pd.DataFrame()
    return dfs


def parse_ifood_sheet(df):
    """Parse the main IFOOD sheet by column index.
    Actual Excel structure (sheet 'todos', 68 cols):
      0=causa_raiz, 1=num_processo, 2=TRT (ex: 'TRT 2'), 3=status,
      4=cargo, 5=assunto, 6=vara, 7=cidade, 8=UF,
      15=adv1(EscrCNA), 16=juiz_sentenca, 22=fase, 35=arquivado,
      44=autor_recurso, 52=resultado_final, 59=julgador_2inst,
      64=resultado_recente, 67=primeiro_resultado
    """
    if df.empty:
        return []
    rows = []
    for _, row in df.iterrows():
        def c(i):
            try:
                v = row.iloc[i]
                if pd.isna(v):
                    return ""
                return str(v).strip()
            except Exception:
                return ""

        causa_raiz        = c(0)
        num_processo      = c(1)
        trt_raw           = c(2)   # "TRT 2", "TRT 15", etc.
        status_raw        = c(3)
        cargo             = c(4)
        assunto           = c(5)
        vara              = c(6)
        cidade            = c(7)
        uf                = c(8)
        # col 14: Requerente 1, col 15: Escritorio CNA Requerente (law firm name)
        adv1              = c(15)  # Escritorio CNA Requerente
        juiz_sentenca     = c(16)  # Juiz sentença
        fase              = c(22)  # Fase Processual
        arquivado         = c(35)  # Se está arquivado
        autor_recurso     = c(43)  # Autor do Recurso
        resultado_final   = c(51)  # Resultado Final
        julgador_2inst    = c(58)  # Julgador 2a Instancia
        resultado_recente = c(63)  # Resultado Julgamento mais recente
        primeiro_resultado = c(66) # 1o resultado do processo

        # Normalize TRT: "TRT 2" → "TRT-2", "TRT 15" → "TRT-15"
        trt = re.sub(r'TRT\s+', 'TRT-', trt_raw, flags=re.IGNORECASE).strip()
        if not trt:
            trt = get_trt(uf, vara)

        # Normalize status
        s = status_raw.lower()
        if s in ("ativo", "active", "1"):
            status = "ativo"
        elif s in ("encerrado", "encerrada", "closed", "0"):
            status = "encerrado"
        else:
            status = status_raw

        rows.append({
            "causa_raiz": causa_raiz,
            "num_processo": num_processo,
            "status": status,
            "cargo": cargo,
            "assunto": assunto,
            "vara": vara,
            "cidade": cidade,
            "uf": uf,
            "adv1": adv1,
            "adv2": "",
            "adv3": "",
            "juiz_sentenca": juiz_sentenca,
            "fase": fase,
            "arquivado": arquivado,
            "autor_recurso": autor_recurso,
            "resultado_final": resultado_final,
            "julgador": julgador_2inst,
            "resultado_recente": resultado_recente,
            "primeiro_resultado": primeiro_resultado,
            "trt": trt,
        })
    return rows


def parse_decisoes_sheet(df):
    """Parse Decisões sheet.
    Columns: 0=N°Processo, 1=TRT, 2=CAUSA RAIZ, 3=RESULTADO DA DECISÃO,
             4=JUIZ/RELATOR, 5=VARATURMA
    """
    if df.empty:
        return []
    rows = []
    for _, row in df.iterrows():
        def c(i):
            try:
                v = row.iloc[i]
                if pd.isna(v):
                    return ""
                return str(v).strip()
            except Exception:
                return ""
        causa_raiz = c(2).strip()
        resultado  = c(3).strip()
        if causa_raiz:
            rows.append({
                "num_processo":  c(0),
                "trt":           c(1),
                "causa_raiz":    causa_raiz,
                "resultado":     resultado,
                "juiz_relator":  c(4),
                "varaturma":     c(5),
            })
    return rows


def parse_novos_casos(df):
    """Parse Novos Casos Recebidos sheet.
    Sheet has a summary table: row with 'Causa Raiz' header, then CR + count rows.
    Returns dict: { 'total': int, 'por_causa_raiz': { CR: count } }
    """
    if df.empty:
        return {"total": 0, "por_causa_raiz": {}}
    total = 0
    por_causa_raiz = {}
    in_table = False
    for _, row in df.iterrows():
        try:
            v0 = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ""
            # Detect total line
            if v0.lower().startswith("total:") and "novos" in v0.lower():
                try:
                    total = int(re.search(r'\d+', v0).group())
                except Exception:
                    pass
            # Detect table header
            if v0.lower() == "causa raiz":
                in_table = True
                continue
            # Parse table rows
            if in_table and v0 and v0.lower() != "total":
                try:
                    count_raw = row.iloc[1]
                    count = int(float(count_raw)) if not pd.isna(count_raw) else 0
                    if count > 0:
                        por_causa_raiz[v0.upper().strip()] = count
                except Exception:
                    pass
        except Exception:
            pass
    if not total:
        total = sum(por_causa_raiz.values())
    return {"total": total, "por_causa_raiz": por_causa_raiz}


def parse_suspensos(df):
    if df.empty:
        return []
    rows = []
    cols = list(df.columns)
    for _, row in df.iterrows():
        r = {}
        for i, col in enumerate(cols):
            try:
                v = row.iloc[i]
                r[str(col)] = "" if pd.isna(v) else str(v).strip()
            except Exception:
                r[str(col)] = ""
        rows.append(r)
    return rows


# ─── ANALYTICS ─────────────────────────────────────────────────────────────────

def parse_pedidos(assunto_str, top_n=20):
    """Extract and rank individual claims (pedidos) from the Assunto field."""
    import re
    if not assunto_str or str(assunto_str).strip() in ("", "nan", "None"):
        return Counter()
    s = str(assunto_str)
    # Split on » separator (various encodings) or bullet-like chars
    for sep in ['\xbb', '\u00bb', '\u00b7', '»', ' – ', ' - ']:
        s = s.replace(sep, '|')
    parts = [p.strip().title() for p in s.split('|') if p.strip()]
    # If no separator found, try double-space or known keyword patterns
    if len(parts) <= 1:
        parts = [p.strip().title() for p in re.split(r'\s{2,}', s) if p.strip()]
    counter = Counter()
    skip = {'De', 'Do', 'Da', 'Dos', 'Das', 'E', 'Em', 'A', 'O', 'Nan', 'None', ''}
    for p in parts:
        p = re.sub(r'\s+', ' ', p).strip()
        if p and len(p) > 3 and p not in skip:
            counter[p] += 1
    return counter

def compute_pedidos_ranking(records, causa_raiz, top_n=15):
    """Return top N pedidos (claims) for a given causa raiz."""
    total_counter = Counter()
    for rec in records:
        if rec.get("causa_raiz", "") != causa_raiz:
            continue
        total_counter.update(parse_pedidos(rec.get("assunto", "")))
    # Deduplicate near-identical entries (e.g. 'Horas Extras' vs 'Horas Extras Horas Extras ...')
    cleaned = Counter()
    for pedido, count in total_counter.items():
        # Truncate very long strings
        p = pedido[:70] if len(pedido) > 70 else pedido
        cleaned[p] += count
    return cleaned.most_common(top_n)

def is_arquivado(rec):
    arq = rec.get("arquivado", "")
    fase = rec.get("fase", "")
    return (arq not in ("", "0", "False", "false", "NaN", "nan")
            or "arquivamento" in fase.lower()
            or fase == "6-Arquivamento")

def classify_result(res):
    if res in FAVORAVEL_VALUES:
        return "Favorável"
    if res in DESFAVORAVEL_VALUES:
        return "Desfavorável"
    if res in NEUTRO_VALUES:
        return "Neutro"
    return None

def classify_decisao(res):
    if res in DECISAO_FAVORAVEL_VALUES:
        return "Favorável"
    if res in DECISAO_DESFAVORAVEL_VALUES:
        return "Desfavorável"
    return "Outro"

FASE_EXCLUIR = {"arquivamento", "arquivado"}

def _fase_valida(fase_str):
    """Return True if this fase should be included (exclude arquivamento)."""
    return "arquiv" not in fase_str.lower()

def compute_kpis(records, suspensos_1389, suspensos_1291, novos_casos, decisoes):
    total = len(records)
    ativos = sum(1 for r in records if r["status"].lower() == "ativo")
    encerrados = sum(1 for r in records if r["status"].lower() == "encerrado")
    arquivados = sum(1 for r in records if is_arquivado(r))
    suspensos = len(suspensos_1389) + len(suspensos_1291)
    novos = novos_casos.get("total", 0) if isinstance(novos_casos, dict) else len(novos_casos)

    dec_fav = sum(1 for d in decisoes if classify_decisao(d["resultado"]) == "Favorável")
    dec_desf = sum(1 for d in decisoes if classify_decisao(d["resultado"]) == "Desfavorável")
    dec_total = len(decisoes)

    return {
        "total": total,
        "ativos": ativos,
        "encerrados": encerrados,
        "arquivados": arquivados,
        "suspensos": suspensos,
        "novos": novos,
        "dec_fav": dec_fav,
        "dec_desf": dec_desf,
        "dec_total": dec_total,
    }

def compute_causa_raiz_kpis(records, decisoes, causa_raiz, suspensos_1389, suspensos_1291):
    recs = [r for r in records if r["causa_raiz"] == causa_raiz]
    ativos = sum(1 for r in recs if r["status"].lower() == "ativo")
    encerrados = sum(1 for r in recs if r["status"].lower() == "encerrado")
    arquivados = sum(1 for r in recs if is_arquivado(r))

    def match_susp(s):
        return safe_str(s.get("CAUSA RAIZ", "")).upper() == causa_raiz.upper()

    susp = sum(1 for s in suspensos_1389 if match_susp(s)) + \
           sum(1 for s in suspensos_1291 if match_susp(s))

    decs = [d for d in decisoes if d["causa_raiz"] == causa_raiz]
    dec_fav  = sum(1 for d in decs if classify_decisao(d["resultado"]) == "Favorável")
    dec_desf = sum(1 for d in decs if classify_decisao(d["resultado"]) == "Desfavorável")

    # Fase — exclude arquivamento
    fase_counter = Counter(r["fase"] for r in recs if r["fase"] and _fase_valida(r["fase"]))

    # UF distribution (top 10)
    uf_counter = Counter(r["uf"] for r in recs if r["uf"])
    top_ufs = uf_counter.most_common(10)

    # Escritórios CNA adversários (col 15)
    adv_counter = Counter()
    for r in recs:
        for a in [r["adv1"]]:
            if a and a.lower() not in ("nan", "none", "", "0"):
                adv_counter[a] += 1
    top_advs = adv_counter.most_common(10)

    # Julgadores — da aba todos, sem relação com aba Decisões.
    # Para cada processo classifica pelo resultado_recente (col 63).
    # Nome do juiz: usa julgador_2inst (col 58) se preenchido, senão juiz_sentenca (col 16).
    julg_counter      = Counter()
    julg_fav_counter  = Counter()
    julg_desf_counter = Counter()

    for r in recs:
        j = r.get("julgador", "") or r.get("juiz_sentenca", "")
        if not j or j.lower() in ("nan", "none", "", "0"):
            continue
        julg_counter[j] += 1
        cls = classify_result(r["resultado_recente"])
        if cls == "Favorável":
            julg_fav_counter[j] += 1
        elif cls == "Desfavorável":
            julg_desf_counter[j] += 1

    top_julgs      = julg_counter.most_common(10)
    top_julgs_fav  = julg_fav_counter.most_common(10)
    top_julgs_desf = julg_desf_counter.most_common(10)

    # Turma/VT e Julgador/Relator — da aba Decisões
    vt_fav_c   = Counter()
    vt_desf_c  = Counter()
    jr_fav_c   = Counter()
    jr_desf_c  = Counter()
    for d in decs:
        cls = classify_decisao(d["resultado"])
        vt = d.get("varaturma", "")
        jr = d.get("juiz_relator", "")
        if vt:
            if cls == "Favorável":    vt_fav_c[vt]  += 1
            elif cls == "Desfavorável": vt_desf_c[vt] += 1
        if jr:
            if cls == "Favorável":    jr_fav_c[jr]  += 1
            elif cls == "Desfavorável": jr_desf_c[jr] += 1

    all_vts = sorted(set(list(vt_fav_c) + list(vt_desf_c)))
    all_jrs = sorted(set(list(jr_fav_c) + list(jr_desf_c)),
                     key=lambda j: -(jr_fav_c[j] + jr_desf_c[j]))[:15]

    # Suspensos list for this causa raiz
    susp_list_1389 = [s for s in suspensos_1389 if match_susp(s)]
    susp_list_1291 = [s for s in suspensos_1291 if match_susp(s)]

    # Pedidos ranking — todas as causas raiz
    pedidos_ranking = compute_pedidos_ranking(records, causa_raiz, top_n=15)

    return {
        "total": len(recs),
        "ativos": ativos,
        "encerrados": encerrados,
        "arquivados": arquivados,
        "suspensos": susp,
        "dec_fav": dec_fav,
        "dec_desf": dec_desf,
        "dec_total": len(decs),
        "fase_labels": [k for k, _ in fase_counter.most_common()],
        "fase_data":   [v for _, v in fase_counter.most_common()],
        "uf_labels": [k for k, _ in top_ufs],
        "uf_data":   [v for _, v in top_ufs],
        "adv_labels": [k for k, _ in top_advs],
        "adv_data":   [v for _, v in top_advs],
        "julg_labels":      [k for k, _ in top_julgs],
        "julg_data":        [v for _, v in top_julgs],
        "julg_fav_labels":  [k for k, _ in top_julgs_fav],
        "julg_fav_data":    [v for _, v in top_julgs_fav],
        "julg_desf_labels": [k for k, _ in top_julgs_desf],
        "julg_desf_data":   [v for _, v in top_julgs_desf],
        "susp_list_1389": susp_list_1389,
        "susp_list_1291": susp_list_1291,
        "decisoes": decs,
        "pedidos_ranking": pedidos_ranking,
        # Turma/VT e Julgador — da aba Decisões
        "vt_fav_c":   dict(vt_fav_c),
        "vt_desf_c":  dict(vt_desf_c),
        "all_vts":    all_vts,
        "jr_fav_c":   dict(jr_fav_c),
        "jr_desf_c":  dict(jr_desf_c),
        "all_jrs":    all_jrs,
    }

def compute_global_charts(records, decisoes):
    # Causa raiz distribution (ativos)
    cr_counter = Counter(r["causa_raiz"] for r in records if r["status"].lower() == "ativo" and r["causa_raiz"])
    cr_labels = list(cr_counter.keys())
    cr_data   = list(cr_counter.values())

    # Fase processual — exclude arquivamento
    fase_counter = Counter(r["fase"] for r in records if r["fase"] and r["status"].lower() == "ativo" and _fase_valida(r["fase"]))
    fase_labels = [k for k, _ in fase_counter.most_common()]
    fase_data   = [v for _, v in fase_counter.most_common()]

    # Top 15 julgadores
    julg_counter = Counter(r["julgador"] for r in records if r["julgador"] and r["julgador"].lower() not in ("nan", "none", "", "0"))
    top15_julg = julg_counter.most_common(15)
    julg_labels = [k for k, _ in top15_julg]
    julg_data   = [v for _, v in top15_julg]

    # Top 10 escritórios CNA adversários
    adv_counter = Counter()
    for r in records:
        if r["adv1"] and r["adv1"].lower() not in ("nan", "none", "", "0"):
            adv_counter[r["adv1"]] += 1
    top10_adv = adv_counter.most_common(10)
    adv_labels = [k for k, _ in top10_adv]
    adv_data   = [v for _, v in top10_adv]

    # Resultado processos com decisão (col 148)
    res_counter = Counter()
    for r in records:
        c = classify_result(r["resultado_recente"])
        if c:
            res_counter[c] += 1
    res_labels = list(res_counter.keys())
    res_data   = list(res_counter.values())

    # Decisões do período por causa raiz (fav vs desf)
    dec_by_cr_fav  = defaultdict(int)
    dec_by_cr_desf = defaultdict(int)
    for d in decisoes:
        cr = d["causa_raiz"] or "Outros"
        cl = classify_decisao(d["resultado"])
        if cl == "Favorável":
            dec_by_cr_fav[cr] += 1
        elif cl == "Desfavorável":
            dec_by_cr_desf[cr] += 1

    all_crs = sorted(set(list(dec_by_cr_fav.keys()) + list(dec_by_cr_desf.keys())))
    dec_cr_labels = all_crs
    dec_cr_fav   = [dec_by_cr_fav.get(c, 0) for c in all_crs]
    dec_cr_desf  = [dec_by_cr_desf.get(c, 0) for c in all_crs]

    # UF → count and top causa raiz per UF
    uf_counter = Counter(r["uf"] for r in records if r["uf"] and r["status"].lower() == "ativo")

    uf_cr_map = defaultdict(Counter)
    for r in records:
        if r["uf"] and r["status"].lower() == "ativo" and r["causa_raiz"]:
            uf_cr_map[r["uf"]][r["causa_raiz"]] += 1
    uf_causa_raiz_top = {
        uf: counter.most_common(1)[0][0]
        for uf, counter in uf_cr_map.items() if counter
    }

    # TRT ranking with chart data
    trt_recs = defaultdict(list)
    for r in records:
        if r["status"].lower() == "ativo":
            trt_recs[r["trt"]].append(r)

    trt_ranking = []
    for trt, recs in sorted(trt_recs.items(), key=lambda x: -len(x[1])):
        total = len(recs)
        fav  = sum(1 for r in recs if r["resultado_recente"] in FAVORAVEL_VALUES)
        desf = sum(1 for r in recs if r["resultado_recente"] in DESFAVORAVEL_VALUES)
        pct = round(100 * fav / total, 1) if total > 0 else 0.0
        trt_ranking.append({
            "trt": trt,
            "total": total,
            "fav": fav,
            "desf": desf,
            "pct_fav": f"{pct}%" if total > 0 else "N/A",
            "pct_fav_num": pct,
        })

    # TRT × Causa Raiz stacked (ativos)
    trt_cr_map = defaultdict(lambda: defaultdict(int))
    for r in records:
        if r["status"].lower() == "ativo" and r["trt"] and r["causa_raiz"]:
            trt_cr_map[r["trt"]][r["causa_raiz"]] += 1
    # top 12 TRTs by total
    trt_cr_sorted = sorted(trt_cr_map.items(), key=lambda x: -sum(x[1].values()))[:12]

    return {
        "cr_labels": cr_labels,
        "cr_data": cr_data,
        "trt_cr_sorted": trt_cr_sorted,
        "fase_labels": fase_labels,
        "fase_data": fase_data,
        "julg_labels": julg_labels,
        "julg_data": julg_data,
        "adv_labels": adv_labels,
        "adv_data": adv_data,
        "res_labels": res_labels,
        "res_data": res_data,
        "dec_cr_labels": dec_cr_labels,
        "dec_cr_fav": dec_cr_fav,
        "dec_cr_desf": dec_cr_desf,
        "uf_counter": dict(uf_counter),
        "uf_causa_raiz_top": uf_causa_raiz_top,
        "trt_ranking": trt_ranking,
    }


# ─── HTML GENERATION ───────────────────────────────────────────────────────────

BRAZIL_SVG = """
<svg id="brazil-map" viewBox="0 0 800 860" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-height:520px;">
  <!-- Simplified Brazil state polygons (approximate) -->
  <g id="states">
    <path id="AC" d="M80,420 L120,400 L160,410 L170,450 L130,460 L80,450 Z" stroke="#fff" stroke-width="1"/>
    <path id="AL" d="M620,340 L650,330 L660,360 L630,370 L610,355 Z" stroke="#fff" stroke-width="1"/>
    <path id="AM" d="M100,280 L200,240 L300,250 L320,300 L280,350 L200,360 L130,340 L90,310 Z" stroke="#fff" stroke-width="1"/>
    <path id="AP" d="M440,170 L480,155 L500,185 L470,210 L440,200 Z" stroke="#fff" stroke-width="1"/>
    <path id="BA" d="M530,330 L610,310 L640,370 L630,440 L580,480 L530,470 L500,430 L490,370 Z" stroke="#fff" stroke-width="1"/>
    <path id="CE" d="M590,260 L640,250 L660,280 L640,310 L610,310 L590,290 Z" stroke="#fff" stroke-width="1"/>
    <path id="DF" d="M515,415 L530,415 L530,430 L515,430 Z" stroke="#fff" stroke-width="1.5"/>
    <path id="ES" d="M630,430 L660,420 L670,460 L645,470 L620,455 Z" stroke="#fff" stroke-width="1"/>
    <path id="GO" d="M490,380 L550,370 L560,420 L540,460 L490,455 L470,420 Z" stroke="#fff" stroke-width="1"/>
    <path id="MA" d="M480,250 L540,240 L560,280 L530,310 L490,300 L470,275 Z" stroke="#fff" stroke-width="1"/>
    <path id="MG" d="M530,440 L610,420 L650,460 L640,520 L590,550 L540,540 L510,500 L510,460 Z" stroke="#fff" stroke-width="1"/>
    <path id="MS" d="M420,480 L490,470 L510,510 L490,550 L440,555 L410,520 Z" stroke="#fff" stroke-width="1"/>
    <path id="MT" d="M330,360 L420,340 L450,380 L450,440 L400,460 L340,450 L310,400 Z" stroke="#fff" stroke-width="1"/>
    <path id="PA" d="M330,220 L440,190 L490,220 L500,270 L470,300 L390,310 L330,290 L300,260 Z" stroke="#fff" stroke-width="1"/>
    <path id="PB" d="M640,280 L680,272 L688,298 L656,305 L638,295 Z" stroke="#fff" stroke-width="1"/>
    <path id="PE" d="M570,300 L638,292 L650,318 L615,328 L572,320 Z" stroke="#fff" stroke-width="1"/>
    <path id="PI" d="M535,268 L580,258 L590,298 L558,310 L532,300 Z" stroke="#fff" stroke-width="1"/>
    <path id="PR" d="M480,570 L545,558 L565,595 L540,625 L485,625 L460,600 Z" stroke="#fff" stroke-width="1"/>
    <path id="RJ" d="M615,500 L655,488 L670,515 L648,535 L615,525 Z" stroke="#fff" stroke-width="1"/>
    <path id="RN" d="M652,260 L688,252 L698,275 L670,282 L650,272 Z" stroke="#fff" stroke-width="1"/>
    <path id="RO" d="M200,370 L280,355 L300,400 L270,430 L210,430 L185,400 Z" stroke="#fff" stroke-width="1"/>
    <path id="RR" d="M200,170 L270,155 L300,190 L280,230 L220,240 L185,210 Z" stroke="#fff" stroke-width="1"/>
    <path id="RS" d="M445,650 L510,640 L535,680 L510,720 L455,725 L425,690 Z" stroke="#fff" stroke-width="1"/>
    <path id="SC" d="M470,625 L540,615 L558,648 L530,668 L472,665 L452,645 Z" stroke="#fff" stroke-width="1"/>
    <path id="SE" d="M650,360 L672,352 L678,375 L655,382 L640,370 Z" stroke="#fff" stroke-width="1"/>
    <path id="SP" d="M520,495 L600,478 L625,510 L610,555 L555,565 L520,545 L505,515 Z" stroke="#fff" stroke-width="1"/>
    <path id="TO" d="M450,300 L510,285 L530,330 L515,375 L470,378 L445,340 Z" stroke="#fff" stroke-width="1"/>
  </g>
  <!-- State labels -->
  <g id="state-labels" font-family="IBM Plex Mono, monospace" font-size="9" fill="#fff" text-anchor="middle" font-weight="600" pointer-events="none">
    <text x="128" y="435">AC</text>
    <text x="632" y="353">AL</text>
    <text x="205" y="305">AM</text>
    <text x="468" y="185">AP</text>
    <text x="565" y="400">BA</text>
    <text x="622" y="283">CE</text>
    <text x="522" y="424">DF</text>
    <text x="645" y="447">ES</text>
    <text x="515" y="418">GO</text>
    <text x="510" y="278">MA</text>
    <text x="580" y="490">MG</text>
    <text x="455" y="515">MS</text>
    <text x="385" y="402">MT</text>
    <text x="405" y="255">PA</text>
    <text x="662" y="290">PB</text>
    <text x="608" y="312">PE</text>
    <text x="558" y="288">PI</text>
    <text x="510" y="595">PR</text>
    <text x="638" y="513">RJ</text>
    <text x="672" y="268">RN</text>
    <text x="245" y="400">RO</text>
    <text x="242" y="198">RR</text>
    <text x="478" y="685">RS</text>
    <text x="503" y="645">SC</text>
    <text x="657" y="368">SE</text>
    <text x="563" y="523">SP</text>
    <text x="487" y="335">TO</text>
  </g>
</svg>
"""

def escape_js(s):
    return s.replace("\\", "\\\\").replace("`", "\\`").replace("$", "\\$")

def build_kpi_cards(kpis_list):
    """Build HTML for a row of KPI cards."""
    cards = []
    for label, value, unit, trend_cls, trend_txt in kpis_list:
        trend_html = f'<div class="kpi-trend {trend_cls}">{trend_txt}</div>' if trend_txt else ""
        cards.append(f"""
        <div class="kpi-card">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value">{value}</div>
          <div class="kpi-unit">{unit}</div>
          {trend_html}
        </div>""")
    return "\n".join(cards)

def build_table(headers, rows, max_rows=200):
    """Build an HTML table."""
    th = "".join(f"<th>{h}</th>" for h in headers)
    trs = []
    for row in rows[:max_rows]:
        tds = "".join(f"<td>{cell}</td>" for cell in row)
        trs.append(f"<tr>{tds}</tr>")
    tbody = "\n".join(trs)
    return f"""
    <div class="table-wrap">
      <table class="data-table">
        <thead><tr>{th}</tr></thead>
        <tbody>{tbody}</tbody>
      </table>
    </div>"""

def build_causa_raiz_tab(tab_id, label, data):
    """Build a complete causa raiz tab panel."""
    # Change 8: Filter zero values in KPI cards
    kpi_cards = []
    if data["total"] > 0:
        kpi_cards.insert(0, ("Total de Processos", fmt_br(data["total"]), "processos", "", ""))
    if data["ativos"] > 0:
        kpi_cards.append(("Processos Ativos", fmt_br(data["ativos"]), "processos", "", ""))
    if data["encerrados"] > 0:
        kpi_cards.append(("Processos Encerrados", fmt_br(data["encerrados"]), "processos", "", ""))
    if data["suspensos"] > 0:
        kpi_cards.append(("Suspensos", fmt_br(data["suspensos"]), "processos", "", ""))
    kpis_row1 = build_kpi_cards(kpi_cards) if kpi_cards else ""

    dec_cards = []
    if data["dec_fav"] > 0 or data["dec_desf"] > 0:
        dec_pct_fav = fmt_pct(data["dec_fav"], data["dec_total"]) if data["dec_total"] else "N/A"
        dec_cards.append(("Decisões Favoráveis", fmt_br(data["dec_fav"]), f"({dec_pct_fav})", "up", "Favorável"))
        dec_cards.append(("Decisões Desfavoráveis", fmt_br(data["dec_desf"]), "decisões do período", "down", ""))
        dec_cards.append(("Total Decisões", fmt_br(data["dec_total"]), "no período", "flat", ""))
    kpis_row2 = build_kpi_cards(dec_cards) if dec_cards else ""

    # Turma/VT e Julgador/Relator — da aba Decisões
    dec_vt_html = ""
    if data.get("all_vts") or data.get("all_jrs"):
        dec_vt_html = '<h3 class="section-title" style="margin:28px 0 16px">Decisões do Período — Turma/VT e Julgador/Relator</h3>'
        dec_vt_html += f"""
    <div class="charts-grid">
      <div class="chart-card">
        <div class="chart-card-title">Decisões por Turma/VT</div>
        <div class="chart-wrap" style="height:280px"><canvas id="{dec_vt_id}"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-card-title">Decisões por Julgador/Relator</div>
        <div class="chart-wrap" style="height:280px"><canvas id="{dec_jr_id}"></canvas></div>
      </div>
    </div>"""

    # Suspensos tables
    susp_html = ""
    if data["susp_list_1389"]:
        cols_1389 = list(data["susp_list_1389"][0].keys()) if data["susp_list_1389"] else []
        rows_1389 = [[s.get(c, "") for c in cols_1389] for s in data["susp_list_1389"]]
        susp_html += f'<h3 class="section-title" style="margin:24px 0 12px">Processos Suspensos — Tema 1389</h3>'
        susp_html += build_table(cols_1389, rows_1389)

    if data["susp_list_1291"]:
        cols_1291 = list(data["susp_list_1291"][0].keys()) if data["susp_list_1291"] else []
        rows_1291 = [[s.get(c, "") for c in cols_1291] for s in data["susp_list_1291"]]
        susp_html += f'<h3 class="section-title" style="margin:24px 0 12px">Processos Suspensos — Tema 1291</h3>'
        susp_html += build_table(cols_1291, rows_1291)

    fase_id      = f"chart_{tab_id}_fase"
    uf_id        = f"chart_{tab_id}_uf"
    dec_id       = f"chart_{tab_id}_dec"
    pedidos_id   = f"chart_{tab_id}_pedidos"
    adv_id       = f"chart_{tab_id}_adv"
    julg_id      = f"chart_{tab_id}_julg"
    julg_fav_id  = f"chart_{tab_id}_julgfav"
    julg_desf_id = f"chart_{tab_id}_julgdesf"
    dec_vt_id    = f"chart_{tab_id}_decvt"
    dec_jr_id    = f"chart_{tab_id}_decjr"

    # Pedidos ranking — somente aba Ex-Foodlover
    pedidos_html = ""
    if tab_id == "ex-foodlover" and data.get("pedidos_ranking"):
        ped_rows = [[rank+1, p, fmt_br(n)] for rank, (p, n) in enumerate(data["pedidos_ranking"])]
        ped_table = build_table(["#", "Pedido / Assunto", "Ocorrências"], ped_rows)
        pedidos_html = f"""
    <h3 class="section-title" style="margin:28px 0 16px">Ranking de Pedidos dos Reclamantes (Top 15)</h3>
    <div class="charts-grid">
      <div class="chart-card" style="grid-column:span 2">
        <div class="chart-card-title">Pedidos Mais Frequentes — {label}</div>
        <div class="chart-wrap" style="height:340px"><canvas id="{pedidos_id}"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-card-title">Tabela — Ranking de Pedidos</div>
        {ped_table}
      </div>
    </div>"""

    return f"""
  <div id="tab-{tab_id}" class="tab-panel">
    <div class="page-header">
      <div class="page-label">Causa Raiz</div>
      <h1 class="page-title">{label}</h1>
      <p class="page-subtitle">Visão consolidada dos processos desta categoria</p>
    </div>

    <div class="kpi-grid">{kpis_row1}</div>
    <div class="kpi-grid">{kpis_row2}</div>

    <div class="charts-grid">
      <div class="chart-card">
        <div class="chart-card-title">Decisões — Favorável × Desfavorável</div>
        <div class="chart-wrap"><canvas id="{dec_id}"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-card-title">Fase Processual</div>
        <div class="chart-wrap"><canvas id="{fase_id}"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-card-title">Distribuição por UF (Top 10)</div>
        <div class="chart-wrap"><canvas id="{uf_id}"></canvas></div>
      </div>
    </div>

    <div class="charts-grid">
      <div class="chart-card">
        <div class="chart-card-title">Top 10 Advogados Adversários</div>
        <div class="chart-wrap" style="height:260px"><canvas id="{adv_id}"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-card-title">Top 10 Julgadores (geral)</div>
        <div class="chart-wrap" style="height:260px"><canvas id="{julg_id}"></canvas></div>
      </div>
    </div>

    <h3 class="section-title" style="margin:8px 0 16px">Julgadores por Resultado</h3>
    <div class="charts-grid">
      <div class="chart-card">
        <div class="chart-card-title" style="color:var(--green)">✓ Julgadores — Decisões Favoráveis</div>
        <div class="chart-wrap" style="height:260px"><canvas id="{julg_fav_id}"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-card-title" style="color:var(--red)">✕ Julgadores — Decisões Desfavoráveis</div>
        <div class="chart-wrap" style="height:260px"><canvas id="{julg_desf_id}"></canvas></div>
      </div>
    </div>

    {pedidos_html}
    {dec_vt_html}
    {susp_html}
  </div>"""

def build_causa_raiz_js(tab_id, data):
    """Build Chart.js initialization code for a causa raiz tab."""
    # Change 3: apply filter_nonzero to all chart data
    fase_labels_f, fase_data_f = filter_nonzero(data["fase_labels"], data["fase_data"])
    uf_labels_f, uf_data_f     = filter_nonzero(data["uf_labels"],   data["uf_data"])
    adv_labels_f, adv_data_f   = filter_nonzero(data["adv_labels"],  data["adv_data"])
    julg_labels_f, julg_data_f = filter_nonzero(data["julg_labels"], data["julg_data"])

    fase_labels = json.dumps(fase_labels_f, ensure_ascii=False)
    fase_data   = json.dumps(fase_data_f)
    uf_labels   = json.dumps(uf_labels_f, ensure_ascii=False)
    uf_data     = json.dumps(uf_data_f)
    # Decisões vêm exclusivamente da aba Decisões (não do campo resultado_recente)
    dec_outros  = max(0, data["dec_total"] - data["dec_fav"] - data["dec_desf"])
    dec_labels  = json.dumps(["Favorável", "Desfavorável", "Outro"], ensure_ascii=False)
    dec_data    = json.dumps([data["dec_fav"], data["dec_desf"], dec_outros])

    pedidos_js = ""
    if tab_id == "ex-foodlover" and data.get("pedidos_ranking"):
        ped_labels = json.dumps([p for p, _ in data["pedidos_ranking"]], ensure_ascii=False)
        ped_data   = json.dumps([n for _, n in data["pedidos_ranking"]])
        pedidos_js = f"""
    new Chart(document.getElementById('chart_{tab_id}_pedidos'), {{
      type: 'bar',
      data: {{ labels: {ped_labels}, datasets: [{{ label:'Ocorrências', data: {ped_data},
        backgroundColor: ['#7B1A2E','#0D2B5E','#0F9E76','#E67E22','#8E44AD',
                          '#C0392B','#2563EB','#059669','#D97706','#7C3AED',
                          '#B91C1C','#1D4ED8','#065F46','#92400E','#5B21B6'],
        borderRadius: 5 }}] }},
      options: {{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
        plugins:{{ legend:{{ display:false }},
          tooltip:{{ callbacks:{{ label: ctx => ' ' + ctx.parsed.x + ' processos' }} }} }},
        scales:{{
          x:{{ grid:{{ color:'rgba(0,0,0,0.05)' }}, ticks:{{ font:{{ family:'IBM Plex Mono', size:10 }} }} }},
          y:{{ grid:{{ display:false }}, ticks:{{ font:{{ family:'IBM Plex Mono', size:10 }}, color:'#3D2B2E' }} }}
        }}
      }}
    }});"""

    # Julgadores fav/desf
    julg_fav_labels_f, julg_fav_data_f   = filter_nonzero(data.get("julg_fav_labels", []),  data.get("julg_fav_data", []))
    julg_desf_labels_f, julg_desf_data_f = filter_nonzero(data.get("julg_desf_labels", []), data.get("julg_desf_data", []))

    adv_js = ""
    if adv_labels_f:
        adv_labels_s = json.dumps(adv_labels_f, ensure_ascii=False)
        adv_data_s   = json.dumps(adv_data_f)
        adv_js = f"""
    new Chart(document.getElementById('chart_{tab_id}_adv'), {{
      type: 'bar',
      data: {{ labels: {adv_labels_s}, datasets: [{{ label:'Processos', data: {adv_data_s},
        backgroundColor: '#7B1A2E', borderRadius: 4 }}] }},
      options: {{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
        plugins:{{ legend:{{ display:false }} }},
        scales:{{ x:{{ grid:{{ color:'rgba(0,0,0,0.05)' }} }}, y:{{ grid:{{ display:false }}, ticks:{{ font:{{ size:9 }} }} }} }}
      }}
    }});"""

    julg_js = ""
    if julg_labels_f:
        julg_labels_s = json.dumps(julg_labels_f, ensure_ascii=False)
        julg_data_s   = json.dumps(julg_data_f)
        julg_js = f"""
    new Chart(document.getElementById('chart_{tab_id}_julg'), {{
      type: 'bar',
      data: {{ labels: {julg_labels_s}, datasets: [{{ label:'Processos', data: {julg_data_s},
        backgroundColor: '#6366f1', borderRadius: 4 }}] }},
      options: {{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
        plugins:{{ legend:{{ display:false }} }},
        scales:{{ x:{{ grid:{{ color:'rgba(0,0,0,0.05)' }} }}, y:{{ grid:{{ display:false }}, ticks:{{ font:{{ size:9 }} }} }} }}
      }}
    }});"""

    julg_fav_js = ""
    if julg_fav_labels_f:
        s_lbl = json.dumps(julg_fav_labels_f, ensure_ascii=False)
        s_dat = json.dumps(julg_fav_data_f)
        julg_fav_js = f"""
    new Chart(document.getElementById('chart_{tab_id}_julgfav'), {{
      type: 'bar',
      data: {{ labels: {s_lbl}, datasets: [{{ label:'Favoráveis', data: {s_dat},
        backgroundColor: '#0F9E76', borderRadius: 4 }}] }},
      options: {{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
        plugins:{{ legend:{{ display:false }} }},
        scales:{{ x:{{ grid:{{ color:'rgba(0,0,0,0.05)' }} }}, y:{{ grid:{{ display:false }}, ticks:{{ font:{{ size:9 }} }} }} }}
      }}
    }});"""

    julg_desf_js = ""
    if julg_desf_labels_f:
        s_lbl = json.dumps(julg_desf_labels_f, ensure_ascii=False)
        s_dat = json.dumps(julg_desf_data_f)
        julg_desf_js = f"""
    new Chart(document.getElementById('chart_{tab_id}_julgdesf'), {{
      type: 'bar',
      data: {{ labels: {s_lbl}, datasets: [{{ label:'Desfavoráveis', data: {s_dat},
        backgroundColor: '#7B1A2E', borderRadius: 4 }}] }},
      options: {{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
        plugins:{{ legend:{{ display:false }} }},
        scales:{{ x:{{ grid:{{ color:'rgba(0,0,0,0.05)' }} }}, y:{{ grid:{{ display:false }}, ticks:{{ font:{{ size:9 }} }} }} }}
      }}
    }});"""

    # Turma/VT e Julgador/Relator da aba Decisões
    dec_vt_js = ""
    all_vts = data.get("all_vts", [])
    if all_vts:
        vt_fav_c  = data.get("vt_fav_c",  {})
        vt_desf_c = data.get("vt_desf_c", {})
        vt_labels = json.dumps(all_vts, ensure_ascii=False)
        vt_fav    = json.dumps([vt_fav_c.get(v, 0)  for v in all_vts])
        vt_desf   = json.dumps([vt_desf_c.get(v, 0) for v in all_vts])
        dec_vt_js = f"""
    new Chart(document.getElementById('chart_{tab_id}_decvt'), {{
      type: 'bar',
      data: {{ labels: {vt_labels}, datasets: [
        {{ label:'Favorável',    data:{vt_fav},  backgroundColor:'#0F9E76', borderRadius:4 }},
        {{ label:'Desfavorável', data:{vt_desf}, backgroundColor:'#7B1A2E', borderRadius:4 }}
      ]}},
      options: {{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
        plugins:{{ legend:{{ position:'top', labels:{{ font:{{ family:'IBM Plex Mono', size:10 }} }} }} }},
        scales:{{ x:{{ grid:{{ color:'rgba(0,0,0,0.05)' }}, beginAtZero:true }},
                  y:{{ grid:{{ display:false }}, ticks:{{ font:{{ size:9 }} }} }} }} }}
    }});"""

    dec_jr_js = ""
    all_jrs = data.get("all_jrs", [])
    if all_jrs:
        jr_fav_c  = data.get("jr_fav_c",  {})
        jr_desf_c = data.get("jr_desf_c", {})
        jr_labels = json.dumps(all_jrs, ensure_ascii=False)
        jr_fav    = json.dumps([jr_fav_c.get(j, 0)  for j in all_jrs])
        jr_desf   = json.dumps([jr_desf_c.get(j, 0) for j in all_jrs])
        dec_jr_js = f"""
    new Chart(document.getElementById('chart_{tab_id}_decjr'), {{
      type: 'bar',
      data: {{ labels: {jr_labels}, datasets: [
        {{ label:'Favorável',    data:{jr_fav},  backgroundColor:'#0F9E76', borderRadius:4 }},
        {{ label:'Desfavorável', data:{jr_desf}, backgroundColor:'#7B1A2E', borderRadius:4 }}
      ]}},
      options: {{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
        plugins:{{ legend:{{ position:'top', labels:{{ font:{{ family:'IBM Plex Mono', size:10 }} }} }} }},
        scales:{{ x:{{ grid:{{ color:'rgba(0,0,0,0.05)' }}, beginAtZero:true }},
                  y:{{ grid:{{ display:false }}, ticks:{{ font:{{ size:9 }} }} }} }} }}
    }});"""

    return f"""
    // — {tab_id} charts —
    new Chart(document.getElementById('chart_{tab_id}_dec'), {{
      type: 'doughnut',
      data: {{ labels: {dec_labels}, datasets: [{{ data: {dec_data},
        backgroundColor: ['#0F9E76','#7B1A2E','#f59e0b'],
        borderWidth: 0 }}] }},
      options: {{ responsive:true, plugins:{{ legend:{{ position:'right', labels:{{ font:{{ family:'IBM Plex Mono', size:10 }}, color:'#1A0A0D' }} }} }} }}
    }});
    new Chart(document.getElementById('chart_{tab_id}_fase'), {{
      type: 'bar',
      data: {{ labels: {fase_labels}, datasets: [{{ label:'Processos', data: {fase_data},
        backgroundColor: '#0D2B5E', borderRadius: 4 }}] }},
      options: {{ indexAxis:'y', responsive:true, plugins:{{ legend:{{ display:false }} }},
        scales:{{ x:{{ grid:{{ color:'rgba(0,0,0,0.05)' }} }}, y:{{ grid:{{ display:false }} }} }} }}
    }});
    new Chart(document.getElementById('chart_{tab_id}_uf'), {{
      type: 'bar',
      data: {{ labels: {uf_labels}, datasets: [{{ label:'Processos', data: {uf_data},
        backgroundColor: '#7B1A2E', borderRadius: 4 }}] }},
      options: {{ responsive:true, plugins:{{ legend:{{ display:false }} }},
        scales:{{ x:{{ grid:{{ display:false }} }}, y:{{ grid:{{ color:'rgba(0,0,0,0.05)' }} }} }} }}
    }});{adv_js}{julg_js}{julg_fav_js}{julg_desf_js}{pedidos_js}{dec_vt_js}{dec_jr_js}"""


def build_html(kpis, global_charts, causa_raiz_data, suspensos_1389, suspensos_1291,
               decisoes, period_str):

    # ── VISÃO GERAL KPIs ──
    # Change 4: updated KPI rows to match screenshot layout
    kpi_row1 = build_kpi_cards([
        ("Total de Processos",    fmt_br(kpis["total"]),     "processos",  "", ""),
        ("Processos Ativos",      fmt_br(kpis["ativos"]),    "ativos",     "", ""),
        ("Processos Encerrados",  fmt_br(kpis["encerrados"]),"encerrados", "", ""),
        ("Novos Casos no Mês",    fmt_br(kpis["novos"]),     "novos",      "up", "+Período"),
        ("Decisões no Período",   fmt_br(kpis["dec_total"]), "decisões",   "flat", ""),
    ])
    dec_pct = fmt_pct(kpis["dec_fav"], kpis["dec_total"]) if kpis["dec_total"] else "N/A"
    kpi_row2 = build_kpi_cards([
        ("Decisões Favoráveis",    fmt_br(kpis["dec_fav"]),  f"({dec_pct})",   "up",   "Favorável"),
        ("Decisões Desfavoráveis", fmt_br(kpis["dec_desf"]), f"({fmt_pct(kpis['dec_desf'], kpis['dec_total'])})" if kpis["dec_total"] else "", "down", ""),
        ("Suspensos",              fmt_br(kpis["suspensos"]), "processos",     "flat", ""),
    ])

    # ── TRT RANKING TABLE + CHART DATA ──
    trt_rows = [[r["trt"], fmt_br(r["total"]), r["pct_fav"]]
                for r in global_charts["trt_ranking"]]
    trt_table = build_table(["TRT", "Processos", "% Favorável"], trt_rows)

    trt_chart_labels = json.dumps([r["trt"] for r in global_charts["trt_ranking"]], ensure_ascii=False)
    trt_chart_vol    = json.dumps([r["total"]       for r in global_charts["trt_ranking"]])
    trt_chart_fav    = json.dumps([r["pct_fav_num"] for r in global_charts["trt_ranking"]])

    # ── DECISÕES TABLE ──
    dec_rows = [[d["num_processo"], d["causa_raiz"], d["resultado"], d["tipo"]]
                for d in decisoes]
    dec_table = build_table(
        ["Nº Processo", "Causa Raiz", "Resultado", "Tipo de Decisão"],
        dec_rows, max_rows=1000
    )

    # ── SUSPENSOS TABLES ──
    susp_cols_1389 = list(suspensos_1389[0].keys()) if suspensos_1389 else ["Nenhum dado"]
    susp_rows_1389 = [[s.get(c,"") for c in susp_cols_1389] for s in suspensos_1389]

    susp_cols_1291 = list(suspensos_1291[0].keys()) if suspensos_1291 else ["Nenhum dado"]
    susp_rows_1291 = [[s.get(c,"") for c in susp_cols_1291] for s in suspensos_1291]

    susp_table_1389 = build_table(susp_cols_1389, susp_rows_1389, 500)
    susp_table_1291 = build_table(susp_cols_1291, susp_rows_1291, 500)

    # ── CAUSA RAIZ TABS HTML ──
    cr_tabs_html = ""
    cr_tabs_buttons = ""
    for causa, label in CAUSA_RAIZ_TABS:
        tab_id = re.sub(r'[^a-z0-9]', '_', label.lower())
        data = causa_raiz_data.get(causa, {
            "total":0,"ativos":0,"encerrados":0,"arquivados":0,"suspensos":0,
            "dec_fav":0,"dec_desf":0,"dec_total":0,
            "fase_labels":[],"fase_data":[],"uf_labels":[],"uf_data":[],
            "adv_labels":[],"adv_data":[],"julg_labels":[],"julg_data":[],
            "julg_fav_labels":[],"julg_fav_data":[],"julg_desf_labels":[],"julg_desf_data":[],
            "res_fav":0,"res_desf":0,"res_neutro":0,
            "susp_list_1389":[],"susp_list_1291":[],"decisoes":[],"pedidos_ranking":[]
        })
        cr_tabs_html += build_causa_raiz_tab(tab_id, label, data)
        cr_tabs_buttons += f'<button class="tab-btn" data-tab="{tab_id}">{label}</button>\n'

    # ── CAUSA RAIZ JS ──
    cr_js = ""
    for causa, label in CAUSA_RAIZ_TABS:
        tab_id = re.sub(r'[^a-z0-9]', '_', label.lower())
        data = causa_raiz_data.get(causa, {
            "fase_labels":[],"fase_data":[],"uf_labels":[],"uf_data":[],
            "adv_labels":[],"adv_data":[],"julg_labels":[],"julg_data":[],
            "julg_fav_labels":[],"julg_fav_data":[],"julg_desf_labels":[],"julg_desf_data":[],
            "res_fav":0,"res_desf":0,"res_neutro":0,"pedidos_ranking":[],
        })
        cr_js += build_causa_raiz_js(tab_id, data)

    # ── TRT × CAUSA RAIZ JSON ──
    _TRT_CR_PAL = ['#0D2B5E','#7B1A2E','#0F9E76','#f59e0b','#6366f1','#ec4899']
    trt_cr_data = global_charts["trt_cr_sorted"]
    trt_cr_trts = [t for t, _ in trt_cr_data]
    all_causas  = [c for c, _ in CAUSA_RAIZ_TABS]
    all_labels  = [l for _, l in CAUSA_RAIZ_TABS]
    trt_cr_datasets = []
    for i, (causa, label) in enumerate(CAUSA_RAIZ_TABS):
        dataset_data = [trt_cr_data[j][1].get(causa, 0) for j in range(len(trt_cr_data))]
        if any(v > 0 for v in dataset_data):
            trt_cr_datasets.append({
                "label": label,
                "data": dataset_data,
                "backgroundColor": _TRT_CR_PAL[i % len(_TRT_CR_PAL)],
                "borderRadius": 3,
                "stack": "s",
            })
    trt_cr_labels_js   = json.dumps(trt_cr_trts, ensure_ascii=False)
    trt_cr_datasets_js = json.dumps(json_safe(trt_cr_datasets), ensure_ascii=False)

    # ── UF → count + causa raiz JSON for SVG map ──
    uf_json    = json.dumps(global_charts["uf_counter"], ensure_ascii=False)
    uf_cr_json = json.dumps(global_charts["uf_causa_raiz_top"], ensure_ascii=False)
    max_uf     = max(global_charts["uf_counter"].values()) if global_charts["uf_counter"] else 1

    # ── GLOBAL CHART DATA — Change 7: apply filter_nonzero ──
    cr_labels_f, cr_data_f = filter_nonzero(global_charts["cr_labels"], global_charts["cr_data"])
    fase_labels_f, fase_data_f = filter_nonzero(global_charts["fase_labels"], global_charts["fase_data"])
    julg_labels_f, julg_data_f = filter_nonzero(global_charts["julg_labels"], global_charts["julg_data"])
    adv_labels_f, adv_data_f = filter_nonzero(global_charts["adv_labels"], global_charts["adv_data"])
    res_labels_f, res_data_f = filter_nonzero(global_charts["res_labels"], global_charts["res_data"])

    cr_labels_js    = json.dumps(cr_labels_f, ensure_ascii=False)
    cr_data_js      = json.dumps(cr_data_f)
    fase_labels_js  = json.dumps(fase_labels_f, ensure_ascii=False)
    fase_data_js    = json.dumps(fase_data_f)
    julg_labels_js  = json.dumps(julg_labels_f, ensure_ascii=False)
    julg_data_js    = json.dumps(julg_data_f)
    adv_labels_js   = json.dumps(adv_labels_f, ensure_ascii=False)
    adv_data_js     = json.dumps(adv_data_f)
    res_labels_js   = json.dumps(res_labels_f, ensure_ascii=False)
    res_data_js     = json.dumps(res_data_f)

    # Change 7: filter dec_cr grouped bar per-dataset
    dec_cr_pairs = [(l, f, d) for l, f, d in zip(global_charts["dec_cr_labels"], global_charts["dec_cr_fav"], global_charts["dec_cr_desf"]) if f > 0 or d > 0]
    if dec_cr_pairs:
        dec_cr_labels_f = [x[0] for x in dec_cr_pairs]
        dec_cr_fav_f    = [x[1] for x in dec_cr_pairs]
        dec_cr_desf_f   = [x[2] for x in dec_cr_pairs]
    else:
        dec_cr_labels_f, dec_cr_fav_f, dec_cr_desf_f = [], [], []

    dec_cr_labels_js = json.dumps(dec_cr_labels_f, ensure_ascii=False)
    dec_cr_fav_js   = json.dumps(dec_cr_fav_f)
    dec_cr_desf_js  = json.dumps(dec_cr_desf_f)

    # ── DECISÕES SUMMARY ──
    dec_fav_count  = sum(1 for d in decisoes if classify_decisao(d["resultado"]) == "Favorável")
    dec_desf_count = sum(1 for d in decisoes if classify_decisao(d["resultado"]) == "Desfavorável")
    dec_tipo_counter = Counter(d["tipo"] for d in decisoes if d["tipo"])
    dec_tipo_summary = "".join(
        f'<div class="kpi-card"><div class="kpi-label">{t}</div><div class="kpi-value">{fmt_br(c)}</div><div class="kpi-unit">decisões</div></div>'
        for t, c in dec_tipo_counter.most_common()
    )

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>Dashboard iFood — Falaw Advogados — {period_str}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com"/>
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin/>
  <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet"/>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    :root {{
      --bg:       #FAF8F5;
      --bg-card:  #FFFFFF;
      --border:   rgba(0,0,0,0.08);
      --navy:     #0D2B5E;
      --crimson:  #7B1A2E;
      --crimson-l:#9B2238;
      --text:     #1A0A0D;
      --muted:    rgba(26,10,13,0.52);
      --green:    #0F9E76;
      --red:      #7B1A2E;
      --sidebar-w: 248px;
    }}
    html, body {{ height:100%; background:var(--bg); color:var(--text); font-family:'DM Sans',sans-serif; font-size:14px; }}
    a {{ text-decoration:none; color:inherit; }}

    /* LAYOUT */
    .layout {{ display:flex; min-height:100vh; }}

    /* SIDEBAR */
    .sidebar {{
      width:var(--sidebar-w); background:linear-gradient(175deg,#0D2B5E 0%,#3D0F1D 100%);
      display:flex; flex-direction:column; position:fixed; top:0; left:0; bottom:0; z-index:100;
      transition:transform .25s;
    }}
    .sidebar-top {{ padding:28px 20px 20px; border-bottom:1px solid rgba(255,255,255,0.1); }}
    .sidebar-logo {{ font-family:'IBM Plex Mono',monospace; font-size:13px; font-weight:600;
      letter-spacing:.08em; color:#fff; margin-bottom:18px; }}
    .sidebar-logo span {{ color:var(--crimson-l); }}
    .sidebar-badge {{
      font-family:'IBM Plex Mono',monospace; font-size:8px; letter-spacing:.18em;
      text-transform:uppercase; color:var(--crimson-l);
      background:rgba(155,34,56,0.2); border:1px solid rgba(155,34,56,0.35);
      padding:3px 8px; border-radius:2px; display:inline-block; margin-bottom:10px;
    }}
    .sidebar-client {{ font-size:14px; font-weight:700; color:#fff; }}
    .sidebar-sub {{ font-family:'IBM Plex Mono',monospace; font-size:10px; color:rgba(255,255,255,0.4); margin-top:3px; }}
    .sidebar-nav {{ flex:1; padding:16px 10px; overflow-y:auto; }}
    .nav-group-label {{
      font-family:'IBM Plex Mono',monospace; font-size:8px; letter-spacing:.2em;
      text-transform:uppercase; color:rgba(255,255,255,0.25);
      padding:0 10px; margin:14px 0 6px;
    }}
    .nav-item {{
      display:flex; align-items:center; gap:8px; padding:9px 10px; border-radius:5px;
      font-size:12.5px; color:rgba(255,255,255,0.55); cursor:pointer;
      transition:background .15s,color .15s; margin-bottom:1px; user-select:none;
    }}
    .nav-item:hover {{ background:rgba(255,255,255,0.08); color:#fff; }}
    .nav-item.active {{ background:rgba(255,255,255,0.1); color:var(--crimson-l); font-weight:600; }}
    .nav-icon {{ font-size:14px; width:18px; text-align:center; flex-shrink:0; }}
    .sidebar-footer {{ padding:14px 10px; border-top:1px solid rgba(255,255,255,0.1); }}
    .btn-period-info {{
      width:100%; padding:9px 12px; background:rgba(255,255,255,0.06);
      border:1px solid rgba(255,255,255,0.15); border-radius:5px;
      font-family:'IBM Plex Mono',monospace; font-size:9px; letter-spacing:.1em;
      text-transform:uppercase; color:rgba(255,255,255,0.55); text-align:center;
    }}

    /* MAIN */
    .main {{ margin-left:var(--sidebar-w); flex:1; display:flex; flex-direction:column; min-height:100vh; }}

    /* TOPBAR */
    .topbar {{
      background:linear-gradient(90deg,#0D2B5E 0%,#3D0F1D 100%);
      padding:0 28px; height:58px;
      display:flex; align-items:center; justify-content:space-between;
      position:sticky; top:0; z-index:50;
    }}
    .topbar-left {{ display:flex; align-items:center; gap:14px; }}
    .topbar-badge {{
      font-family:'IBM Plex Mono',monospace; font-size:9px; letter-spacing:.15em;
      text-transform:uppercase; color:#fff; background:var(--crimson); padding:3px 10px; border-radius:2px;
    }}
    .topbar-title {{
      font-family:'IBM Plex Mono',monospace; font-size:10px; letter-spacing:.1em;
      text-transform:uppercase; color:rgba(255,255,255,0.45);
    }}
    .topbar-title b {{ color:#fff; font-weight:500; }}
    .topbar-right {{ display:flex; align-items:center; gap:10px; }}
    .btn-export {{
      display:flex; align-items:center; gap:5px; cursor:pointer; border-radius:4px;
      padding:7px 12px; border:1px solid rgba(255,255,255,0.28);
      background:rgba(255,255,255,0.07); color:rgba(255,255,255,0.75);
      font-family:'IBM Plex Mono',monospace; font-size:9px; letter-spacing:.1em;
      text-transform:uppercase; transition:all .2s;
    }}
    .btn-export:hover {{ background:rgba(255,255,255,0.15); color:#fff; }}
    .period-label {{
      font-family:'IBM Plex Mono',monospace; font-size:10px; letter-spacing:.08em;
      color:rgba(255,255,255,0.6); padding:6px 12px;
      border:1px solid rgba(255,255,255,0.15); border-radius:4px;
    }}

    /* TAB BAR */
    .tab-bar {{
      background:var(--bg-card); border-bottom:1px solid var(--border);
      display:flex; padding:0 28px; overflow-x:auto;
      position:sticky; top:58px; z-index:40;
    }}
    .tab-bar::-webkit-scrollbar {{ height:2px; }}
    .tab-bar::-webkit-scrollbar-thumb {{ background:var(--crimson); }}
    .tab-btn {{
      font-family:'IBM Plex Mono',monospace; font-size:9.5px; letter-spacing:.1em;
      text-transform:uppercase; color:var(--muted); padding:0 18px; height:46px;
      display:flex; align-items:center; white-space:nowrap; cursor:pointer;
      border-bottom:2px solid transparent; transition:color .2s,border-color .2s; flex-shrink:0;
      background:none; border-top:none; border-left:none; border-right:none;
    }}
    .tab-btn:hover {{ color:var(--text); }}
    .tab-btn.active {{ color:var(--crimson); border-bottom-color:var(--crimson); }}

    /* CONTENT */
    .content {{ flex:1; padding:32px 28px 80px; }}
    .tab-panel {{ display:none; }}
    .tab-panel.active {{ display:block; }}

    /* PAGE HEADER */
    .page-header {{ margin-bottom:28px; }}
    .page-label {{ font-family:'IBM Plex Mono',monospace; font-size:9px; letter-spacing:.18em; text-transform:uppercase; color:var(--crimson); margin-bottom:7px; }}
    .page-title {{ font-family:'DM Sans',sans-serif; font-size:clamp(18px,2.5vw,26px); font-weight:700; color:var(--text); letter-spacing:-.02em; margin-bottom:5px; }}
    .page-subtitle {{ font-size:13px; color:var(--muted); }}

    /* KPI */
    .kpi-grid {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(170px,1fr)); gap:12px; margin-bottom:28px; }}
    .kpi-card {{
      background:var(--bg-card); border:1px solid var(--border); border-radius:10px;
      padding:16px 18px; position:relative; overflow:hidden;
    }}
    .kpi-card::before {{ content:''; position:absolute; top:0; left:0; right:0; height:2px;
      background:linear-gradient(90deg,var(--crimson),transparent); }}
    .kpi-label {{ font-family:'IBM Plex Mono',monospace; font-size:8.5px; letter-spacing:.15em; text-transform:uppercase; color:var(--muted); margin-bottom:9px; }}
    .kpi-value {{ font-family:'DM Sans',sans-serif; font-size:24px; font-weight:700; color:var(--text); line-height:1; }}
    .kpi-unit {{ font-family:'IBM Plex Mono',monospace; font-size:9px; color:var(--muted); margin-top:4px; }}
    .kpi-trend {{
      display:inline-flex; align-items:center; gap:3px;
      font-family:'IBM Plex Mono',monospace; font-size:9px; letter-spacing:.04em;
      margin-top:7px; padding:2px 6px; border-radius:3px;
    }}
    .kpi-trend.up   {{ background:rgba(15,158,118,0.1); color:var(--green); }}
    .kpi-trend.down {{ background:rgba(123,26,46,0.1);  color:var(--red); }}
    .kpi-trend.flat {{ background:rgba(0,0,0,0.04);     color:var(--muted); }}

    /* CHARTS */
    .charts-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(360px,1fr)); gap:18px; margin-bottom:28px; }}
    .chart-card {{
      background:var(--bg-card); border:1px solid var(--border); border-radius:10px; padding:22px;
    }}
    .chart-card-title {{ font-size:13px; font-weight:700; color:var(--text); margin-bottom:18px; }}
    .chart-wrap {{ position:relative; }}
    .chart-wrap canvas {{ max-height:280px; }}
    .chart-card.full-width {{ grid-column:1/-1; }}

    /* TABLES */
    .table-wrap {{ overflow-x:auto; max-height:400px; overflow-y:auto; }}
    .data-table {{ width:100%; border-collapse:collapse; font-size:12.5px; }}
    .data-table thead th {{
      font-family:'IBM Plex Mono',monospace; font-size:8.5px; letter-spacing:.12em;
      text-transform:uppercase; color:var(--muted); padding:8px 12px;
      border-bottom:1px solid var(--border); background:var(--bg); text-align:left;
      position:sticky; top:0;
    }}
    .data-table tbody tr {{ border-bottom:1px solid rgba(0,0,0,0.04); transition:background .1s; }}
    .data-table tbody tr:hover {{ background:rgba(0,0,0,0.02); }}
    .data-table tbody td {{ padding:8px 12px; color:var(--text); vertical-align:middle; }}

    /* MAP */
    .map-card {{
      background:var(--bg-card); border:1px solid var(--border); border-radius:10px;
      padding:22px; margin-bottom:28px;
    }}
    .map-card-title {{ font-size:13px; font-weight:700; color:var(--text); margin-bottom:18px; }}
    .map-legend {{ display:flex; align-items:center; gap:8px; margin-top:12px; font-family:'IBM Plex Mono',monospace; font-size:9px; color:var(--muted); }}
    .map-legend-gradient {{ width:120px; height:10px; border-radius:3px;
      background:linear-gradient(90deg,#f0e8ea,#7B1A2E); }}

    /* SECTION TITLE */
    .section-title {{ font-size:14px; font-weight:700; color:var(--text); margin:20px 0 12px; }}

    /* MOBILE */
    .hamburger {{ display:none; background:none; border:none; cursor:pointer;
      color:#fff; font-size:22px; padding:4px; }}
    @media(max-width:768px) {{
      .sidebar {{ transform:translateX(-100%); }}
      .sidebar.open {{ transform:translateX(0); }}
      .main {{ margin-left:0; }}
      .hamburger {{ display:block; }}
      .charts-grid {{ grid-template-columns:1fr; }}
      .kpi-grid {{ grid-template-columns:repeat(2,1fr); }}
    }}
  </style>
</head>
<body>
<div class="layout">

  <!-- SIDEBAR -->
  <aside class="sidebar" id="sidebar">
    <div class="sidebar-top">
      <div class="sidebar-logo">FA<span>LAW</span></div>
      <div class="sidebar-badge">Dashboard BI</div>
      <div class="sidebar-client">iFood</div>
      <div class="sidebar-sub">Relatório Processual</div>
    </div>
    <nav class="sidebar-nav">
      <div class="nav-group-label">Visão Geral</div>
      <div class="nav-item active" data-tab="geral">
        <span class="nav-icon">📊</span> Visão Geral
      </div>
      <div class="nav-group-label">Por Causa Raiz</div>
      <div class="nav-item" data-tab="nuvem">
        <span class="nav-icon">☁️</span> Nuvem
      </div>
      <div class="nav-item" data-tab="ol">
        <span class="nav-icon">🚴</span> OL
      </div>
      <div class="nav-item" data-tab="ex_foodlover">
        <span class="nav-icon">⭐</span> Ex-Foodlover
      </div>
      <div class="nav-item" data-tab="terceiriza__o">
        <span class="nav-icon">🏢</span> Terceirização
      </div>
      <div class="nav-item" data-tab="marketplace">
        <span class="nav-icon">🛒</span> Marketplace
      </div>
      <div class="nav-item" data-tab="restaurantes">
        <span class="nav-icon">🍽️</span> Restaurantes
      </div>
      <div class="nav-group-label">Outros</div>
      <div class="nav-item" data-tab="suspensos">
        <span class="nav-icon">⏸️</span> Suspensos
      </div>
      <div class="nav-item" data-tab="decisoes_periodo">
        <span class="nav-icon">⚖️</span> Decisões do Período
      </div>
    </nav>
    <div class="sidebar-footer">
      <div class="btn-period-info">Período: {period_str}</div>
    </div>
  </aside>

  <!-- MAIN -->
  <div class="main">
    <!-- TOPBAR -->
    <header class="topbar">
      <div class="topbar-left">
        <button class="hamburger" id="hamburger">☰</button>
        <span class="topbar-badge">iFood</span>
        <span class="topbar-title">Falaw Advogados — <b>Dashboard Processual</b></span>
      </div>
      <div class="topbar-right">
        <span class="period-label">{period_str}</span>
        <button class="btn-export" id="btn-export-pdf">⬇ Exportar PDF</button>
      </div>
    </header>

    <!-- TAB BAR -->
    <div class="tab-bar">
      <button class="tab-btn active" data-tab="geral">Visão Geral</button>
      {cr_tabs_buttons}
      <button class="tab-btn" data-tab="suspensos">Suspensos</button>
      <button class="tab-btn" data-tab="decisoes_periodo">Decisões do Período</button>
    </div>

    <!-- CONTENT -->
    <div class="content" id="content-area">

      <!-- ═══════════════════════════ VISÃO GERAL ═══════════════════════════ -->
      <div id="tab-geral" class="tab-panel active">
        <div class="page-header">
          <div class="page-label">Dashboard iFood</div>
          <h1 class="page-title">Visão Geral Processual</h1>
          <p class="page-subtitle">Consolidado de todos os processos — Gerado em {datetime.now().strftime("%d/%m/%Y")}</p>
        </div>

        <div class="kpi-grid">{kpi_row1}</div>
        <div class="kpi-grid">{kpi_row2}</div>

        <div class="charts-grid">
          <div class="chart-card">
            <div class="chart-card-title">Distribuição por Causa Raiz (Ativos)</div>
            <div class="chart-wrap"><canvas id="chart-cr-donut"></canvas></div>
          </div>
          <div class="chart-card">
            <div class="chart-card-title">Resultado dos Processos com Decisão</div>
            <div class="chart-wrap"><canvas id="chart-resultados"></canvas></div>
          </div>
        </div>

        <div class="charts-grid">
          <div class="chart-card full-width">
            <div class="chart-card-title">Decisões do Período por Causa Raiz</div>
            <div class="chart-wrap"><canvas id="chart-dec-cr" style="max-height:260px"></canvas></div>
          </div>
        </div>

        <div class="charts-grid">
          <div class="chart-card">
            <div class="chart-card-title">Fase Processual (Ativos)</div>
            <div class="chart-wrap"><canvas id="chart-fase"></canvas></div>
          </div>
          <div class="chart-card">
            <div class="chart-card-title">Top 10 Advogados Adversários</div>
            <div class="chart-wrap"><canvas id="chart-adv"></canvas></div>
          </div>
        </div>

        <div class="charts-grid">
          <div class="chart-card full-width">
            <div class="chart-card-title">Top 15 Julgadores</div>
            <div class="chart-wrap"><canvas id="chart-julgadores" style="max-height:340px"></canvas></div>
          </div>
        </div>

        <!-- TRT CHARTS -->
        <h3 class="section-title" style="margin:0 0 16px">Volumetria por Tribunal Regional do Trabalho</h3>
        <div class="charts-grid">
          <div class="chart-card full-width">
            <div class="chart-card-title">Volume de Processos por TRT (Ativos)</div>
            <div class="chart-wrap" style="max-height:300px"><canvas id="chart-trt-vol"></canvas></div>
          </div>
        </div>
        <div class="charts-grid">
          <div class="chart-card full-width">
            <div class="chart-card-title">Taxa de Decisões Favoráveis por TRT (%)</div>
            <div class="chart-wrap" style="max-height:300px"><canvas id="chart-trt-fav"></canvas></div>
          </div>
        </div>
        <div class="charts-grid">
          <div class="chart-card full-width">
            <div class="chart-card-title">Processos Ativos por TRT e Causa Raiz</div>
            <div class="chart-wrap" style="height:340px"><canvas id="chart-trt-cr"></canvas></div>
          </div>
        </div>
        <div class="chart-card" style="margin-bottom:28px">
          <div class="chart-card-title">Tabela — Ranking TRT</div>
          {trt_table}
        </div>

        <!-- BRAZIL MAP -->
        <div class="map-card" style="position:relative">
          <div class="map-card-title">Distribuição Geográfica por UF (Processos Ativos)</div>
          {BRAZIL_SVG}
          <div class="map-legend">
            <span>0</span>
            <div class="map-legend-gradient"></div>
            <span id="map-max-label">{max_uf}+</span>
            <span style="margin-left:8px;font-size:10px">processos por estado</span>
          </div>
          <div id="map-tooltip" style="display:none;position:absolute;background:#0D2B5E;color:#fff;
            padding:6px 10px;border-radius:4px;font-family:'IBM Plex Mono',monospace;font-size:10px;
            pointer-events:none;white-space:nowrap;z-index:10;box-shadow:0 2px 8px rgba(0,0,0,.25)"></div>
        </div>
      </div>
      <!-- /VISÃO GERAL -->

      <!-- ═══════════════════════════ CAUSA RAIZ TABS ═══════════════════════════ -->
      {cr_tabs_html}

      <!-- ═══════════════════════════ SUSPENSOS ═══════════════════════════ -->
      <div id="tab-suspensos" class="tab-panel">
        <div class="page-header">
          <div class="page-label">Processos Suspensos</div>
          <h1 class="page-title">Suspensos por Tema de Repercussão Geral</h1>
          <p class="page-subtitle">Processos suspensos aguardando julgamento no STF/TST</p>
        </div>
        <div class="kpi-grid">
          <div class="kpi-card">
            <div class="kpi-label">Suspensos Tema 1389</div>
            <div class="kpi-value">{fmt_br(len(suspensos_1389))}</div>
            <div class="kpi-unit">processos</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-label">Suspensos Tema 1291</div>
            <div class="kpi-value">{fmt_br(len(suspensos_1291))}</div>
            <div class="kpi-unit">processos</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-label">Total Suspensos</div>
            <div class="kpi-value">{fmt_br(len(suspensos_1389)+len(suspensos_1291))}</div>
            <div class="kpi-unit">processos</div>
          </div>
        </div>
        <h3 class="section-title">Tema 1389</h3>
        {susp_table_1389}
        <h3 class="section-title" style="margin-top:28px">Tema 1291</h3>
        {susp_table_1291}
      </div>

      <!-- ═══════════════════════════ DECISÕES DO PERÍODO ═══════════════════════════ -->
      <div id="tab-decisoes_periodo" class="tab-panel">
        <div class="page-header">
          <div class="page-label">Decisões do Período</div>
          <h1 class="page-title">Todas as Decisões do Período</h1>
          <p class="page-subtitle">Acórdãos e sentenças coletados no período de referência</p>
        </div>
        <div class="kpi-grid">
          <div class="kpi-card">
            <div class="kpi-label">Total Decisões</div>
            <div class="kpi-value">{fmt_br(len(decisoes))}</div>
            <div class="kpi-unit">no período</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-label">Favoráveis</div>
            <div class="kpi-value" style="color:var(--green)">{fmt_br(dec_fav_count)}</div>
            <div class="kpi-unit">{fmt_pct(dec_fav_count, len(decisoes)) if decisoes else "N/A"}</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-label">Desfavoráveis</div>
            <div class="kpi-value" style="color:var(--red)">{fmt_br(dec_desf_count)}</div>
            <div class="kpi-unit">{fmt_pct(dec_desf_count, len(decisoes)) if decisoes else "N/A"}</div>
          </div>
        </div>
        <div class="kpi-grid">{dec_tipo_summary}</div>
        <div class="chart-card">
          <div class="chart-card-title">Todas as Decisões</div>
          {dec_table}
        </div>
      </div>

    </div><!-- /content -->
  </div><!-- /main -->
</div><!-- /layout -->

<script>
// ── DATA ──────────────────────────────────────────────────────────────────────
const UF_DATA    = {uf_json};
const UF_CR_DATA = {uf_cr_json};
const MAX_UF     = {max_uf};

// ── TAB SWITCHING ────────────────────────────────────────────────────────────
function switchTab(tabId) {{
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn, .nav-item').forEach(b => b.classList.remove('active'));

  const panel = document.getElementById('tab-' + tabId);
  if (panel) panel.classList.add('active');

  document.querySelectorAll('[data-tab="' + tabId + '"]').forEach(b => b.classList.add('active'));
}}

document.querySelectorAll('.tab-btn, .nav-item').forEach(btn => {{
  btn.addEventListener('click', () => {{
    const tab = btn.getAttribute('data-tab');
    if (tab) switchTab(tab);
    if (window.innerWidth <= 768) document.getElementById('sidebar').classList.remove('open');
  }});
}});

// Hamburger
document.getElementById('hamburger').addEventListener('click', () => {{
  document.getElementById('sidebar').classList.toggle('open');
}});

// ── EXPORT PDF ───────────────────────────────────────────────────────────────
document.getElementById('btn-export-pdf').addEventListener('click', async () => {{
  const {{ jsPDF }} = window.jspdf;
  const panel = document.querySelector('.tab-panel.active');
  if (!panel) return;
  const canvas = await html2canvas(panel, {{ scale:1.5, useCORS:true, backgroundColor:'#FAF8F5' }});
  const imgData = canvas.toDataURL('image/png');
  const pdf = new jsPDF({{ orientation:'landscape', unit:'mm', format:'a4' }});
  const pw = pdf.internal.pageSize.getWidth();
  const ph = pdf.internal.pageSize.getHeight();
  const ratio = canvas.width / canvas.height;
  const imgH = pw / ratio;
  let pos = 0;
  if (imgH <= ph) {{
    pdf.addImage(imgData, 'PNG', 0, 0, pw, imgH);
  }} else {{
    let remaining = canvas.height;
    while (remaining > 0) {{
      pdf.addImage(imgData, 'PNG', 0, pos, pw, imgH);
      remaining -= canvas.height * (ph / imgH);
      pos -= ph;
      if (remaining > 0) pdf.addPage();
    }}
  }}
  pdf.save('falaw-ifood-dashboard-{period_str}.pdf');
}});

// ── BRAZIL MAP COLORING ──────────────────────────────────────────────────────
function hexToRgb(hex) {{
  const r = parseInt(hex.slice(1,3),16);
  const g = parseInt(hex.slice(3,5),16);
  const b = parseInt(hex.slice(5,7),16);
  return [r,g,b];
}}
function interpolateColor(t) {{
  // white(#f0e8ea) → crimson(#7B1A2E)
  const c0 = [240,232,234], c1 = [123,26,46];
  const r = Math.round(c0[0] + (c1[0]-c0[0])*t);
  const g = Math.round(c0[1] + (c1[1]-c0[1])*t);
  const b = Math.round(c0[2] + (c1[2]-c0[2])*t);
  return `rgb(${{r}},${{g}},${{b}})`;
}}

const mapTooltip = document.getElementById('map-tooltip');
document.querySelectorAll('#states path').forEach(path => {{
  const uf = path.id;
  const count = UF_DATA[uf] || 0;
  const cr    = UF_CR_DATA[uf] || '—';
  const t = MAX_UF > 0 ? count / MAX_UF : 0;
  path.style.fill = interpolateColor(t);
  path.style.cursor = 'pointer';
  path.addEventListener('mouseenter', (e) => {{
    path.style.opacity = '0.75';
    path.style.strokeWidth = '2';
    if (mapTooltip) {{
      mapTooltip.textContent = uf + ' · ' + count + ' processos · ' + cr;
      mapTooltip.style.display = 'block';
    }}
  }});
  path.addEventListener('mousemove', (e) => {{
    if (!mapTooltip) return;
    const rect = mapTooltip.parentElement.getBoundingClientRect();
    mapTooltip.style.left = (e.clientX - rect.left + 12) + 'px';
    mapTooltip.style.top  = (e.clientY - rect.top  - 28) + 'px';
  }});
  path.addEventListener('mouseleave', () => {{
    path.style.opacity = '1';
    path.style.strokeWidth = '1';
    if (mapTooltip) mapTooltip.style.display = 'none';
  }});
}});

// ── CHARTS ───────────────────────────────────────────────────────────────────
Chart.defaults.font.family = "'DM Sans', sans-serif";
Chart.defaults.color = '#1A0A0D';

const PALETTE_CR = ['#0D2B5E','#7B1A2E','#0F9E76','#f59e0b','#6366f1','#ec4899'];

// Visão Geral: Causa Raiz Doughnut
new Chart(document.getElementById('chart-cr-donut'), {{
  type: 'doughnut',
  data: {{
    labels: {cr_labels_js},
    datasets: [{{ data: {cr_data_js}, backgroundColor: PALETTE_CR, borderWidth: 2, borderColor: '#fff' }}]
  }},
  options: {{
    responsive: true,
    plugins: {{
      legend: {{ position:'right', labels:{{ font:{{ family:'IBM Plex Mono', size:10 }}, padding:14, color:'#1A0A0D' }} }}
    }}
  }}
}});

// Visão Geral: Resultados Doughnut
new Chart(document.getElementById('chart-resultados'), {{
  type: 'doughnut',
  data: {{
    labels: {res_labels_js},
    datasets: [{{ data: {res_data_js},
      backgroundColor: ['#0F9E76','#7B1A2E','#f59e0b','#6366f1'],
      borderWidth: 2, borderColor: '#fff' }}]
  }},
  options: {{
    responsive: true,
    plugins: {{
      legend: {{ position:'right', labels:{{ font:{{ family:'IBM Plex Mono', size:10 }}, padding:14, color:'#1A0A0D' }} }}
    }}
  }}
}});

// Visão Geral: Decisões por CR grouped bar
new Chart(document.getElementById('chart-dec-cr'), {{
  type: 'bar',
  data: {{
    labels: {dec_cr_labels_js},
    datasets: [
      {{ label:'Favorável',    data: {dec_cr_fav_js},  backgroundColor:'#0F9E76', borderRadius:3 }},
      {{ label:'Desfavorável', data: {dec_cr_desf_js}, backgroundColor:'#7B1A2E', borderRadius:3 }}
    ]
  }},
  options: {{
    responsive: true,
    plugins: {{ legend:{{ position:'top', labels:{{ font:{{ family:'IBM Plex Mono', size:10 }} }} }} }},
    scales: {{ x:{{ grid:{{ display:false }} }}, y:{{ grid:{{ color:'rgba(0,0,0,0.05)' }}, beginAtZero:true }} }}
  }}
}});

// Visão Geral: Fase Processual
new Chart(document.getElementById('chart-fase'), {{
  type: 'bar',
  data: {{
    labels: {fase_labels_js},
    datasets: [{{ label:'Processos', data: {fase_data_js}, backgroundColor:'#0D2B5E', borderRadius:4 }}]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true,
    plugins: {{ legend:{{ display:false }} }},
    scales: {{ x:{{ grid:{{ color:'rgba(0,0,0,0.05)' }} }}, y:{{ grid:{{ display:false }} }} }}
  }}
}});

// Visão Geral: Top 10 Advogados
new Chart(document.getElementById('chart-adv'), {{
  type: 'bar',
  data: {{
    labels: {adv_labels_js},
    datasets: [{{ label:'Processos', data: {adv_data_js}, backgroundColor:'#7B1A2E', borderRadius:4 }}]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true,
    plugins: {{ legend:{{ display:false }} }},
    scales: {{ x:{{ grid:{{ color:'rgba(0,0,0,0.05)' }} }}, y:{{ grid:{{ display:false }}, ticks:{{ font:{{ size:9 }} }} }} }}
  }}
}});

// Visão Geral: Top 15 Julgadores
new Chart(document.getElementById('chart-julgadores'), {{
  type: 'bar',
  data: {{
    labels: {julg_labels_js},
    datasets: [{{ label:'Decisões', data: {julg_data_js}, backgroundColor:'#6366f1', borderRadius:4 }}]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true,
    plugins: {{ legend:{{ display:false }} }},
    scales: {{ x:{{ grid:{{ color:'rgba(0,0,0,0.05)' }} }}, y:{{ grid:{{ display:false }}, ticks:{{ font:{{ size:9 }} }} }} }}
  }}
}});

// ── TRT CHARTS ───────────────────────────────────────────────────────────────
new Chart(document.getElementById('chart-trt-vol'), {{
  type: 'bar',
  data: {{
    labels: {trt_chart_labels},
    datasets: [{{ label:'Processos', data: {trt_chart_vol},
      backgroundColor: PALETTE_CR.map((_, i) => i % 2 === 0 ? '#0D2B5E' : '#7B1A2E'),
      borderRadius: 4 }}]
  }},
  options: {{
    responsive: true,
    plugins: {{ legend:{{ display:false }},
      tooltip:{{ callbacks:{{ label: ctx => ' ' + ctx.parsed.y + ' processos' }} }} }},
    scales: {{
      x:{{ grid:{{ display:false }}, ticks:{{ font:{{ family:'IBM Plex Mono', size:10 }} }} }},
      y:{{ grid:{{ color:'rgba(0,0,0,0.05)' }}, beginAtZero:true }}
    }}
  }}
}});

new Chart(document.getElementById('chart-trt-fav'), {{
  type: 'bar',
  data: {{
    labels: {trt_chart_labels},
    datasets: [{{ label:'% Favorável', data: {trt_chart_fav},
      backgroundColor: {trt_chart_fav}.map(v => v >= 30 ? '#0F9E76' : v >= 15 ? '#f59e0b' : '#7B1A2E'),
      borderRadius: 4 }}]
  }},
  options: {{
    responsive: true,
    plugins: {{ legend:{{ display:false }},
      tooltip:{{ callbacks:{{ label: ctx => ' ' + ctx.parsed.y.toFixed(1) + '% favorável' }} }} }},
    scales: {{
      x:{{ grid:{{ display:false }}, ticks:{{ font:{{ family:'IBM Plex Mono', size:10 }} }} }},
      y:{{ max:100, grid:{{ color:'rgba(0,0,0,0.05)' }}, beginAtZero:true,
        ticks:{{ callback: v => v + '%' }} }}
    }}
  }}
}});

// ── TRT × CAUSA RAIZ ─────────────────────────────────────────────────────────
new Chart(document.getElementById('chart-trt-cr'), {{
  type: 'bar',
  data: {{ labels: {trt_cr_labels_js}, datasets: {trt_cr_datasets_js} }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{ legend:{{ position:'bottom', labels:{{ font:{{ family:'IBM Plex Mono', size:10 }}, padding:12 }} }} }},
    scales: {{
      x:{{ stacked:true, grid:{{ display:false }}, ticks:{{ font:{{ family:'IBM Plex Mono', size:10 }} }} }},
      y:{{ stacked:true, grid:{{ color:'rgba(0,0,0,0.05)' }}, beginAtZero:true }}
    }}
  }}
}});

// ── CAUSA RAIZ CHARTS ────────────────────────────────────────────────────────
{cr_js}

</script>
</body>
</html>"""

    return html


# ─── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) >= 2:
        excel_path = sys.argv[1]
    else:
        excel_path = DEFAULT_EXCEL
        if not os.path.exists(excel_path):
            # Search current dir
            candidates = [f for f in os.listdir(".") if f.upper().endswith(".XLSX") and "IFOOD" in f.upper()]
            if candidates:
                excel_path = candidates[0]
                print(f"[INFO] Usando arquivo encontrado: {excel_path}")
            else:
                print(f"[ERROR] Arquivo não encontrado: {excel_path}")
                print("Uso: python processar_ifood.py 'caminho/para/RELATORIO IFOOD.xlsx'")
                sys.exit(1)

    if not os.path.exists(excel_path):
        print(f"[ERROR] Arquivo não encontrado: {excel_path}")
        sys.exit(1)

    # Load
    dfs = load_excel(excel_path)

    # Identify sheets (case-insensitive, multiple keyword fallbacks)
    def find_sheet(*keywords):
        """Return first sheet whose name contains any of the given keywords (case-insensitive)."""
        for kw in keywords:
            for k in dfs:
                if kw.lower() in k.lower():
                    print(f"[INFO] Aba identificada: '{k}' (keyword: {kw!r})")
                    return dfs[k]
        available = list(dfs.keys())
        print(f"[WARN] Nenhuma aba encontrada para keywords {keywords}. Abas disponíveis: {available}")
        return pd.DataFrame()

    df_ifood    = find_sheet("todos", "ifood", "processos", "base", "relatório")
    df_decisoes = find_sheet("decis", "decisão", "decisao", "julgamento", "acórdão", "acordao")
    df_novos    = find_sheet("novos casos recebidos", "novos", "novo caso", "entrada", "ajuizamento")
    df_susp1389 = find_sheet("1389", "susp 1389", "suspensos 1389", "tema 1389")
    df_susp1291 = find_sheet("1291", "susp 1291", "suspensos 1291", "tema 1291")

    print("[INFO] Processando aba IFOOD...")
    records = parse_ifood_sheet(df_ifood)
    print(f"[INFO] {len(records)} registros processados")

    print("[INFO] Processando aba Decisões...")
    decisoes = parse_decisoes_sheet(df_decisoes)
    print(f"[INFO] {len(decisoes)} decisões")

    print("[INFO] Processando Novos Casos...")
    novos_casos = parse_novos_casos(df_novos)
    print(f"[INFO] {len(novos_casos)} novos casos")

    print("[INFO] Processando Suspensos...")
    suspensos_1389 = parse_suspensos(df_susp1389)
    suspensos_1291 = parse_suspensos(df_susp1291)
    print(f"[INFO] {len(suspensos_1389)} suspensos Tema 1389, {len(suspensos_1291)} Tema 1291")

    # Compute analytics
    print("[INFO] Calculando KPIs...")
    kpis = compute_kpis(records, suspensos_1389, suspensos_1291, novos_casos, decisoes)

    print("[INFO] Calculando gráficos globais...")
    global_charts = compute_global_charts(records, decisoes)

    print("[INFO] Calculando dados por Causa Raiz...")
    causa_raiz_data = {}
    for causa, label in CAUSA_RAIZ_TABS:
        causa_raiz_data[causa] = compute_causa_raiz_kpis(
            records, decisoes, causa, suspensos_1389, suspensos_1291
        )
        print(f"  > {causa}: {causa_raiz_data[causa]['total']} processos, "
              f"{causa_raiz_data[causa]['dec_fav']} fav, {causa_raiz_data[causa]['dec_desf']} desf")

    # Period string
    period_str = datetime.now().strftime("%m/%Y")
    period_ym  = datetime.now().strftime("%Y%m")

    # Generate HTML
    print("[INFO] Gerando HTML...")
    html = build_html(
        kpis=kpis,
        global_charts=global_charts,
        causa_raiz_data=causa_raiz_data,
        suspensos_1389=suspensos_1389,
        suspensos_1291=suspensos_1291,
        decisoes=decisoes,
        period_str=period_str,
    )

    # Output path
    out_dir  = os.path.dirname(os.path.abspath(excel_path))
    out_file = os.path.join(out_dir, f"falaw-ifood-dashboard-{period_ym}.html")
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n[OK] Dashboard gerado: {out_file}")
    print(f"     Tamanho: {os.path.getsize(out_file):,} bytes")
    print(f"\n     KPIs Resumo:")
    print(f"       Ativos:    {kpis['ativos']:>6,}")
    print(f"       Encerrados:{kpis['encerrados']:>6,}")
    print(f"       Arquivados:{kpis['arquivados']:>6,}")
    print(f"       Suspensos: {kpis['suspensos']:>6,}")
    print(f"       Dec.Fav:   {kpis['dec_fav']:>6,}")
    print(f"       Dec.Desf:  {kpis['dec_desf']:>6,}")


if __name__ == "__main__":
    main()
