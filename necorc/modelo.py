"""Modelo preditivo: reescrita, linha a linha, das colunas calculadas da tabela
view_preditivo_orc_pac_2026 (colunas M e AN até BV) e das tabelas resumo da aba
"parametros e resultados gerais".

Cada função/trecho indica a coluna do Excel que reproduz. A ordem dos testes (SE aninhados)
e a ordem das multiplicações seguem exatamente as fórmulas originais.

Onde a fórmula do Excel é uma soma/subtração "pura" (sem SE em volta), usamos sub()/soma(),
que reproduzem o arredondamento para zero que o Excel faz nesses casos (ver excel.sub).
"""
from __future__ import annotations

from typing import Any

from .excel import (NA, ErroExcel, Procv, arred, corresp_exata, diferente, igual, maximo, n,
                    soma, somases, sub)
from .leitura import SECRETARIAS, Entrada, Parametros, Regra

COLUNAS_CALCULADAS = [
    "ação ajustada", "empenhado_SIAFI", "empenho diferença", "vlr_empenhado aj",
    "exec mensal R$ referencia", "média utilizada", "prc_exec_12m", "media mensal %", "bln exec",
    "a_desembolsar", "a empenhar atual", "disponível RAP", "a executar total",
    "classificação 2025", "vlr disponível", "Ritmo de execução parâmetros", "Coluna1",
    "% exec mensal parâmetros", "necessidade orçamentária preliminar", "ritmo parametros SEMOB",
    "ritmo parametros SNSA", "ritmo parametros SNP", "ritmo parametros SNH",
    "ritmo parametros SNs", "% exec mensal parâmetros SNs",
    "necessidade preliminar regra secretaria", "critério inicial aplicado",
    "necessidade orçamento inicial", "Nec. Financeira CAIXA", "vlr disponível pós NecFin CAIXA",
    "necessidade orçamento pós NecFin", "bln simplificado", "Simplificado inicial",
    "simplificado final", "necessidade orçamento final", "bln_execução (situação obra)",
]

EXEC_100 = "1_100% executado"
DESEMB_100 = "2_100% desembolsado"
EMPENH_100 = "3_100% empenhado"
SEM_EXEC = "Sem exec 12 meses"
COM_EXEC = "Com exec 12 meses"


def _ritmo(r: dict, regra: Regra) -> float:
    """Núcleo comum das colunas de ritmo (BB e BF:BI) depois do teste 'a executar total > 0'."""
    repasse, ref, mm = n(r["vlr_repasse"]), r["exec mensal R$ referencia"], r["media mensal %"]
    if igual(r["bln exec"], "Não"):
        return repasse * regra.piso * regra.meses
    if igual(r["bln exec"], "Sim") and mm > regra.teto:
        return repasse * regra.meses * regra.teto
    if igual(r["bln exec"], "Sim") and mm < regra.piso:
        return repasse * regra.meses * regra.piso
    if igual(r["bln exec"], "Sim"):
        return ref * regra.meses * (1 + regra.margem_exec)
    return 0


def calcular_linha(e: dict[str, Any], p: Parametros, lk: dict[str, Procv]) -> dict[str, Any]:
    g = p.geral
    r = dict(e)

    # M  ação ajustada = SEERRO(PROCV(acao_orcamentaria; acao_ajustada!A1:B29;2;0);
    #                           PROCV(cod_tci; acao_ajustada!A29:D32;2;0))
    try:
        r["ação ajustada"] = lk["acao"](r["acao_orcamentaria"])
    except ErroExcel:
        try:
            r["ação ajustada"] = lk["ajustes"](r["cod_tci"])
        except ErroExcel:
            r["ação ajustada"] = NA

    repasse = n(r["vlr_repasse"])
    # AN empenhado_SIAFI = SEERRO(PROCV(cod_tci; SIAFI!B5:E5000; 4; 0); 0)
    try:
        r["empenhado_SIAFI"] = lk["siafi"](r["cod_tci"])
    except ErroExcel:
        r["empenhado_SIAFI"] = 0
    # AO
    r["empenho diferença"] = sub(n(r["empenhado_SIAFI"]), n(r["empenhado_2026"]))
    # AP
    if igual(r["fase_pac"], "NOVO PAC - Seleção"):
        r["vlr_empenhado aj"] = n(r["vlr_empenhado"])
    else:
        r["vlr_empenhado aj"] = n(r["vlr_empenhado"]) - n(r["empenhado_2026"]) + n(r["empenhado_SIAFI"])
    # AQ
    medias = [r["media_mensal_12m"], r["media_mensal_06m"], r["media_mensal_03m"]]
    ref = maximo(medias)
    r["exec mensal R$ referencia"] = ref
    # AR
    try:
        r["média utilizada"] = {1: 12, 2: 6, 3: 3}.get(corresp_exata(ref, medias), "")
    except ErroExcel:
        r["média utilizada"] = NA
    # AS
    if ref == 0:
        r["prc_exec_12m"] = 0
    elif ref * 12 >= repasse:
        r["prc_exec_12m"] = 1
    else:
        r["prc_exec_12m"] = arred((ref * 12) / repasse, 4)
    # AT  (vlr_repasse = 0 ou vazio -> 0)
    r["media mensal %"] = 0 if repasse == 0 else arred(ref / repasse, 4)
    mm = r["media mensal %"]
    # AU
    r["bln exec"] = "sim" if ref > 0 else "não"
    # AV
    a_desemb = sub(repasse, n(r["vlr_desembolsado"]))
    r["a_desembolsar"] = a_desemb
    # AW
    emp_aj = r["vlr_empenhado aj"]
    a_emp = 0 if emp_aj > repasse else repasse - emp_aj
    r["a empenhar atual"] = a_emp
    # AX
    r["disponível RAP"] = 0 if emp_aj - n(r["vlr_desembolsado"]) <= 0 else emp_aj - n(r["vlr_desembolsado"])
    # AY
    a_exec = sub(repasse, n(r["vlr_desbloqueado"]))
    r["a executar total"] = a_exec
    # AZ
    if a_exec < 1:
        cls = EXEC_100
    elif a_desemb <= 1:
        cls = DESEMB_100
    elif a_emp < 1:
        cls = EMPENH_100
    elif a_exec > 0 and mm == 0:
        cls = SEM_EXEC
    elif a_exec > 0 and mm > 0:
        cls = COM_EXEC
    else:
        cls = "nao previsto"
    r["classificação 2025"] = cls
    # BA
    disp = soma(n(r["vlr_saldo_conta"]), r["disponível RAP"])
    r["vlr disponível"] = disp

    # BB  Ritmo de execução parâmetros (regra geral)
    ritmo_g = 0 if a_exec <= 0 else _ritmo(r, g)
    r["Ritmo de execução parâmetros"] = ritmo_g
    # BC  Coluna1 (rótulo da situação)
    if a_exec <= 0:
        r["Coluna1"] = "situacao0"
    elif igual(r["bln exec"], "Não"):
        r["Coluna1"] = "situacao1"
    elif igual(r["bln exec"], "Sim") and mm > g.teto:
        r["Coluna1"] = "situação2"
    elif igual(r["bln exec"], "Sim") and mm < g.piso:
        r["Coluna1"] = "situação3"
    elif igual(r["bln exec"], "Sim"):
        r["Coluna1"] = "situação4"
    else:
        r["Coluna1"] = "situação5"
    # BD
    pct_g = 0 if repasse == 0 else arred((ritmo_g / repasse) / g.meses, 4)
    r["% exec mensal parâmetros"] = pct_g
    # BE  necessidade orçamentária preliminar (regra geral)
    if igual(cls, EXEC_100) or igual(cls, DESEMB_100) or igual(cls, EMPENH_100):
        be = 0
    elif igual(cls, SEM_EXEC) and a_exec > 0:
        be = ritmo_g - disp
    elif igual(cls, COM_EXEC) and a_exec > 0 and ritmo_g - disp >= a_emp:
        be = a_emp
    elif igual(cls, COM_EXEC) and a_exec > 0 and pct_g < g.piso:
        be = (ref * g.piso * g.meses) - disp
    elif igual(cls, COM_EXEC) and a_exec > 0 and pct_g > g.teto:
        be = (ref * g.teto * g.meses) - disp
    else:
        be = ritmo_g - disp
    r["necessidade orçamentária preliminar"] = be

    # BF:BI  ritmo por secretaria (regra específica se bln_<secretaria> = VERDADEIRO)
    for sec in SECRETARIAS:
        col = f"ritmo parametros {sec}"
        if a_exec <= 0 or not igual(r["secretaria"], sec):
            r[col] = 0
        else:
            regra = p.secretaria[sec] if p.bln_secretaria[sec] else g
            r[col] = _ritmo(r, regra)
    # BJ
    ritmo_sn = 0
    for sec in SECRETARIAS:
        ritmo_sn += r[f"ritmo parametros {sec}"]
    r["ritmo parametros SNs"] = ritmo_sn
    # BK
    r["% exec mensal parâmetros SNs"] = 0 if repasse == 0 else arred((ritmo_sn / repasse) / g.meses, 4)
    # BL  necessidade preliminar regra secretaria
    if igual(cls, EXEC_100) or igual(cls, DESEMB_100) or igual(cls, EMPENH_100):
        bl = 0
    elif igual(cls, COM_EXEC) and a_exec > 0 and ritmo_sn - disp >= a_emp:
        bl = a_emp
    else:
        bl = ritmo_sn - disp
    r["necessidade preliminar regra secretaria"] = bl
    # BM  critério inicial aplicado (texto explicativo, baseado na regra geral)
    if igual(cls, EXEC_100):
        crit = "R$0, pois já está 100% executado"
    elif igual(cls, DESEMB_100):
        crit = "R$0, pois já está 100% desembolsado"
    elif igual(cls, EMPENH_100):
        crit = "R$0, pois já está 100% empenhado"
    elif igual(cls, SEM_EXEC) and a_exec > 0:
        crit = "Ritmo definido para sem execução menos valor disponível"
    elif igual(cls, COM_EXEC) and a_exec > 0 and ritmo_g - disp >= a_emp:
        crit = " Valor limitado pelo a empenhar do contrato"
    elif igual(cls, COM_EXEC) and a_exec > 0 and pct_g < g.piso:
        crit = "Regra do piso mínimo menos vlr disponível"
    elif igual(cls, COM_EXEC) and a_exec > 0 and pct_g > g.teto:
        crit = "Regra do teto máximo menos vlr disponível"
    else:
        crit = "Ritmo de execução estimada menos vlr disponível"
    r["critério inicial aplicado"] = crit
    # BN
    inicial = bl if bl >= 0 else 0
    r["necessidade orçamento inicial"] = inicial
    # BO  Nec. Financeira CAIXA = SEERRO(PROCV(cod_operacao num; 'NecFin CAIXA'!A3:M5000; 13; 0); 0)
    try:
        necfin = lk["necfin"](r["cod_operacao num"])
    except ErroExcel:
        necfin = 0
    r["Nec. Financeira CAIXA"] = necfin
    # BP
    disp_pos = sub(disp, n(necfin))
    r["vlr disponível pós NecFin CAIXA"] = disp_pos
    # BQ
    if necfin == 0:
        pos = inicial
    elif necfin > 0 and inicial >= disp_pos:
        pos = inicial - disp_pos
    elif necfin > 0 and inicial < disp_pos:
        pos = 0
    elif necfin > 0 and disp_pos < 0:
        pos = inicial + abs(disp_pos)
    else:
        pos = 0
    r["necessidade orçamento pós NecFin"] = pos
    # BR  bln simplificado
    inv = n(r["vlr_investimento"])
    sit_c, sit_o = r["dsc_situacao_contrato_mcid"], r["dsc_situacao_objeto_mcid"]
    if inv > 1500000:
        bs = "Não"
    elif diferente(r["fase_pac"], "Novo PAC - Seleção"):
        bs = "Não"
    elif igual(sit_c, "Contratado - Suspensiva") and igual(sit_o, "Não iniciada"):
        bs = "Sim_porém sem execução"
    elif igual(sit_c, "Contratado - Suspensiva e Liminar") and igual(sit_o, "Não iniciada"):
        bs = "Sim_porém sem execução"
    elif igual(sit_c, "Contratado - Suspensiva") and igual(sit_o, "Em Execução"):
        bs = "Sim"
    elif igual(sit_c, "Contratado - Suspensiva e Liminar") and igual(sit_o, "Em Execução"):
        bs = "Sim"
    elif diferente(sit_c, "Contratado - Suspensiva") and igual(sit_o, "Em Execução"):
        bs = "Sim"
    elif igual(sit_c, "Contratado - Normal") and igual(sit_o, "Não iniciada"):
        bs = "Sim"
    else:
        bs = "verificar"
    r["bln simplificado"] = bs
    # BS  (regra geral: simplif)
    simp_ini = repasse * g.simplif
    r["Simplificado inicial"] = simp_ini
    # BT
    emp = n(r["vlr_empenhado"])
    if igual(bs, "Não") or igual(bs, "Sim_porém sem execução"):
        simp = 0
    elif igual(bs, "Sim") and a_emp == 0:
        simp = 0
    elif igual(bs, "Sim") and emp > simp_ini:
        simp = 0
    elif igual(bs, "Sim") and emp < simp_ini and simp_ini - emp >= a_emp:
        simp = a_emp
    elif igual(bs, "Sim") and emp < simp_ini and simp_ini - emp < a_emp:
        simp = simp_ini - emp
    else:
        simp = 0
    r["simplificado final"] = simp
    # BU
    r["necessidade orçamento final"] = max(pos, simp)
    # BV
    r["bln_execução (situação obra)"] = bool(p.bln_execucao is True and igual(sit_o, "Em execução"))
    return r


def calcular(entrada: Entrada, p: Parametros | None = None) -> list[dict[str, Any]]:
    """Calcula todas as linhas da view com os parâmetros informados (ou os da planilha)."""
    p = p or entrada.parametros
    lk = {
        "acao": Procv(entrada.acao_tabela),
        "ajustes": Procv(entrada.acao_ajustes),
        "siafi": Procv(entrada.siafi),
        "necfin": Procv(entrada.necfin),
    }
    return [calcular_linha(e, p, lk) for e in entrada.linhas]


# ---------------------------------------------------------------------------
# Tabelas resumo da aba "parametros e resultados gerais"
# ---------------------------------------------------------------------------
LINHAS_RESUMO = [  # (linha Excel, rótulo coluna K, coluna somada da view)
    (12, "Ritmo de execução", "ritmo parametros SNs"),
    (13, "Valor disponível", "vlr disponível"),
    (14, "Necessidade orçamentária preliminar", "necessidade orçamento inicial"),
    (15, "Necessidade Financeira CAIXA", "Nec. Financeira CAIXA"),
    (16, "Necessidade orçamentária após CAIXA", "necessidade orçamento pós NecFin"),
    (17, "Necessidade orçamentária simplificado", "simplificado final"),
    (18, "Necessidade orçamentária final", "necessidade orçamento final"),
]


def _col(linhas, nome):
    return [r[nome] for r in linhas]


def resumo_secretarias(entrada: Entrada, linhas: list[dict], p: Parametros) -> list[dict]:
    """K11:P18 e K20:P20 (valores em milhões)."""
    nomes = entrada.nomes_secretarias_resumo
    sec = _col(linhas, "secretaria")
    out = []
    for lin, rotulo, coluna in LINHAS_RESUMO:
        vals = _col(linhas, coluna)
        valores = [somases(vals, (sec, nome)) / 1000000 for nome in nomes]
        out.append({"linha": lin, "rotulo": rotulo, "valores": valores, "total": _soma(valores)})
    vals = _col(linhas, "necessidade orçamento final")
    blnx = _col(linhas, "bln_execução (situação obra)")
    valores = [somases(vals, (sec, nome), (blnx, p.bln_execucao)) / 1000000 for nome in nomes]
    out.append({"linha": 20, "rotulo": "Nec orçamentária final (bln_execucao)", "valores": valores,
                "total": _soma(valores)})
    return out


def _soma(valores):
    t = 0
    for v in valores:
        t += v
    return t


def resumo_acoes(entrada: Entrada, linhas: list[dict]) -> dict:
    """K25:P47 (valores em milhões)."""
    final = _col(linhas, "necessidade orçamento final")
    acao = _col(linhas, "ação ajustada")
    painel_c = [c for c, _ in entrada.painel]
    painel_h = [h for _, h in entrada.painel]
    itens = []
    for lin, desc, cod in entrada.acoes_resumo:
        nec = somases(final, (acao, cod)) / 1000000
        disp = somases(painel_h, (painel_c, cod)) / 1000000
        itens.append({"linha": lin, "descricao": desc, "codigo": cod, "nec_orc_final": nec,
                      "disponivel_loa": disp, "dentro_disp": "Não" if nec > disp else "Sim",
                      "saldo": disp - nec})
    tot_nec = _soma([i["nec_orc_final"] for i in itens])
    tot_disp = _soma([i["disponivel_loa"] for i in itens])
    faltando = _soma([i["saldo"] for i in itens if i["dentro_disp"] == "Não"])
    sobrando = _soma([i["saldo"] for i in itens if i["dentro_disp"] == "Sim"])
    return {"itens": itens, "total_nec": tot_nec, "total_disp": tot_disp,
            "faltando": faltando, "sobrando": sobrando}
