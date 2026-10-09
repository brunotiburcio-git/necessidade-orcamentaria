"""Dados para a versão web (Artifact) do sistema.

A página web recalcula só a parte do modelo que depende dos parâmetros. Tudo o que não depende
deles (colunas da view vindas do banco, PROCVs, classificação 2025, valor disponível, etc.) é
calculado aqui em Python, pelo mesmo modelo.py validado contra o Excel, e vai pronto para a página.

As fórmulas que a página recalcula (motor.js) são as mesmas de modelo.py:
_ritmo (BB/BF:BI), BJ, BL, BN, BQ, BS, BT, BU, BV e as tabelas resumo.
"""
from __future__ import annotations

from datetime import datetime

from .excel import criterio_ok, eh_numero, igual
from .leitura import SECRETARIAS, Entrada, Parametros
from .modelo import COM_EXEC, DESEMB_100, EMPENH_100, EXEC_100, SEM_EXEC, calcular

# Colunas da aba "detalhe contratos" que dependem dos parâmetros: a página recalcula (motor.js).
# As demais vão prontas em "detalhe.fixas", na ordem da view.
COLUNAS_RECALCULADAS = [
    "Ritmo de execução parâmetros", "Coluna1", "% exec mensal parâmetros",
    "necessidade orçamentária preliminar", "ritmo parametros SEMOB", "ritmo parametros SNSA",
    "ritmo parametros SNP", "ritmo parametros SNH", "ritmo parametros SNs",
    "% exec mensal parâmetros SNs", "necessidade preliminar regra secretaria",
    "critério inicial aplicado", "necessidade orçamento inicial", "necessidade orçamento pós NecFin",
    "Simplificado inicial", "simplificado final", "necessidade orçamento final",
    "bln_execução (situação obra)",
]

CLASSES = [EXEC_100, DESEMB_100, EMPENH_100, SEM_EXEC, COM_EXEC]   # índice 5 = outros
BLN_SIMPLIFICADO = ["Não", "Sim_porém sem execução", "Sim"]        # índice 3 = verificar


def _indice(valor, opcoes):
    for i, o in enumerate(opcoes):
        if igual(valor, o):
            return i
    return len(opcoes)


def _num(v, nome):
    if not eh_numero(v):
        raise ValueError(f"valor não numérico em {nome}: {v!r}")
    return v


def _regra(r):
    return {"margem_exec": r.margem_exec, "simplif": r.simplif, "piso": r.piso,
            "teto": r.teto, "meses": r.meses}


def parametros_json(p: Parametros) -> dict:
    return {"geral": _regra(p.geral), "secretaria": {s: _regra(p.secretaria[s]) for s in SECRETARIAS},
            "bln_secretaria": dict(p.bln_secretaria), "bln_execucao": p.bln_execucao}


def exportar(entrada: Entrada) -> dict:
    linhas = calcular(entrada)  # parâmetros da planilha; só as colunas independentes são usadas
    col = {k: [] for k in ("repasse", "ref", "mm", "execSim", "aExec", "cls", "disp", "aEmp",
                           "necfin", "dispPos", "bs", "emp", "emExec", "sec")}
    for r in linhas:
        col["repasse"].append(_num(r["vlr_repasse"] or 0, "vlr_repasse"))
        col["ref"].append(_num(r["exec mensal R$ referencia"], "ref"))
        col["mm"].append(_num(r["media mensal %"], "media mensal %"))
        col["execSim"].append(igual(r["bln exec"], "Sim"))
        col["aExec"].append(_num(r["a executar total"], "a executar total"))
        col["cls"].append(_indice(r["classificação 2025"], CLASSES))
        col["disp"].append(_num(r["vlr disponível"], "vlr disponível"))
        col["aEmp"].append(_num(r["a empenhar atual"], "a empenhar atual"))
        col["necfin"].append(_num(r["Nec. Financeira CAIXA"], "Nec. Financeira CAIXA"))
        col["dispPos"].append(_num(r["vlr disponível pós NecFin CAIXA"], "vlr disponível pós NecFin"))
        col["bs"].append(_indice(r["bln simplificado"], BLN_SIMPLIFICADO))
        col["emp"].append(_num(r["vlr_empenhado"] or 0, "vlr_empenhado"))
        col["emExec"].append(igual(r["dsc_situacao_objeto_mcid"], "Em execução"))
        col["sec"].append(_indice(r["secretaria"], SECRETARIAS))

    nomes = entrada.nomes_secretarias_resumo
    sec_txt = [r["secretaria"] for r in linhas]
    linhas_sec = [[i for i, s in enumerate(sec_txt) if criterio_ok(s, nome)] for nome in nomes]

    painel_c = [c for c, _ in entrada.painel]
    painel_h = [h for _, h in entrada.painel]
    from .excel import somases
    acao = [r["ação ajustada"] for r in linhas]
    acoes = [{"descricao": desc, "codigo": str(cod),
              "disponivel_loa": somases(painel_h, (painel_c, cod)) / 1000000,
              "linhas": [i for i, a in enumerate(acao) if criterio_ok(a, cod)]}
             for _, desc, cod in entrada.acoes_resumo]

    colunas = list(linhas[0].keys()) if linhas else []
    fixas, datas = {}, []
    for nome in colunas:
        if nome in COLUNAS_RECALCULADAS:
            continue
        vals = [r[nome] for r in linhas]
        if any(isinstance(v, datetime) for v in vals):
            datas.append(nome)
            vals = [v.isoformat() if isinstance(v, datetime) else v for v in vals]
        fixas[nome] = vals

    return {
        "arquivo": entrada.arquivo.name,
        "gerado_em": datetime.fromtimestamp(entrada.arquivo.stat().st_mtime).strftime("%d/%m/%Y %H:%M"),
        "contratos": len(linhas),
        "secretarias": SECRETARIAS,
        "nomes_resumo": [str(n) for n in nomes],
        "linhas_resumo": linhas_sec,
        "acoes": acoes,
        "parametros": parametros_json(entrada.parametros),
        "col": col,
        "detalhe": {"colunas": colunas, "fixas": fixas, "datas": datas},
    }
