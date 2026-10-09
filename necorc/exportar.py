"""Tabelas de resultado em pandas (para a tela e para exportar em Excel)."""
from __future__ import annotations

import io

import pandas as pd

from .leitura import Entrada, Parametros
from .modelo import resumo_acoes, resumo_secretarias


def tabela_secretarias(entrada: Entrada, linhas: list[dict], p: Parametros) -> pd.DataFrame:
    """K11:P18 + K20:P20."""
    colunas = [str(c) for c in entrada.nomes_secretarias_resumo] + ["MCID"]
    dados, indice = [], []
    for item in resumo_secretarias(entrada, linhas, p):
        dados.append(item["valores"] + [item["total"]])
        indice.append(item["rotulo"])
    return pd.DataFrame(dados, index=indice, columns=colunas)


def tabela_acoes(entrada: Entrada, linhas: list[dict]) -> tuple[pd.DataFrame, dict]:
    """K25:P47."""
    ra = resumo_acoes(entrada, linhas)
    df = pd.DataFrame([{
        "Descrição ação": it["descricao"],
        "cod ação": str(it["codigo"]),
        "nec orc final": it["nec_orc_final"],
        "disponível LOA": it["disponivel_loa"],
        "Dentro disp": it["dentro_disp"],
        "Saldo": it["saldo"],
    } for it in ra["itens"]])
    totais = {"TOTAL nec orc final": ra["total_nec"], "TOTAL disponível LOA": ra["total_disp"],
              "Faltando": ra["faltando"], "Sobrando": ra["sobrando"]}
    return df, totais


def tabela_detalhe(linhas: list[dict]) -> pd.DataFrame:
    return pd.DataFrame(linhas)


def gerar_excel(entrada: Entrada, linhas: list[dict], p: Parametros) -> bytes:
    buf = io.BytesIO()
    acoes, totais = tabela_acoes(entrada, linhas)
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        tabela_secretarias(entrada, linhas, p).to_excel(xw, sheet_name="resumo secretarias")
        acoes.to_excel(xw, sheet_name="resumo ações", index=False)
        pd.DataFrame(list(totais.items()), columns=["item", "valor"]).to_excel(
            xw, sheet_name="resumo ações", index=False, startrow=len(acoes) + 2)
        par = [("geral", k, v) for k, v in vars(p.geral).items()]
        for sec, regra in p.secretaria.items():
            par += [(sec, k, v) for k, v in vars(regra).items()]
            par.append((sec, "bln_secretaria", p.bln_secretaria[sec]))
        par.append(("geral", "bln_execucao", p.bln_execucao))
        pd.DataFrame(par, columns=["regra", "parâmetro", "valor"]).to_excel(
            xw, sheet_name="parâmetros usados", index=False)
        tabela_detalhe(linhas).to_excel(xw, sheet_name="detalhe contratos", index=False)
    return buf.getvalue()
