"""Leitura da planilha atualizada (mesmo arquivo usado hoje no Excel).

O sistema lê apenas os DADOS DE ENTRADA, exatamente como estão salvos:
- aba "view_preditivo_orc_pac_2026": colunas A até AM (view SQL, colunas em verde)
- aba "SIAFI": B5:E5000 (cod_tci -> empenhado_2026)
- aba "NecFin CAIXA": A3:L5000 (a coluna M = MÁXIMO(I:J) é recalculada aqui)
- aba "Painel SPOA - RP3": C8:H50 (ação -> disponível)
- aba "acao_ajustada": A1:B29 (de-para de ações) e A29:D32 (ajustes pontuais)
- aba "parametros e resultados gerais": valores atuais dos parâmetros (usados como padrão)
  e a lista de ações da tabela K25:P47.

As colunas calculadas da view (AN em diante) e os resultados da aba de parâmetros também são
lidos, mas só para a validação (comparar o resultado do Python com o do Excel).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter

ABA_VIEW = "view_preditivo_orc_pac_2026"
ABA_PARAM = "parametros e resultados gerais"
ABA_SIAFI = "SIAFI"
ABA_NECFIN = "NecFin CAIXA"
ABA_PAINEL = "Painel SPOA - RP3"
ABA_ACAO = "acao_ajustada"

# Colunas A..AM da view (dados de entrada vindos do banco)
COLUNAS_ENTRADA = [
    "cod_operacao num", "cod_tci", "cod_contrato", "cod_operacao", "cod_proposta",
    "num_proposta", "num_convenio", "fase_pac", "secretaria", "objeto", "ano_assinatura",
    "acao_orcamentaria", "ação ajustada", "dsc_situacao_contrato_mcid",
    "dsc_situacao_objeto_mcid", "txt_uf", "txt_municipio", "txt_tomador", "txt_tipo_tomador",
    "vlr_investimento", "vlr_repasse", "vlr_empenhado", "vlr_desembolsado", "vlr_desbloqueado",
    "vlr_saldo_conta", "empenhado_2026", "desembolsado_2026", "execucao_2026", "execucao_12m",
    "execucao_06m", "execucao_03m", "media_mensal_12m", "media_mensal_06m", "media_mensal_03m",
    "data_aio", "qtd_dias_aio", "dias_sem_medicao", "mes_referencia", "txt_link",
]
# "ação ajustada" (coluna M) é fórmula no Excel; é recalculada pelo modelo.
COLUNAS_FORMULA_NA_ENTRADA = {"ação ajustada"}

SECRETARIAS = ["SEMOB", "SNSA", "SNP", "SNH"]  # colunas F, G, H, I da aba de parâmetros


@dataclass
class Regra:
    margem_exec: float
    simplif: float
    piso: float
    teto: float
    meses: float


@dataclass
class Parametros:
    geral: Regra
    secretaria: dict[str, Regra]
    bln_secretaria: dict[str, bool]
    bln_execucao: bool
    # Módulo "Geral / Ação" (só no sistema; não existe na planilha): regra específica por ação ajustada.
    # modo = "secretaria" (padrão, igual ao Excel) ou "acao"; chaves = código da ação como texto.
    modo: str = "secretaria"
    acao: dict[str, Regra] = field(default_factory=dict)
    bln_acao: dict[str, bool] = field(default_factory=dict)


@dataclass
class Entrada:
    arquivo: Path
    linhas: list[dict[str, Any]]                  # linhas da view (entradas)
    siafi: list[tuple[Any, Any]]                  # (cod_tci, empenhado_2026)
    necfin: list[tuple[Any, Any]]                 # (OPERACAO, valor_necessidade_max)
    acao_tabela: list[tuple[Any, Any]]            # acao_ajustada!A1:B29
    acao_ajustes: list[tuple[Any, Any]]           # acao_ajustada!A29:B32
    painel: list[tuple[Any, Any]]                 # Painel SPOA C8:C50, H8:H50
    parametros: Parametros                        # valores salvos na planilha
    nomes_secretarias_resumo: list[Any]           # L11:O11
    acoes_resumo: list[tuple[int, Any, Any]]      # (linha, K descrição, L código) de 26 a 45
    acao_secretaria: dict[str, str]               # acao_ajustada!B:D: ação LOA atual -> secretaria
    excel_view: list[dict[str, Any]] = field(default_factory=list)   # valores calculados pelo Excel
    excel_param: dict[str, Any] = field(default_factory=dict)        # células da aba de parâmetros


def _celulas(ws, max_linha: int, max_col: int) -> dict[str, Any]:
    """Valores de A1 até (max_linha, max_col) indexados pelo endereço ("E4", "L11"...)."""
    out = {}
    for i, linha in enumerate(ws.iter_rows(min_row=1, max_row=max_linha, max_col=max_col, values_only=True), start=1):
        for j in range(max_col):
            out[f"{get_column_letter(j + 1)}{i}"] = linha[j] if j < len(linha) else None
    for i in range(1, max_linha + 1):
        for j in range(max_col):
            out.setdefault(f"{get_column_letter(j + 1)}{i}", None)
    return out


def ler_planilha(caminho: str | Path) -> Entrada:
    caminho = Path(caminho)
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    try:
        # ---------------- view ----------------
        ws = wb[ABA_VIEW]
        linhas_brutas = ws.iter_rows(values_only=True)
        cabecalho = None
        for _ in range(50):
            r = next(linhas_brutas)
            if r and r[0] == "cod_operacao num":
                cabecalho = [str(h) if h is not None else None for h in r]
                break
        if cabecalho is None:
            raise ValueError(f"Cabeçalho 'cod_operacao num' não encontrado na aba {ABA_VIEW}")
        faltando = [c for c in COLUNAS_ENTRADA if c not in cabecalho]
        if faltando:
            raise ValueError(f"Colunas ausentes na aba {ABA_VIEW}: {faltando}")
        pos = {nome: i for i, nome in enumerate(cabecalho) if nome is not None}
        linhas, excel_view = [], []
        for r in linhas_brutas:
            r = list(r) + [None] * (len(cabecalho) - len(r))
            entrada = {c: r[pos[c]] for c in COLUNAS_ENTRADA if c not in COLUNAS_FORMULA_NA_ENTRADA}
            if all(v is None for v in entrada.values()):
                break  # fim da tabela
            linhas.append(entrada)
            excel_view.append({nome: r[i] for nome, i in pos.items()})

        # ---------------- SIAFI!B5:E5000 ----------------
        siafi = [(r[1], r[4]) for r in wb[ABA_SIAFI].iter_rows(min_row=5, max_row=5000, max_col=5, values_only=True)
                 if r and len(r) >= 5]

        # ---------------- NecFin CAIXA!A3:M5000 (M = MAX(I:J)) ----------------
        necfin = []
        for i, r in enumerate(wb[ABA_NECFIN].iter_rows(min_row=3, max_row=5000, max_col=13, values_only=True), start=3):
            r = list(r) + [None] * (13 - len(r))
            if i == 3:
                m = r[12]  # linha de cabeçalho (texto)
            elif r[0] is None and r[8] is None and r[9] is None:
                m = r[12]
            else:
                nums = [v for v in (r[8], r[9]) if isinstance(v, (int, float)) and not isinstance(v, bool)]
                m = max(nums) if nums else 0
            necfin.append((r[0], m))

        # ---------------- acao_ajustada ----------------
        ac = [list(r) for r in wb[ABA_ACAO].iter_rows(min_row=1, max_row=32, max_col=4, values_only=True)]
        ac += [[None] * 4] * (32 - len(ac))
        acao_tabela = [(r[0], r[1]) for r in ac[0:29]]     # A1:B29
        acao_ajustes = [(r[0], r[1]) for r in ac[28:32]]   # A29:B32
        acao_secretaria = {str(r[1]): str(r[3]) for r in ac[1:28] if r[1] is not None and r[3] is not None}

        # ---------------- Painel SPOA - RP3!C8:H50 ----------------
        painel = [(r[2], r[7]) for r in wb[ABA_PAINEL].iter_rows(min_row=8, max_row=50, max_col=8, values_only=True)
                  if r and len(r) >= 8]

        # ---------------- parâmetros ----------------
        wp = wb[ABA_PARAM]
        p = _celulas(wp, 50, 16)

        def regra(col: str) -> Regra:
            return Regra(margem_exec=p[f"{col}4"], simplif=p[f"{col}5"], piso=p[f"{col}6"],
                         teto=p[f"{col}7"], meses=p[f"{col}8"])

        parametros = Parametros(
            geral=regra("E"),
            secretaria={s: regra(c) for s, c in zip(SECRETARIAS, "FGHI")},
            bln_secretaria={s: bool(p[f"{c}9"]) for s, c in zip(SECRETARIAS, "FGHI")},
            bln_execucao=bool(p["E10"]),
        )
        nomes = [p[f"{c}11"] for c in "LMNO"]
        acoes = [(lin, p[f"K{lin}"], p[f"L{lin}"]) for lin in range(26, 46) if p[f"L{lin}"] is not None]
    finally:
        wb.close()

    return Entrada(arquivo=caminho, linhas=linhas, siafi=siafi, necfin=necfin,
                   acao_tabela=acao_tabela, acao_ajustes=acao_ajustes, painel=painel,
                   parametros=parametros, nomes_secretarias_resumo=nomes, acoes_resumo=acoes,
                   acao_secretaria=acao_secretaria,
                   excel_view=excel_view, excel_param=p)
