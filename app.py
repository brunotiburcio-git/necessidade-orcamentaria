"""Front-end (Streamlit) do modelo preditivo de necessidade orçamentária.

Rodar:  streamlit run app.py
"""
from __future__ import annotations

import copy
import json
from datetime import datetime
from pathlib import Path

import streamlit as st

from necorc.exportar import gerar_excel, tabela_acoes, tabela_detalhe, tabela_secretarias
from necorc.leitura import SECRETARIAS, Entrada, Parametros, Regra, ler_planilha
from necorc.modelo import calcular

RAIZ = Path(__file__).resolve().parent
CONFIG = json.loads((RAIZ / "config.json").read_text(encoding="utf-8"))
PASTA = (RAIZ / CONFIG.get("pasta_planilhas", "dados")).resolve()
PADRAO = CONFIG.get("padrao_arquivo", "*.xlsx")

CAMPOS = [  # (atributo, rótulo da aba de parâmetros, linha)
    ("margem_exec", "Margem sobre desempenho atual (com execução)", 4),
    ("simplif", "Regime simplificado (com execução)", 5),
    ("piso", "Piso mínimo (sem execução)", 6),
    ("teto", "Teto máximo (com alta execução)", 7),
    ("meses", "Qtd de meses", 8),
]

st.set_page_config(page_title="Necessidade orçamentária", layout="wide")
st.markdown(f"<style>{(RAIZ / 'assets' / 'style.css').read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)


@st.cache_data(show_spinner="Lendo a planilha...")
def carregar(caminho: str, _mtime: float) -> Entrada:
    return ler_planilha(caminho)


def fmt_br(v, casas=2):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return v


def _num(v):
    """Mantém inteiros como int (ex.: meses = 3) para o widget e para as contas."""
    return int(v) if isinstance(v, (int, float)) and float(v).is_integer() and not isinstance(v, bool) else float(v)


def chave(regra: str, campo: str) -> str:
    return f"p_{regra}_{campo}"


def carregar_parametros_no_estado(p: Parametros):
    for campo, _, _ in CAMPOS:
        conv = _num if campo == "meses" else float
        st.session_state[chave("geral", campo)] = conv(getattr(p.geral, campo))
        for sec in SECRETARIAS:
            st.session_state[chave(sec, campo)] = conv(getattr(p.secretaria[sec], campo))
    for sec in SECRETARIAS:
        st.session_state[chave(sec, "bln")] = p.bln_secretaria[sec]
    st.session_state["p_bln_execucao"] = p.bln_execucao


def parametros_do_estado() -> Parametros:
    def regra(nome):
        return Regra(**{campo: st.session_state[chave(nome, campo)] for campo, _, _ in CAMPOS})
    return Parametros(
        geral=regra("geral"),
        secretaria={sec: regra(sec) for sec in SECRETARIAS},
        bln_secretaria={sec: bool(st.session_state[chave(sec, "bln")]) for sec in SECRETARIAS},
        bln_execucao=bool(st.session_state["p_bln_execucao"]),
    )


def campo_numero(nome_regra: str, campo: str, rotulo: str):
    v = st.session_state[chave(nome_regra, campo)]
    if campo == "meses":
        st.number_input(rotulo, key=chave(nome_regra, campo), min_value=1 if isinstance(v, int) else 0.0,
                        step=1 if isinstance(v, int) else 1.0)
    else:
        st.number_input(rotulo, key=chave(nome_regra, campo), min_value=0.0, step=0.001, format="%.4f",
                        help="Informe como fração: 0,2 = 20%")


# ---------------------------------------------------------------------------
# Barra lateral: escolha da planilha
# ---------------------------------------------------------------------------
st.sidebar.header("Base de dados")
arquivos = sorted((a for a in PASTA.glob(PADRAO) if not a.name.startswith("~$")),
                  key=lambda a: a.stat().st_mtime, reverse=True)
if not arquivos:
    st.sidebar.error(f"Nenhuma planilha encontrada em {PASTA}")
    st.info(f"Salve a planilha atualizada (.xlsx) na pasta **{PASTA}** e recarregue a página. "
            "A pasta pode ser alterada no arquivo config.json.")
    st.stop()

escolhido = st.sidebar.selectbox(
    "Planilha (mais recente primeiro)", arquivos,
    format_func=lambda a: f"{a.name}  ({datetime.fromtimestamp(a.stat().st_mtime):%d/%m/%Y %H:%M})")
if st.sidebar.button("Recarregar planilha"):
    carregar.clear()
entrada = carregar(str(escolhido), escolhido.stat().st_mtime)
st.sidebar.caption(f"{len(entrada.linhas):,} contratos lidos da aba da view".replace(",", "."))

if st.session_state.get("_arquivo") != (str(escolhido), escolhido.stat().st_mtime):
    carregar_parametros_no_estado(entrada.parametros)
    st.session_state["_arquivo"] = (str(escolhido), escolhido.stat().st_mtime)

# ---------------------------------------------------------------------------
# Parâmetros
# ---------------------------------------------------------------------------
st.title("Necessidade orçamentária: modelo preditivo")

with st.expander("Parâmetros", expanded=True):
    cols = st.columns(5)
    with cols[0]:
        st.markdown('<div class="legenda">Regra geral</div>', unsafe_allow_html=True)
        for campo, rotulo, _ in CAMPOS:
            campo_numero("geral", campo, rotulo)
    for col, sec in zip(cols[1:], SECRETARIAS):
        with col:
            st.markdown(f'<div class="legenda">{sec}</div>', unsafe_allow_html=True)
            for campo, rotulo, _ in CAMPOS:
                campo_numero(sec, campo, rotulo)
            st.checkbox("Regra específica", key=chave(sec, "bln"),
                        help=f"bln_{sec.lower()}: marcado = regra específica; desmarcado = regra geral")
    st.checkbox("Somente contratos em execução (bln_execucao)", key="p_bln_execucao",
                help="Afeta a linha 'Nec orçamentária final (bln_execucao)'.")
    st.button("Restaurar parâmetros da planilha",
              on_click=carregar_parametros_no_estado, args=(copy.deepcopy(entrada.parametros),))

parametros = parametros_do_estado()
linhas = calcular(entrada, parametros)

# ---------------------------------------------------------------------------
# Resultados
# ---------------------------------------------------------------------------
st.subheader("Resumo por secretaria (valores em milhões)")
df_sec = tabela_secretarias(entrada, linhas, parametros)
st.dataframe(df_sec.style.format(fmt_br), width="stretch")

st.subheader("Resumo por ação orçamentária (valores em milhões)")
df_acao, totais = tabela_acoes(entrada, linhas)
st.dataframe(
    df_acao.style.format({c: fmt_br for c in ["nec orc final", "disponível LOA", "Saldo"]})
    .map(lambda v: "color: #f87171; font-weight: 600" if v == "Não" else "", subset=["Dentro disp"]),
    width="stretch", hide_index=True, height=35 * (len(df_acao) + 1) + 3)
c1, c2, c3, c4 = st.columns(4)
c1.metric("TOTAL nec orc final", fmt_br(totais["TOTAL nec orc final"]))
c2.metric("TOTAL disponível LOA", fmt_br(totais["TOTAL disponível LOA"]))
c3.metric("Faltando", fmt_br(totais["Faltando"]))
c4.metric("Sobrando", fmt_br(totais["Sobrando"]))

with st.expander("Detalhe por contrato"):
    det = tabela_detalhe(linhas)
    secs = st.multiselect("Secretaria", sorted(det["secretaria"].dropna().unique()))
    if secs:
        det = det[det["secretaria"].isin(secs)]
    misto = [c for c in det.columns if det[c].dtype == object and det[c].map(type).nunique() > 1]
    det[misto] = det[misto].astype(str)  # colunas com número e texto (ex.: ação 8865)
    st.dataframe(det, width="stretch", hide_index=True)

if st.button("Gerar arquivo Excel com os resultados"):
    with st.spinner("Gerando arquivo..."):
        dados = gerar_excel(entrada, linhas, parametros)
    st.download_button("Baixar resultados em Excel", data=dados,
                       file_name=f"necorc_resultado_{datetime.now():%Y%m%d_%H%M}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
