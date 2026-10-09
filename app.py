"""Front-end (Streamlit) do modelo preditivo de necessidade orçamentária.

Rodar:  streamlit run app.py
"""
from __future__ import annotations

import copy
import json
from datetime import datetime
from decimal import Decimal
from pathlib import Path

import streamlit as st

from necorc.exportar import gerar_excel, tabela_acoes, tabela_detalhe, tabela_secretarias
from necorc.leitura import SECRETARIAS, Entrada, Parametros, Regra, ler_planilha
from necorc.modelo import calcular
from necorc.usuarios import conferir, gerar_codigo, ler_usuarios, usuario_do_codigo

RAIZ = Path(__file__).resolve().parent
CONFIG = json.loads((RAIZ / "config.json").read_text(encoding="utf-8"))
PASTA = (RAIZ / CONFIG.get("pasta_planilhas", "dados")).resolve()
PADRAO = CONFIG.get("padrao_arquivo", "*.xlsx")
ARQ_USUARIOS = PASTA / CONFIG.get("arquivo_usuarios", "usuarios.xlsx")

CAMPOS = [  # (atributo, rótulo na tela, explicação mostrada no "?" da regra geral)
    ("margem_exec", "Margem (%)",
     "Percentual de margem a ser aplicado sobre ritmo de execução do contrato.\n\n"
     "**Exemplo:**\n"
     "- Contrato executa 1% ao mês\n"
     "- Usuário escolheu 25% de margem\n"
     "- Usuário escolheu 3 meses no campo Meses\n\n"
     "Então o ritmo de execução prevista será de 1,25%.\n\n"
     "**Resultado:** execução real x (1 + margem) x meses"),
    ("simplif", "Simplificado (%)",
     "Para os contratos do regime simplificado, a necessidade orçamentária prevê um valor padrão de "
     "necessidade orçamentária.\n\n"
     "Ou seja, 50% significa que para os contratos em execução, deve-se garantir 50% do valor do "
     "repasse do contrato."),
    ("piso", "Sem exec/piso (%)",
     "Para os contratos sem execução, será considerado um percentual mínimo do valor de repasse a ser "
     "garantido vezes a quantidade de meses.\n\n"
     "**Exemplo:**\n"
     "- Contrato executou 0% nos últimos 3, 6 ou 12 meses (sem execução)\n"
     "- Usuário escolheu 1% no campo Sem exec/piso\n"
     "- Usuário escolheu 3 meses no campo Meses\n\n"
     "Então o modelo irá prever um valor mínimo de 3% do valor de repasse para cada contrato.\n\n"
     "**Resultado:** valor repasse x sem exec/piso x meses"),
    ("teto", "Teto máximo (%)",
     "Percentual máximo de execução prevista.\n\n"
     "Ou seja, a execução prevista não pode ser superior a 100% ao mês."),
    ("meses", "Qtd de meses",
     "Quantidade de meses para qual o usuário quer projetar a necessidade orçamentária."),
]

# rótulos curtos para os cards estreitos das secretarias
PRINCIPAIS = ("margem_exec", "piso", "meses")  # parâmetros principais do modelo (recebem destaque)

ROTULO_CURTO = {"margem_exec": "Margem (%)", "simplif": "Simplif. (%)", "piso": "Sem exec/piso (%)",
                "teto": "Teto (%)", "meses": "Meses"}

st.set_page_config(page_title="Necessidade orçamentária", layout="wide")
st.markdown(f"<style>{(RAIZ / 'assets' / 'style.css').read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)
# Liga o modo claro do style.css quando o tema escolhido no menu ⋮ (Light / Dark / System) é claro.
# Lê a cor do texto do app (escura = tema claro) e repete a checagem para acompanhar a troca de tema.
st.html("""<script>
if (!window.__temaNecorc) {
  window.__temaNecorc = setInterval(function () {
    var app = document.querySelector('.stApp'); if (!app) return;
    var m = getComputedStyle(app).color.match(/\\d+/g); if (!m) return;
    var lum = (0.299 * m[0] + 0.587 * m[1] + 0.114 * m[2]) / 255;
    var tema = lum < 0.5 ? 'claro' : 'escuro';
    if (document.documentElement.dataset.tema !== tema) document.documentElement.dataset.tema = tema;
  }, 300);
}
</script>""", unsafe_allow_javascript=True)
COR_NAO = "#ef4444"  # vermelho legível nos dois modos
# linhas de total: fundo verde translúcido + negrito (legível no modo claro e no escuro)
ESTILO_TOTAL = "background-color: rgba(74, 222, 128, 0.22); font-weight: 700"


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


def para_pct(fracao) -> float:
    """0,006 -> 0,6 (%). Feito em decimal para não criar resíduos de ponto flutuante."""
    return float(Decimal(repr(float(fracao))) * 100)


def de_pct(pct) -> float:
    """0,6 (%) -> 0,006: devolve exatamente o mesmo número que o Excel guarda para 0,6%."""
    return float(Decimal(repr(float(pct))) / 100)


def chave(regra: str, campo: str) -> str:
    return f"p_{regra}_{campo}"


def carregar_parametros_no_estado(p: Parametros):
    for campo, _, _ in CAMPOS:
        conv = _num if campo == "meses" else para_pct
        st.session_state[chave("geral", campo)] = conv(getattr(p.geral, campo))
        for sec in SECRETARIAS:
            st.session_state[chave(sec, campo)] = conv(getattr(p.secretaria[sec], campo))
    for sec in SECRETARIAS:
        st.session_state[chave(sec, "bln")] = False  # padrão: regra geral; o usuário liga a regra específica
    st.session_state["p_bln_execucao"] = p.bln_execucao


def parametros_do_estado() -> Parametros:
    def regra(nome):
        return Regra(**{campo: (st.session_state[chave(nome, campo)] if campo == "meses"
                                else de_pct(st.session_state[chave(nome, campo)]))
                        for campo, _, _ in CAMPOS})
    return Parametros(
        geral=regra("geral"),
        secretaria={sec: regra(sec) for sec in SECRETARIAS},
        bln_secretaria={sec: bool(st.session_state[chave(sec, "bln")]) for sec in SECRETARIAS},
        bln_execucao=bool(st.session_state["p_bln_execucao"]),
    )


def campo_numero(nome_regra: str, campo: str, rotulo: str, ajuda: str, desabilitado: bool = False):
    v = st.session_state[chave(nome_regra, campo)]
    if campo == "meses":
        st.number_input(rotulo, key=chave(nome_regra, campo), min_value=1 if isinstance(v, int) else 0.0,
                        step=1 if isinstance(v, int) else 1.0, help=ajuda, disabled=desabilitado)
    else:
        st.number_input(rotulo, key=chave(nome_regra, campo), min_value=0.0, step=0.1, format="%.1f",
                        help=ajuda, disabled=desabilitado)


# ---------------------------------------------------------------------------
# Acesso: usuários e senhas em dados/usuarios.xlsx (perfis admin e consulta)
# ---------------------------------------------------------------------------
if not ARQ_USUARIOS.exists():
    st.error(f"Arquivo de usuários não encontrado: {ARQ_USUARIOS}")
    st.stop()
try:
    usuarios = ler_usuarios(ARQ_USUARIOS)
except Exception as erro:
    st.error(f"Não foi possível ler {ARQ_USUARIOS.name}: {erro}")
    st.stop()

usuario = usuario_do_codigo(PASTA, usuarios, st.query_params.get("acesso"))
if usuario is None:
    st.title("Necessidade orçamentária: modelo preditivo")
    with st.columns([1, 1.2, 1])[1], st.form("login"):
        st.subheader("Entrar")
        nome_usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
        if st.form_submit_button("Entrar", type="primary", width="stretch"):
            u = conferir(usuarios, nome_usuario, senha)
            if u:
                st.query_params["acesso"] = gerar_codigo(PASTA, u)  # mantém o acesso ao apertar F5
                st.rerun()
            st.error("Usuário ou senha inválidos.")
    st.stop()

st.sidebar.markdown(f"**{usuario.nome}**  \n{'Administrador' if usuario.admin else 'Consulta'}")
if st.sidebar.button("Sair"):
    st.query_params.clear()
    st.session_state.clear()
    st.rerun()
if not usuario.admin:
    # perfil consulta: esconde do menu ⋮ as opções de desenvolvedor
    st.markdown("""<style>
[data-testid="stMainMenuItem-rerun"], [data-testid="stMainMenuItem-autoRerun"],
[data-testid="stMainMenuItem-clearCache"] { display: none !important; }
</style>""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Barra lateral: escolha da planilha
# ---------------------------------------------------------------------------
st.sidebar.header("Base de dados")
arquivos = sorted((a for a in PASTA.glob(PADRAO) if not a.name.startswith("~$") and a != ARQ_USUARIOS),
                  key=lambda a: a.stat().st_mtime, reverse=True)
if not arquivos:
    st.sidebar.error(f"Nenhuma planilha encontrada em {PASTA}")
    st.info(f"Salve a planilha atualizada (.xlsx) na pasta **{PASTA}** e recarregue a página. "
            "A pasta pode ser alterada no arquivo config.json.")
    st.stop()

def rotulo_arquivo(a: Path) -> str:
    return f"{a.name}  ({datetime.fromtimestamp(a.stat().st_mtime):%d/%m/%Y %H:%M})"


if usuario.admin:
    escolhido = st.sidebar.selectbox("Planilha (mais recente primeiro)", arquivos, format_func=rotulo_arquivo)
    if st.sidebar.button("Recarregar planilha"):
        carregar.clear()
else:  # consulta: sempre a planilha salva mais recentemente
    escolhido = arquivos[0]
    st.sidebar.caption(f"Planilha: {rotulo_arquivo(escolhido)}")
entrada = carregar(str(escolhido), escolhido.stat().st_mtime)
st.sidebar.caption(f"{len(entrada.linhas):,} contratos lidos da aba da view".replace(",", "."))
# Botões de tema: gravam a escolha onde o Streamlit guarda o tema (a mesma do menu ⋮) e recarregam a página.
st.sidebar.html("""
<div class="tema-botoes">
  <button type="button" data-escolha="Light">☀ Claro</button>
  <button type="button" data-escolha="Dark">☾ Escuro</button>
</div>
<script>
document.querySelectorAll('.tema-botoes button').forEach(function (b) {
  if (b.dataset.ligado) return; b.dataset.ligado = '1';
  b.addEventListener('click', function () {
    localStorage.setItem('stActiveTheme-' + window.location.pathname + '-v2', JSON.stringify(b.dataset.escolha));
    window.location.reload();
  });
});
</script>""", unsafe_allow_javascript=True)
st.sidebar.caption("Ao trocar o tema a página recarrega e os parâmetros voltam aos da planilha.")

if st.session_state.get("_arquivo") != (str(escolhido), escolhido.stat().st_mtime):
    carregar_parametros_no_estado(entrada.parametros)
    st.session_state["_arquivo"] = (str(escolhido), escolhido.stat().st_mtime)

# ---------------------------------------------------------------------------
# Parâmetros
# ---------------------------------------------------------------------------
st.title("Necessidade orçamentária: modelo preditivo")

st.subheader("Parâmetros")

# Layout: 5 colunas lado a lado (regra geral + SEMOB, SNSA, SNP, SNH). Em cada bloco, um campo por linha
# (rótulo à esquerda, valor à direita), com os 3 principais primeiro: margem, sem exec/piso, meses.
ORDEM = list(PRINCIPAIS) + ["simplif", "teto"]
AJUDA = {c: a for c, _, a in CAMPOS}


def linha_campo(nome_regra: str, campo: str, desabilitado: bool = False):
    # rótulo e valor na mesma linha: feito no style.css (classe .st-key-<chave> do campo)
    campo_numero(nome_regra, campo, ROTULO_CURTO[campo], AJUDA[campo] if nome_regra == "geral" else None,
                 desabilitado=desabilitado)


leg = st.columns([1.2, 4], gap="small")
leg[0].markdown('<div class="legenda">Regra geral</div>', unsafe_allow_html=True)
leg[1].markdown('<div class="legenda">Regras específicas por secretaria</div>', unsafe_allow_html=True)
blocos = st.columns([1.2, 1, 1, 1, 1], gap="small")

with blocos[0], st.container(border=True, key="card_geral"):
    st.markdown('<div class="sec-titulo">GERAL</div>', unsafe_allow_html=True)
    st.caption("Vale para todos, exceto secretarias com regra específica ligada.")
    for campo in ORDEM:
        linha_campo("geral", campo)
    st.toggle("Somente contratos em execução", key="p_bln_execucao",
              help="bln_execucao: afeta a linha 'Nec orçamentária final (bln_execucao)'.")
    st.button("Restaurar parâmetros da planilha", key="restaurar", help="Volta aos valores da planilha (regras específicas desligadas)",
              on_click=carregar_parametros_no_estado, args=(copy.deepcopy(entrada.parametros),))

for col, sec in zip(blocos[1:], SECRETARIAS):
    with col, st.container(border=True, key=f"card_{sec}"):
        st.markdown(f'<div class="sec-titulo">{sec}</div>', unsafe_allow_html=True)
        # ligado = regra da secretaria; desligado = regra geral
        ligada = st.toggle("Regra específica", key=chave(sec, "bln"))
        for campo in ORDEM:
            linha_campo(sec, campo, desabilitado=not ligada)

parametros = parametros_do_estado()
linhas = calcular(entrada, parametros)

# ---------------------------------------------------------------------------
# Resultados
# ---------------------------------------------------------------------------
LARGURA_TABELAS = [4, 0.7]  # tabelas um pouco mais estreitas que a página (2ª coluna fica vazia)

st.subheader("Resumo")
st.markdown('<div class="legenda">Por secretaria (valores em milhões)</div>', unsafe_allow_html=True)
df_sec = tabela_secretarias(entrada, linhas, parametros)
df_sec = df_sec.rename_axis("Item").reset_index()  # rótulos como coluna comum (mesma cor do restante)
FINAIS = df_sec.index[-2:]  # Nec. orçamentária final e a versão bln_execucao: linhas de total
estilo_sec = (df_sec.style.format(fmt_br)
              .apply(lambda r: [ESTILO_TOTAL if r.name in FINAIS else ""] * len(r), axis=1))
st.columns(LARGURA_TABELAS)[0].dataframe(estilo_sec, width="stretch", hide_index=True)

st.markdown('<div class="legenda" style="margin-top:0.8rem">Por ação orçamentária (valores em milhões)</div>',
            unsafe_allow_html=True)
df_acao, totais = tabela_acoes(entrada, linhas)
area = st.columns(LARGURA_TABELAS)[0]
area.dataframe(
    df_acao.style.format({c: fmt_br for c in ["nec orc final", "disponível LOA", "Saldo"]})
    .map(lambda v: f"color: {COR_NAO}; font-weight: 600" if v == "Não" else "", subset=["Dentro disp"]),
    width="stretch", hide_index=True, height="content")
c1, c2, c3, c4 = area.columns(4)
c1.metric("Nec. orc. final", fmt_br(totais["TOTAL nec orc final"]))
c2.metric("Disponível LOA", fmt_br(totais["TOTAL disponível LOA"]))
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
