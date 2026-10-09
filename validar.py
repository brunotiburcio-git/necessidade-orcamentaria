"""Confere se o cálculo em Python bate com os valores calculados e salvos pelo Excel.

Uso:  python validar.py caminho/da/planilha.xlsx

Compara, com os parâmetros salvos na própria planilha:
- todas as colunas calculadas da view (M e AN..BV), linha a linha;
- as tabelas resumo K12:P18, K20:P20 e K26:P47.
Importante: a planilha precisa ter sido salva DEPOIS de o Excel recalcular.
"""
from __future__ import annotations

import sys

from necorc.excel import eh_numero
from necorc.leitura import ler_planilha
from necorc.modelo import COLUNAS_CALCULADAS, calcular, resumo_acoes, resumo_secretarias

TOL = 1e-6  # tolerância absoluta (R$) para considerar "igual" em valores numéricos


def _compara(a, b):
    """Retorna (igual_exato, diferenca_absoluta)."""
    if eh_numero(a) and eh_numero(b):
        return a == b, abs(a - b)
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b, 0 if a == b else float("inf")
    sa = "" if a is None else a
    sb = "" if b is None else b
    if isinstance(sa, str) and isinstance(sb, str) and sb.startswith("#"):
        ok = sa.upper().startswith("#")
        return ok, 0 if ok else float("inf")
    return sa == sb, 0 if sa == sb else float("inf")


def validar(caminho: str) -> bool:
    ent = ler_planilha(caminho)
    linhas = calcular(ent)
    print(f"Arquivo: {ent.arquivo}")
    print(f"Linhas da view: {len(linhas)}\n")

    tudo_ok = True
    print(f"{'coluna':45s} {'exatas':>8s} {'dif>tol':>8s} {'maior dif':>12s}")
    for col in COLUNAS_CALCULADAS:
        exatas, ruins, maior, exemplo = 0, 0, 0.0, None
        for i, (py, xl) in enumerate(zip(linhas, ent.excel_view)):
            eq, dif = _compara(py[col], xl.get(col))
            exatas += eq
            maior = max(maior, dif)
            if dif > TOL:
                ruins += 1
                if exemplo is None:
                    exemplo = (i, py["cod_tci"], py[col], xl.get(col))
        tudo_ok &= ruins == 0
        print(f"{col:45s} {exatas:8d} {ruins:8d} {maior:12.3g}")
        if exemplo:
            print(f"   ex.: linha {exemplo[0]} cod_tci={exemplo[1]} python={exemplo[2]!r} excel={exemplo[3]!r}")

    print("\nResumo por secretaria (milhões)")
    p = ent.excel_param
    for item in resumo_secretarias(ent, linhas, ent.parametros):
        lin = item["linha"]
        for c, v in zip("LMNOP", item["valores"] + [item["total"]]):
            xl = p[f"{c}{lin}"]
            dif = abs(v - xl)
            tudo_ok &= dif <= TOL
            flag = "OK" if dif <= TOL else "DIVERGE"
            print(f"  {c}{lin:<3d} {item['rotulo'][:38]:38s} python={v:20.10f} excel={xl:20.10f} dif={dif:.2e} {flag}")

    print("\nResumo por ação (milhões)")
    ra = resumo_acoes(ent, linhas)
    checks = []
    for it in ra["itens"]:
        lin = it["linha"]
        checks += [(f"M{lin}", it["nec_orc_final"]), (f"N{lin}", it["disponivel_loa"]),
                   (f"O{lin}", it["dentro_disp"]), (f"P{lin}", it["saldo"])]
    checks += [("M46", ra["total_nec"]), ("N46", ra["total_disp"]),
               ("P46", ra["faltando"]), ("P47", ra["sobrando"])]
    for cel, v in checks:
        xl = p[cel]
        if isinstance(v, str):
            ok = v == xl
            print(f"  {cel:5s} python={v!r:>22s} excel={xl!r:>22s} {'OK' if ok else 'DIVERGE'}")
        else:
            ok = abs(v - xl) <= TOL
            print(f"  {cel:5s} python={v:22.10f} excel={xl:22.10f} dif={abs(v - xl):.2e} {'OK' if ok else 'DIVERGE'}")
        tudo_ok &= ok

    print("\nRESULTADO:", "tudo confere com o Excel" if tudo_ok else "há divergências (ver acima)")
    return tudo_ok


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(0 if validar(sys.argv[1]) else 1)
