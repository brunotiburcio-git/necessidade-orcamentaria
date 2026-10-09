"""Funções que reproduzem o comportamento das funções do Excel usadas na planilha.

Detalhes que importam para o resultado bater com o Excel:
- Comparações de texto no Excel ("=" e "<>") ignoram maiúsculas/minúsculas
  ("sim" = "Sim", "Em execução" = "Em Execução").
- Célula vazia vale 0 em contas e "" em comparações de texto.
- ARRED/ROUND arredonda "meio para cima" (0,00005 -> 0,0001), diferente do round() do Python.
- PROCV/VLOOKUP com correspondência exata não converte número em texto (8865 <> "8865"),
  mas SOMASES/SUMIFS com critério numérico aceita células com o número em texto.
"""
from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Iterable

NA = "#N/A"


class ErroExcel(Exception):
    """Equivalente a um erro do Excel (#N/D, #VALOR!...) dentro de SEERRO/IFERROR."""


def eh_numero(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def n(v: Any) -> float:
    """Valor usado em contas: célula vazia vale 0."""
    if v is None:
        return 0
    if isinstance(v, bool):
        return int(v)
    if eh_numero(v):
        return v
    raise ErroExcel(f"#VALOR! (texto em conta: {v!r})")


def _texto(v: Any) -> str:
    if v is None:
        return ""
    return str(v).casefold()


def igual(a: Any, b: Any) -> bool:
    """Operador "=" do Excel."""
    if a is None and b is None:
        return True
    if isinstance(a, str) or isinstance(b, str):
        if a is None or b is None:
            return _texto(a) == _texto(b)
        if isinstance(a, str) and isinstance(b, str):
            return a.casefold() == b.casefold()
        return False  # texto x número nunca é igual
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b and type(a) is type(b)
    return n(a) == n(b)


def diferente(a: Any, b: Any) -> bool:
    """Operador "<>" do Excel."""
    return not igual(a, b)


def arred(x: float, casas: int) -> float:
    """ARRED/ROUND do Excel (meio para longe do zero, sobre a representação decimal)."""
    d = Decimal(repr(float(x)))
    return float(d.quantize(Decimal(1).scaleb(-casas), rounding=ROUND_HALF_UP))


def maximo(valores: Iterable[Any]) -> float:
    """MÁXIMO/MAX sobre um intervalo: ignora vazios e textos; sem números retorna 0."""
    nums = [v for v in valores if eh_numero(v)]
    return max(nums) if nums else 0


def corresp_exata(valor: Any, intervalo: list[Any]) -> int:
    """CORRESP/MATCH(valor; intervalo; 0). Retorna a posição (1..n) ou levanta ErroExcel."""
    for i, v in enumerate(intervalo, start=1):
        if v is None:
            continue
        if eh_numero(valor) and eh_numero(v) and valor == v:
            return i
        if isinstance(valor, str) and isinstance(v, str) and valor.casefold() == v.casefold():
            return i
    raise ErroExcel(NA)


class Procv:
    """PROCV/VLOOKUP(valor; tabela; coluna; FALSO) – devolve o PRIMEIRO registro encontrado."""

    def __init__(self, linhas: list[tuple[Any, Any]]):
        # linhas: (chave, valor_retorno) na ordem da planilha
        self._idx: dict[tuple[str, Any], Any] = {}
        for chave, valor in linhas:
            k = self._chave(chave)
            if k is not None and k not in self._idx:
                self._idx[k] = valor

    @staticmethod
    def _chave(v: Any):
        if v is None:
            return None
        if isinstance(v, bool):
            return ("b", v)
        if eh_numero(v):
            return ("n", float(v))
        return ("t", str(v).casefold())

    def __call__(self, valor: Any) -> Any:
        k = self._chave(valor)
        if k is None or k not in self._idx:
            raise ErroExcel(NA)
        v = self._idx[k]
        return 0 if v is None else v  # célula vazia retornada pelo PROCV vira 0


def _como_numero(v: Any):
    if eh_numero(v):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.strip().replace(",", "."))
        except ValueError:
            return None
    return None


def criterio_ok(celula: Any, criterio: Any) -> bool:
    """Teste de igualdade usado por SOMASES/SUMIFS (critério sem operadores/curingas)."""
    if isinstance(criterio, bool):
        return isinstance(celula, bool) and celula == criterio
    crit_num = _como_numero(criterio)
    if crit_num is not None:
        cel_num = _como_numero(celula) if not isinstance(celula, bool) else None
        return cel_num is not None and cel_num == crit_num
    if isinstance(celula, str):
        return celula.casefold() == str(criterio).casefold()
    return celula is None and criterio == ""


def somases(soma: list[Any], *pares: tuple[list[Any], Any]) -> float:
    """SOMASES/SUMIFS: soma, na ordem das linhas, apenas valores numéricos."""
    total = 0
    for i, v in enumerate(soma):
        if not eh_numero(v):
            continue
        if all(criterio_ok(intervalo[i], crit) for intervalo, crit in pares):
            total += v
    return total


def _quase_igual(a: float, b: float) -> bool:
    return a == b or abs(a - b) < abs(a) * 2.0 ** -48


def sub(a: float, b: float) -> float:
    """Subtração com o ajuste do Excel: se a e b têm o mesmo sinal e são iguais até a
    ~15ª casa significativa, o resultado vira 0 (ex.: 313865,5800000001 - 313865,58 = 0)."""
    if ((a > 0 and b > 0) or (a < 0 and b < 0)) and _quase_igual(a, b):
        return 0
    return a - b


def soma(a: float, b: float) -> float:
    """Adição com o mesmo ajuste do Excel (operandos de sinais opostos)."""
    return sub(a, -b)
