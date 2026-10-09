"""Controle de acesso simples: usuários e senhas numa planilha Excel da pasta de dados.

Planilha (primeira aba), com cabeçalho na linha 1:
    usuario | nome | perfil | senha
- perfil: "admin" (acesso completo) ou "consulta" (parâmetros e resultados)
- linhas com usuario vazio são ignoradas; usuario não diferencia maiúsculas/minúsculas.

A sessão continua aberta ao apertar F5 por meio de um código no endereço da página
(?acesso=...), assinado com uma chave guardada na pasta de dados.
"""
from __future__ import annotations

import hashlib
import hmac
import secrets
from dataclasses import dataclass
from pathlib import Path

import openpyxl

PERFIS = ("admin", "consulta")
COLUNAS = ("usuario", "nome", "perfil", "senha")


@dataclass
class Usuario:
    usuario: str
    nome: str
    perfil: str
    senha: str

    @property
    def admin(self) -> bool:
        return self.perfil == "admin"


def _texto(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)  # senha 1234 digitada como número no Excel
    return str(v).strip()


def ler_usuarios(caminho: Path) -> dict[str, Usuario]:
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    try:
        linhas = list(wb.worksheets[0].iter_rows(values_only=True))
    finally:
        wb.close()
    if not linhas:
        raise ValueError("planilha de usuários vazia")
    cab = [_texto(c).lower() for c in linhas[0]]
    faltando = [c for c in COLUNAS if c not in cab]
    if faltando:
        raise ValueError(f"faltam as colunas {', '.join(faltando)} na linha 1")
    pos = {c: cab.index(c) for c in COLUNAS}
    usuarios = {}
    for linha in linhas[1:]:
        val = {c: _texto(linha[i]) if i < len(linha) else "" for c, i in pos.items()}
        if not val["usuario"]:
            continue
        perfil = val["perfil"].lower()
        usuarios[val["usuario"].lower()] = Usuario(
            usuario=val["usuario"].lower(), nome=val["nome"] or val["usuario"],
            perfil=perfil if perfil in PERFIS else "consulta", senha=val["senha"])
    return usuarios


def conferir(usuarios: dict[str, Usuario], usuario: str, senha: str) -> Usuario | None:
    u = usuarios.get(usuario.strip().lower())
    if u and u.senha and hmac.compare_digest(u.senha, senha.strip()):
        return u
    return None


def _chave(pasta: Path) -> bytes:
    arq = pasta / ".chave_sessao"
    if not arq.exists():
        arq.write_text(secrets.token_hex(32), encoding="utf-8")
    return arq.read_text(encoding="utf-8").strip().encode()


def gerar_codigo(pasta: Path, u: Usuario) -> str:
    """Código da sessão: muda se a senha ou o perfil mudarem na planilha."""
    assinatura = hmac.new(_chave(pasta), f"{u.usuario}|{u.senha}|{u.perfil}".encode(), hashlib.sha256)
    return f"{u.usuario}.{assinatura.hexdigest()[:32]}"


def usuario_do_codigo(pasta: Path, usuarios: dict[str, Usuario], codigo: str | None) -> Usuario | None:
    if not codigo or "." not in codigo:
        return None
    u = usuarios.get(codigo.rsplit(".", 1)[0])
    if u and hmac.compare_digest(gerar_codigo(pasta, u), codigo):
        return u
    return None
