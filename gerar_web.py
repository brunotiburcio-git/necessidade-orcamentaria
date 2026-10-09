"""Gera a versão web (uma página HTML só) com os dados da planilha mais recente da pasta dados/.

Uso:  python gerar_web.py [caminho_da_planilha.xlsx] [caminho_usuarios.xlsx]
Saída: web/saida/necorc.html (Artifact do Claude) e docs/index.html (GitHub Pages)

Os dados vão cifrados (AES-GCM) dentro da página. Cada usuário de usuarios.xlsx recebe uma cópia da
chave dos dados cifrada com a própria senha (PBKDF2), então sem um usuário e senha válidos a página
não consegue ler os números.
"""
from __future__ import annotations

import base64
import gzip
import hashlib
import json
import os
import sys
from pathlib import Path

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

from necorc.artifact import exportar
from necorc.leitura import ler_planilha
from necorc.usuarios import ler_usuarios

ITERACOES = 250_000

RAIZ = Path(__file__).resolve().parent
CONFIG = json.loads((RAIZ / "config.json").read_text(encoding="utf-8"))


def planilha_mais_recente() -> Path:
    pasta = RAIZ / CONFIG.get("pasta_planilhas", "dados")
    usuarios = CONFIG.get("arquivo_usuarios", "usuarios.xlsx")
    arquivos = [a for a in pasta.glob(CONFIG.get("padrao_arquivo", "*.xlsx"))
                if not a.name.startswith("~$") and a.name != usuarios]
    if not arquivos:
        sys.exit(f"Nenhuma planilha em {pasta}")
    return max(arquivos, key=lambda a: a.stat().st_mtime)


def _b64(b: bytes) -> str:
    return base64.b64encode(b).decode()


def _escape_js(c: str) -> str:
    """Caractere fora do ASCII como escape \\uXXXX do JavaScript (par substituto acima de U+FFFF)."""
    b = c.encode("utf-16-be")
    return "".join(f"\\u{int.from_bytes(b[i:i + 2], 'big'):04x}" for i in range(0, len(b), 2))


def proteger(dados: dict, usuarios) -> dict:
    chave = AESGCM.generate_key(bit_length=256)
    iv = os.urandom(12)
    texto = gzip.compress(json.dumps(dados, ensure_ascii=False, separators=(",", ":")).encode(), mtime=0)
    protegido = {"iter": ITERACOES, "dados": {"iv": _b64(iv), "ct": _b64(AESGCM(chave).encrypt(iv, texto, None))},
                 "chaves": {}}
    for u in usuarios.values():
        if not u.senha:
            continue
        sal, iv_u = os.urandom(16), os.urandom(12)
        kek = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=sal, iterations=ITERACOES).derive(u.senha.encode())
        conteudo = json.dumps({"k": _b64(chave), "nome": u.nome}, ensure_ascii=False).encode()
        protegido["chaves"][hashlib.sha256(u.usuario.encode()).hexdigest()] = {
            "salt": _b64(sal), "iv": _b64(iv_u), "ct": _b64(AESGCM(kek).encrypt(iv_u, conteudo, None))}
    if not protegido["chaves"]:
        sys.exit("Nenhum usuário com senha em usuarios.xlsx")
    return protegido


def main():
    caminho = Path(sys.argv[1]) if len(sys.argv) > 1 else planilha_mais_recente()
    arq_usuarios = (Path(sys.argv[2]) if len(sys.argv) > 2
                    else RAIZ / CONFIG.get("pasta_planilhas", "dados") / CONFIG.get("arquivo_usuarios", "usuarios.xlsx"))
    usuarios = ler_usuarios(arq_usuarios)
    dados = exportar(ler_planilha(caminho))
    pagina = (RAIZ / "web" / "pagina.html").read_text(encoding="utf-8")
    motor = (RAIZ / "web" / "motor.js").read_text(encoding="utf-8")
    # SheetJS 0.18.5 (versão mini: só o que a exportação usa); caracteres especiais viram \uXXXX
    xlsx = (RAIZ / "web" / "vendor" / "xlsx.mini.min.js").read_text(encoding="utf-8")
    xlsx = "".join(c if ord(c) < 128 else _escape_js(c) for c in xlsx)
    pagina = pagina.replace("/*MOTOR*/", motor, 1).replace(
        "/*DADOS*/", json.dumps(proteger(dados, usuarios), separators=(",", ":")), 1)
    pagina = pagina.replace("/*XLSX*/", xlsx, 1)  # por último: o código da biblioteca não passa pelas outras trocas
    saida = RAIZ / "web" / "saida" / "necorc.html"   # versão para o Artifact do Claude
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(pagina, encoding="utf-8")
    # versão para o GitHub Pages (pasta docs/): mesma página com o cabeçalho HTML completo
    site = RAIZ / "docs" / "index.html"
    site.parent.mkdir(parents=True, exist_ok=True)
    site.write_text('<!doctype html>\n<html lang="pt-BR">\n<head>\n<meta charset="utf-8">\n'
                    '<meta name="viewport" content="width=device-width,initial-scale=1">\n'
                    '<meta name="robots" content="noindex">\n</head>\n<body>\n'
                    + pagina + "\n</body>\n</html>\n", encoding="utf-8")
    print(f"{saida}  ({saida.stat().st_size / 1e6:.2f} MB, {dados['contratos']} contratos de {caminho.name}, "
          f"{len(usuarios)} usuários)")


if __name__ == "__main__":
    main()
