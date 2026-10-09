"""Gera a versão web (uma página HTML só) com os dados da planilha mais recente da pasta dados/.

Uso:  python gerar_web.py [caminho_da_planilha.xlsx]
Saída: web/saida/necorc.html (é esse arquivo que é publicado como Artifact)
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from necorc.artifact import exportar
from necorc.leitura import ler_planilha

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


def main():
    caminho = Path(sys.argv[1]) if len(sys.argv) > 1 else planilha_mais_recente()
    dados = exportar(ler_planilha(caminho))
    pagina = (RAIZ / "web" / "pagina.html").read_text(encoding="utf-8")
    motor = (RAIZ / "web" / "motor.js").read_text(encoding="utf-8")
    pagina = pagina.replace("/*MOTOR*/", motor).replace(
        "/*DADOS*/", json.dumps(dados, ensure_ascii=False, separators=(",", ":")))
    saida = RAIZ / "web" / "saida" / "necorc.html"
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(pagina, encoding="utf-8")
    print(f"{saida}  ({saida.stat().st_size / 1e6:.2f} MB, {dados['contratos']} contratos de {caminho.name})")


if __name__ == "__main__":
    main()
