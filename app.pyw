"""Ponto de entrada do Sistema de Impressão (abra este arquivo)."""

import logging
import sys
from pathlib import Path
from tkinter import Tk, messagebox

BASE_DIR = Path(__file__).resolve().parent
ARQUIVO_LOG = BASE_DIR / "erros.log"

# Em .pyw não existe console: sem isso, qualquer erro some em silêncio.
logging.basicConfig(
    filename=ARQUIVO_LOG,
    level=logging.WARNING,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8",
)


def main():
    try:
        from tela_ctk import Tela   # dentro do try: pega também erro de import (ex.: pywin32)

        Tela().executar()

    except Exception:
        logging.exception("Falha ao iniciar o sistema")

        raiz = Tk()
        raiz.withdraw()
        messagebox.showerror(
            "Erro",
            f"O sistema não pôde ser iniciado.\n\nDetalhes em:\n{ARQUIVO_LOG}",
        )
        raiz.destroy()
        sys.exit(1)


if __name__ == "__main__":
    main()
