from __future__ import annotations

import traceback
from tkinter import messagebox

from app.ui.main_window import MainWindow


def main() -> None:
    try:
        app = MainWindow()
        app.mainloop()
    except Exception as exc:  # pragma: no cover - proteção final para o usuário
        traceback.print_exc()
        messagebox.showerror(
            "Erro inesperado",
            "Ocorreu um erro inesperado ao iniciar o sistema.\n\n"
            f"Detalhes: {exc}",
        )


if __name__ == "__main__":
    main()
