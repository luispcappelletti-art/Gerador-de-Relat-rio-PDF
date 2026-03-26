"""Compatibilidade de execução.

Use `python main.py` como ponto de entrada principal.
"""

from main import App  # reexport útil para integrações antigas
import customtkinter as ctk


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    app = App()
    app.mainloop()
