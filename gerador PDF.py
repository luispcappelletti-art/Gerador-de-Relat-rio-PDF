"""Compatibilidade: ponto de entrada antigo.

Use `main.py` para escolher os módulos do app.
"""

from gerar_relatorio import App
import customtkinter as ctk


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    app = App()
    app.mainloop()
