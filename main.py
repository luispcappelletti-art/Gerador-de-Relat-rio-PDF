import customtkinter as ctk
import subprocess
import sys
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent


class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Central de Funções")
        self.geometry("520x300")
        self.minsize(460, 260)

        self._processes = []

        container = ctk.CTkFrame(self)
        container.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            container,
            text="Escolha qual função do app deseja abrir",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(pady=(10, 8))

        ctk.CTkLabel(
            container,
            text="Cada opção abre um módulo separado para facilitar a manutenção.",
            text_color=("#4A5568", "#A0AEC0"),
        ).pack(pady=(0, 18))

        ctk.CTkButton(
            container,
            text="Gerar Relatórios",
            width=220,
            height=40,
            command=lambda: self._open_module("gerar_relatorio.py"),
        ).pack(pady=6)

        ctk.CTkButton(
            container,
            text="Gerar Certificados (exemplo)",
            width=220,
            height=40,
            command=lambda: self._open_module("gerar_certificado.py"),
        ).pack(pady=6)

        self.status_label = ctk.CTkLabel(container, text="Pronto.", anchor="w")
        self.status_label.pack(fill="x", pady=(14, 0), padx=8)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _open_module(self, module_filename: str):
        module_path = BASE_DIR / module_filename
        if not module_path.exists():
            self.status_label.configure(text=f"Arquivo não encontrado: {module_filename}")
            return

        try:
            proc = subprocess.Popen([sys.executable, str(module_path)])
            self._processes.append(proc)
            self.status_label.configure(text=f"Módulo aberto: {module_filename}")
        except Exception as exc:
            self.status_label.configure(text=f"Falha ao abrir {module_filename}: {exc}")

    def _on_close(self):
        for proc in self._processes:
            if proc.poll() is None:
                proc.terminate()
        self.destroy()


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    app = MainApp()
    app.mainloop()
