import customtkinter as ctk
import subprocess
import sys
import json
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
LAUNCHER_STATE_FILE = BASE_DIR / "launcher_state.json"
DEFAULT_MODULE = "gerar_relatorio.py"
MODULES = {
    "gerar_relatorio.py": "Gerar Relatórios",
    "gerar_certificado.py": "Gerar Certificados (exemplo)",
}


def load_last_module() -> str:
    if not LAUNCHER_STATE_FILE.exists():
        return DEFAULT_MODULE
    try:
        data = json.loads(LAUNCHER_STATE_FILE.read_text(encoding="utf-8"))
        module = data.get("last_module", DEFAULT_MODULE)
        return module if module in MODULES else DEFAULT_MODULE
    except Exception:
        return DEFAULT_MODULE


def save_last_module(module_filename: str):
    if module_filename not in MODULES:
        return
    LAUNCHER_STATE_FILE.write_text(
        json.dumps({"last_module": module_filename}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Central de Funções")
        self.geometry("520x300")
        self.minsize(460, 260)

        self._processes = []
        self._autostarted = False

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
            command=lambda: self._open_module("gerar_relatorio.py", close_launcher=True),
        ).pack(pady=6)

        ctk.CTkButton(
            container,
            text="Gerar Certificados (exemplo)",
            width=220,
            height=40,
            command=lambda: self._open_module("gerar_certificado.py", close_launcher=True),
        ).pack(pady=6)

        self.status_label = ctk.CTkLabel(container, text="Pronto.", anchor="w")
        self.status_label.pack(fill="x", pady=(14, 0), padx=8)

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(150, self._autostart_last_module)

    def _autostart_last_module(self):
        if self._autostarted:
            return
        self._autostarted = True
        last_module = load_last_module()
        self._open_module(last_module, close_launcher=True)

    def _open_module(self, module_filename: str, close_launcher: bool = False):
        module_path = BASE_DIR / module_filename
        if not module_path.exists():
            self.status_label.configure(text=f"Arquivo não encontrado: {module_filename}")
            return

        try:
            self._terminate_open_modules()
            proc = subprocess.Popen([sys.executable, str(module_path)])
            self._processes.append(proc)
            save_last_module(module_filename)
            self.status_label.configure(text=f"Módulo aberto: {module_filename}")
            if close_launcher:
                self.after(120, self.destroy)
        except Exception as exc:
            self.status_label.configure(text=f"Falha ao abrir {module_filename}: {exc}")

    def _terminate_open_modules(self):
        for proc in self._processes:
            if proc.poll() is None:
                proc.terminate()
        self._processes = []

    def _on_close(self):
        self._terminate_open_modules()
        self.destroy()


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    app = MainApp()
    app.mainloop()
