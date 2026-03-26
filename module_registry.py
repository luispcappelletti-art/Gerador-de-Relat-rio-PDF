from pathlib import Path
from typing import List, Tuple


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_MODULE = "gerar_relatorio.py"
MODULE_LABEL_OVERRIDES = {
    "gerar_relatorio.py": "Gerar Relatórios",
    "gerar_certificado.py": "Gerar Certificados (exemplo)",
}


def discover_modules() -> List[Tuple[str, str]]:
    modules: List[Tuple[str, str]] = []
    for module_path in BASE_DIR.glob("gerar_*.py"):
        if module_path.name == "main.py":
            continue
        label = MODULE_LABEL_OVERRIDES.get(module_path.name, _humanize_filename(module_path.stem))
        modules.append((module_path.name, label))
    modules.sort(key=lambda item: item[1].lower())
    return modules


def available_module_names() -> set[str]:
    return {name for name, _ in discover_modules()}


def _humanize_filename(stem: str) -> str:
    return stem.replace("_", " ").strip().title()
