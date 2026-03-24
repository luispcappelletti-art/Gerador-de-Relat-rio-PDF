import importlib.util
import pathlib
import sys
import types
import unittest


ctk_stub = types.SimpleNamespace(CTk=type("CTk", (), {}), CTkLabel=object, CTkButton=object, CTkTextbox=object, CTkFrame=object, set_appearance_mode=lambda *_: None)
sys.modules.setdefault("customtkinter", ctk_stub)
pypdf_stub = types.SimpleNamespace(PdfReader=object, PdfWriter=object)
sys.modules.setdefault("pypdf", pypdf_stub)

MODULE_PATH = pathlib.Path(__file__).resolve().parents[1] / "gerador PDF.py"
spec = importlib.util.spec_from_file_location("gerador_pdf", MODULE_PATH)
gerador_pdf = importlib.util.module_from_spec(spec)
assert spec.loader is not None
spec.loader.exec_module(gerador_pdf)


class ListaEInfoTests(unittest.TestCase):
    def test_texto_com_topicos_explicitos_nao_redivide_por_ponto(self):
        from reportlab.lib.styles import getSampleStyleSheet

        styles = {"Body": getSampleStyleSheet()["BodyText"]}
        elementos = gerador_pdf.processar_lista(
            "- Primeiro item.\ncontinuação do primeiro.\n- Segundo item.",
            styles,
        )

        lista = elementos[0]
        self.assertEqual(len(lista._flowables), 2)

    def test_ponto_isolado_no_cabecalho_e_considerado_vazio(self):
        self.assertFalse(gerador_pdf._valor_info_preenchido("."))
        self.assertFalse(gerador_pdf._valor_info_preenchido("   .   "))
        self.assertTrue(gerador_pdf._valor_info_preenchido("ok"))


if __name__ == "__main__":
    unittest.main()
