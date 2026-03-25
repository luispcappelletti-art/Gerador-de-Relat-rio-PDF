import importlib.util
import pathlib
import sys
import types
import unittest


ctk_stub = types.SimpleNamespace(CTk=type("CTk", (), {}), CTkLabel=object, CTkButton=object, CTkTextbox=object, CTkFrame=object, set_appearance_mode=lambda *_: None)
sys.modules.setdefault("customtkinter", ctk_stub)
pypdf_stub = types.SimpleNamespace(PdfReader=object, PdfWriter=object)
sys.modules.setdefault("pypdf", pypdf_stub)
fitz_stub = types.SimpleNamespace(open=lambda *_args, **_kwargs: None, Matrix=lambda *_args, **_kwargs: None)
sys.modules.setdefault("fitz", fitz_stub)

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

    def test_parse_header_info_text_aceita_aliases(self):
        info = gerador_pdf._parse_header_info_text(
            "Técnico: Ana\nHorário de término: 12:00\nTempo atendimento: 30 min"
        )
        self.assertEqual(info["tecnico"], "Ana")
        self.assertEqual(info["fim"], "12:00")
        self.assertEqual(info["tempo_atendimento"], "30 min")

    def test_compose_full_text_prioriza_info_editada(self):
        content = gerador_pdf._compose_full_text_with_sections(
            "Cliente: Original",
            {
                "info": {"cliente": "Atualizado"},
                "descricao": "novo escopo",
            },
        )
        self.assertIn("Cliente: Atualizado", content)
        self.assertIn("1 – ESCOPO DO ATENDIMENTO", content)


if __name__ == "__main__":
    unittest.main()
