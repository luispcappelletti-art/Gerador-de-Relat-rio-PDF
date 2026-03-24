import unittest

from app import parse_sections, resolve_standard_title


class NormalizationTests(unittest.TestCase):
    def test_title_variations_map_to_standard(self):
        self.assertEqual(resolve_standard_title("   RESUMO   EXECUTIVO "), "Resumo Executivo")
        self.assertEqual(resolve_standard_title("metodologia"), "Metodologia")
        self.assertEqual(resolve_standard_title("CONCLUSAO E RECOMENDACOES"), "Conclusão e Recomendações")

    def test_parse_orders_standard_sections(self):
        raw = """
metodologia: detalhe da metodologia
objetivo geral: detalhe do objetivo
resumo executivo: detalhe do resumo
"""
        sections = parse_sections(raw)
        self.assertEqual([s.title for s in sections[:3]], ["Resumo Executivo", "Objetivo Geral", "Metodologia"])


if __name__ == "__main__":
    unittest.main()
