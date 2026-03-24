# Gerador de Relatório PDF

Aplicação em Python para transformar um texto de entrada em um relatório PDF com layout profissional.

## Principais melhorias

- **Visual profissional** com capa, título, subtítulo, tabela de seções, cabeçalho e rodapé.
- **Compatibilidade de títulos** com normalização robusta (maiúsculas/minúsculas, acentos e espaços extras).
- **Mapeamento de variações** para os títulos padrão de relatório.

## Instalação

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Uso

```bash
python app.py --input exemplo.txt --output relatorio.pdf
```

Formato de entrada sugerido:

```text
RESUMO EXECUTIVO: Conteúdo do resumo.
Objetivo Geral: Conteúdo do objetivo.
metodologia : Conteúdo da metodologia.
Análise dos Resultados: Conteúdo da análise.
Conclusão e Recomendações: Conteúdo final.
```

## Testes

```bash
python -m unittest discover -s tests -p "test_*.py"
```
