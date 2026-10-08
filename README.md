# Automação de imagens — speciesLink

Web scraper em Python para catalogar padrões de barcode das imagens públicas do CRIA/speciesLink.

O projeto consulta somente endpoints públicos do catálogo, não baixa imagens e registra progresso em SQLite para permitir retomada segura.

## Uso

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:PYTHONPATH = "src"
python -m cria_barcode_scraper.cli
```

O arquivo Excel é gerado em `data/padroes_barcode_cria.xlsx`.
