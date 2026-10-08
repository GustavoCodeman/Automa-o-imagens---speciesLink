# Automação de imagens — speciesLink

Web scraper em Python para catalogar padrões de barcode das imagens públicas do CRIA/speciesLink.

## Executar no Windows

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
py scraper.py
```

O progresso é salvo em `data/progress.sqlite` e a planilha é gerada em `data/padroes_barcode_cria.xlsx`. Para retomar uma coleta interrompida, repita o mesmo comando.

> A primeira versão ainda está em evolução: antes de uma coleta completa, valide a execução com uma coleção pequena.
