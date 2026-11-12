#!/bin/sh
# Maakt beide PDF's en de presentatie opnieuw. Draai eerst scripts/bouw_dataset.py als de data verandert.
cd "$(dirname "$0")"
typst compile --font-path fonts --input excel=nl typst/practicum.typ practicum_leerling_nl-excel.pdf
typst compile --font-path fonts --input excel=en typst/practicum.typ practicum_leerling_en-excel.pdf
typst compile --font-path fonts --input docent=true typst/practicum.typ practicum_begeleiders.pdf
echo "Klaar: practicum_leerling_nl-excel.pdf, practicum_leerling_en-excel.pdf en practicum_begeleiders.pdf"

# Presentatie: gebouwd op een bestaande UU-presentatie als sjabloon (huisstijl, layouts).
SJABLOON="${SJABLOON:-$HOME/Library/Mobile Documents/com~apple~CloudDocs/UU/Meeloopdagen/mld/2026_03_23-twinfo/presOD.pptx}"
PYTHON="${PYTHON:-$HOME/.venvs/practicum-vertraagd/bin/python}"
if [ -f "$SJABLOON" ] && [ -x "$PYTHON" ]; then
  "$PYTHON" scripts/bouw_presentatie.py "$SJABLOON"
else
  echo "Presentatie overgeslagen: sjabloon ($SJABLOON) of python-pptx-omgeving ($PYTHON) niet gevonden."
fi
