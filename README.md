# Butikkpotensial

Python-verktøy for å sammenligne varegrupper mellom `REMA 1000 Bakklandet` og en valgt sammenligningsbutikk.

## Hva programmet gjør

- Leser flere sammenligningsfiler, én per varegruppe
- Leser én bruttofil som grunnlag for `Brutto %` og `Brutto Kr`
- Matcher varer på tallkoden bakerst i varenavnet
- Filtrerer bort varer med `Brutto % <= 18`
- Normaliserer omsetningen for en rettferdig sammenligning
- Viser toppvarer per varegruppe basert på størst potensial i Bakklandet, med omsetning i butikk B som sekundær prioritering
- Eksporterer resultatet til en ny Excel-fil

## Kjør GUI

```bash
python3 run_butikk_potensial.py
```

## Kjør fra kommandolinjen

```bash
python3 run_butikk_potensial.py \
  --brutto "/sti/til/Varer med brutto.xlsx" \
  --compare "/sti/til/Sammenlign drikke.xlsx" "/sti/til/Sammenlign meieri.xlsx" \
  --top-n 10 \
  --minimum-brutto 18
```

Valgfri totalfil:

```bash
python3 run_butikk_potensial.py \
  --brutto "/sti/til/Varer med brutto.xlsx" \
  --compare "/sti/til/Sammenlign drikke.xlsx" "/sti/til/Sammenlign meieri.xlsx" \
  --total "/sti/til/Sammenlign total.xlsx" \
  --normalization-mode total-file
```

## Krav

- Python 3.14+
- `pandas`
- `openpyxl`

## Viktig å vite

Hvis bruttofila bare inneholder et lite utvalg varer, vil programmet gi advarsler om lav treffrate mot sammenligningsfilene. Da bør bruttofila eksporteres på nytt med flere varer for å gi et bedre beslutningsgrunnlag.

## Windows-distribusjon

Dette prosjektet kan bygges til en Windows-exe ved hjelp av `PyInstaller`.

En enkel måte å gjøre det på er å bruke GitHub Actions og fila `.github/workflows/windows-build.yml` som nå er lagt til i repoet.

Når workflowen kjører, vil den bygge `gui.py` til en enkelt Windows-eksekverbar fil og publisere den som et artifact.

Hvis du selv vil bygge lokalt på Windows, kan du kjøre:

```bash
python -m pip install --upgrade pip
pip install pyinstaller
pip install -r requirements.txt
pyinstaller --noconfirm --onefile --windowed gui.py
```

Da får du en `dist/gui.exe` som kan kjøres på Windows uten at brukeren trenger Python.
