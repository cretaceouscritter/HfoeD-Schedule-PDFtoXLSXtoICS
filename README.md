# HfoeD Stundenplan Transformer

Kleine Toolchain, um Stundenpläne aus einer **PDF** erst nach **XLSX** und danach nach **ICS** zu konvertieren.

## 1) PDF → XLSX (Java)

**Voraussetzungen:** JDK + Maven (Dependencies)

**Run:**
- HfoeD Stundenplan in PDF-Form in Ordner `INPUT_PDF`legen
- Ordner `OUTPUT_XLSX` ist angelegt
- `PDF_to_XLSX.java` ausführen
- In der Konsole die Anzahl  der Gruppen eingeben:

**Output:**\
Stundenplaene.xlsx in Ordner `OUTPUT_XLSX`

## 2) XLSX → ICS (Python)

**Voraussetzungen:**
* Das erzeugte Stundenplaene.xlsx liegt in Ordner `OUTPUT_XLSX`
* Ordner `OUTPUT_ICS` ist angelegt
* Libraries `pandas`, `ipytz` und `icalendar` müssen installiert sein

**Run:**
- `XLSX_to_ICS.py` ausführen
- In der Konsole eingeben: Datum des ersten Tages der Stundenpläne (immer Montag)

**Output**
GruppeX_WocheX.ics Dateien im Ordner `OUTPUT_ICS`
