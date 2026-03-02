# HfoeD Stundenplan Transformer

Kleine Toolchain, um Stundenpläne aus einer **PDF** erst nach **XLSX** und danach nach **ICS** zu konvertieren.

## 1) PDF → XLSX (Java)

**Voraussetzungen:** JDK + Maven (Dependencies)

**Run:**
- `PDF_to_XLSX.java` ausführen
- In der Konsole eingeben:
  - Anzahl Gruppen
  - Anzahl Wochen
  - Pfad zur PDF

**Output:**\
Stundenplaene.xlsx

## 2) XLSX → ICS (Python)

**Voraussetzungen:**
* Verzeichnis `Output` angelegt sein
* Libraries `pandas` und `ics` müssen installiert sein

**Run:**
- `XLSX_to_ICS.py` ausführen
- In der Konsole eingeben:
  - Anzahl Gruppen
  - Anzahl Wochen
  - Datum des ersten Tages (Montag!)

**Output**
GruppeX_WocheX.ics Datein im Ordner `Output`