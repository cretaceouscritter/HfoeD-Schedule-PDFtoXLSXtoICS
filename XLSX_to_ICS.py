import pandas as pd
import warnings
import datetime
from ics import Calendar, Event
import pytz


# Suppress specific warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
# Pfad zur Excel-Datei
filepath = "C:\\Users\\Aiden.Hintersberger\\OneDrive\\Desktop\\Stundenplaene_HfoeD2_1.xlsx"

# Variablen für Gruppen und Wochen
groups = 3
weeks = 9
start_date = datetime.date(2026, 3, 2)  #Startdatum

# Dictionary zur Speicherung der Daten
data = {}

# Daten aus den Excel-Tabellenblättern einlesen
for group in range(1, groups + 1):
    for week in range(1, weeks + 1):
        sheet_name = f"Gruppe{group}_Woche{week}"
        # Einlesen des Tabellenblatts in ein DataFrame
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
        # Speicherung des DataFrames im Dictionary
        data[sheet_name] = df

# Beispiel für die Startzeiten jeder Stunde
class_times = [
    "08:00-08:45", "08:45-09:30", "10:00-10:45", "10:45-11:30",
    "12:00-12:45", "12:45-13:30", "13:30-14:15", "14:30-15:15",
    "15:15-16:00", "16:15-17:00", "17:00-17:45", "18:00-18:45",
    "18:45-19:30", "19:45-20:30"
]



def makeICS(cal, cell, group, week, day, hour):
    """
    Diese Funktion fügt ein Event einem Kalender (cal) hinzu.

    :param cal: Der Kalender, dem das Event hinzugefügt wird.
    :param cell: Der Inhalt der Zelle, der den Kurs beschreibt (<fachname>#<lehrkraftname>#<raum>).
    :param group: Die Gruppennummer (int).
    :param week: Die Wochennummer (int).
    :param day: Die Tagesnummer (int)
    :param hour: Die Stundenindexnummer (int).
    """
    # Daten aus der Zelle extrahieren
    if not cell or pd.isna(cell) or str(cell).isspace():
        print(f"Leere Zelle bei: Gruppe{group}_Woche{week}_Tag{day + 1}_Stunde{hour + 1}")
        return  # Wenn die Zelle leer ist, nichts tun

    try:
        fachname, lehrkraftname, raum = cell.split("#")[:3]

        # Besonderer Fall: Fach ist präsenti und damit aufgeteilt in 2 Gruppen
        if str(fachname).__contains__("präsenti"):
            fachname = "präsenti"
            lehrkraftname = str(cell).replace("#", " ").replace("präsenti", "")
            raum = ""

        if str(fachname).__contains__("ZP_VI"):
            fachname = "ZP"
            lehrkraftname = ""
            raum = ""

    except ValueError:
        print(f"Fehler bei der Verarbeitung der Zelle: Gruppe{group}_Woche{week}_Tag{day + 1}_Stunde{hour + 1}")
        return

    # Startzeit und Endzeit bestimmen
    try:
        time_range = class_times[hour]
    except IndexError:
        # Wenn die Stundenindexnummer außerhalb des Bereichs liegt, nichts tun
        return
    start_time_str, end_time_str = time_range.split("-")

    # Datum und Zeiten kombinieren
    day_offset = (week - 1) * 7 + day
    event_date = start_date + datetime.timedelta(days=day_offset)

    # Kombinieren des Datums mit den Zeiten


    start_time = datetime.datetime.strptime(start_time_str, "%H:%M").time()
    end_time = datetime.datetime.strptime(end_time_str, "%H:%M").time()

    start_datetime = datetime.datetime.combine(event_date, start_time)
    end_datetime = datetime.datetime.combine(event_date, end_time)

    # Naive Zeiten, keine Zeitzoneninformationen
    start_datetime_naive = start_datetime.replace(tzinfo=None)
    end_datetime_naive = end_datetime.replace(tzinfo=None)

    # kein bock mehr
    start_datetime_naive = start_datetime_naive - datetime.timedelta(hours=1)
    end_datetime_naive = end_datetime_naive - datetime.timedelta(hours=1)

    #bruh jetzt reichts
    # Überprüfen, ob es Freitag ist und ob die Startzeit nach 11:30 Uhr liegt
    print(str(day) + "  " + str(start_datetime_naive.time()) + " " + str(datetime.time(11, 30)))
    if day == 4 and start_datetime_naive.time() >= datetime.time(11, 00):
        # 15 Minuten abziehen
        start_datetime_naive -= datetime.timedelta(minutes=15)
        end_datetime_naive -= datetime.timedelta(minutes=15)
        print("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA")

    # Debug-Ausgabe zur Überprüfung
    #print(f"DEBUG: start_datetime_naive={start_datetime_naive}, end_datetime_naive={end_datetime_naive}")
    # Erstellen des Ereignisses
    event = Event()
    event.name = str(fachname).strip()
    event.begin = start_datetime_naive
    event.end = end_datetime_naive
    event.description = str(lehrkraftname).strip()
    event.location = str(raum).strip()
    event.created = datetime.datetime.now()  # DTSTAMP

    # Event dem Kalender hinzufügen
    cal.events.add(event)


def week_to_ics(data, group, week):
    """
    Diese Funktion druckt die Daten eines bestimmten Tabellenblatts basierend auf der Gruppen- und Wochennummer und erstellt eine ICS-Datei für die gesamte Woche.

    :param data: Dictionary, das die DataFrames enthält.
    :param group: Die Gruppennummer (int).
    :param week: Die Wochennummer (int).
    """

    sheet_name = f"Gruppe{group}_Woche{week}"
    if sheet_name in data:
        print(f"Daten für {sheet_name}:")
        df = data[sheet_name]
        cal = Calendar()  # Neuer Kalender für die Woche
        for row_idx, row in df.iterrows():
            for col_idx, cell in enumerate(row):
                makeICS(cal, cell, group, week, col_idx, row_idx)
        # Kalender in eine ICS-Datei speichern
        ics_filename = f"HfoeD2\\Gruppe{group}_Woche{week}.ics"
        with open(ics_filename, 'w', encoding='utf-8') as f:
            cal_str = str(cal)
            # Entfernen von zusätzlichen Zeilenumbrüchen
            cal_str = "\n".join([line for line in cal_str.splitlines() if line.strip()])
            f.write(cal_str)
        print(f"ICS-Datei für Woche erstellt: {ics_filename}")
    else:
        print(f"Keine Daten gefunden für {sheet_name}.")

def print_data(data, group, week):
    """
    Diese Funktion druckt die Daten eines bestimmten Tabellenblatts basierend auf der Gruppen- und Wochennummer und erstellt eine ICS-Datei für die gesamte Woche.

    :param data: Dictionary, das die DataFrames enthält.
    :param group: Die Gruppennummer (int).
    :param week: Die Wochennummer (int).
    """
    sheet_name = f"Gruppe{group}_Woche{week}"
    if sheet_name in data:
        print(f"Daten für {sheet_name}:")
        df = data[sheet_name]
        for row_idx, row in df.iterrows():
            row_output = f"stunde {row_idx + 1:02d}:"
            for col_idx, cell in enumerate(row):
                cell_value = str(cell).strip()[:3]  # Nur die ersten 3 Zeichen, ohne führende/trailing Leerzeichen
                row_output += f" {cell_value:<3}"  # Ausgabe mit Padding für Ausrichtung
            print(row_output)

    else:
        print(f"Keine Daten gefunden für {sheet_name}.")

for group in range(1, groups+1):
    for week in range(1, week+1):
        week_to_ics(data, group, week)
