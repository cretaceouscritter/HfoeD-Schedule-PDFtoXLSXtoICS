import pandas as pd
import warnings
import datetime
import uuid
import pytz
import re
from icalendar import Calendar, Event

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def infer_groups_weeks(filepath: str) -> tuple[int, int]:
    """Liest max. Gruppe/Woche aus Sheetnamen wie 'Gruppe3_Woche9'"""
    xls = pd.ExcelFile(filepath)
    mxg = mxw = 0
    for name in xls.sheet_names:
        m = re.match(r"^Gruppe(\d+)_Woche(\d+)$", name, re.I)
        if m:
            mxg = max(mxg, int(m.group(1)))
            mxw = max(mxw, int(m.group(2)))
    return mxg, mxw


filepath = "OUTPUT_XLSX/Stundenplaene.xlsx"

groups, weeks = infer_groups_weeks(filepath)
start_date = datetime.date.fromisoformat(input("Startdatum (YYYY-MM-DD): ").strip())

BERLIN = pytz.timezone("Europe/Berlin")

class_times = [
    "08:00-08:45", "08:45-09:30", "10:00-10:45", "10:45-11:30",
    "12:00-12:45", "12:45-13:30", "13:30-14:15", "14:30-15:15",
    "15:15-16:00", "16:15-17:00", "17:00-17:45", "18:00-18:45",
    "18:45-19:30", "19:45-20:30"
]

data = {}
for group in range(1, groups + 1):
    for week in range(1, weeks + 1):
        sheet_name = f"Gruppe{group}_Woche{week}"
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
        data[sheet_name] = df


def makeICS(cal, cell, group, week, day, hour):
    if not cell or pd.isna(cell) or str(cell).isspace():
        return

    try:
        fachname, lehrkraftname, raum = cell.split("#")[:3]

        if "ZP_VI" in str(fachname):  # Umbenennung Spezialfall
            fachname = "ZP"
            lehrkraftname = ""
            raum = ""

    except ValueError:
        print(f"Fehler bei Zelle: Gruppe{group}_Woche{week}_Tag{day + 1}_Stunde{hour + 1}")
        return

    try:
        start_time_str, end_time_str = class_times[hour].split("-")
    except IndexError:
        return

    day_offset = (week - 1) * 7 + day
    event_date = start_date + datetime.timedelta(days=day_offset)

    start_time = datetime.datetime.strptime(start_time_str, "%H:%M").time()
    end_time = datetime.datetime.strptime(end_time_str, "%H:%M").time()

    start_dt_naive = datetime.datetime.combine(event_date, start_time)
    end_dt_naive = datetime.datetime.combine(event_date, end_time)

    # Freitag = day == 4 (wenn Mo=0..Fr=4) und ab 11:30 -> 15 min früher
    if day == 4 and start_dt_naive.time() >= datetime.time(11, 30):
        start_dt_naive -= datetime.timedelta(minutes=15)
        end_dt_naive -= datetime.timedelta(minutes=15)

    # DST korrekt je nach Datum (Sommer/Winterzeit) via pytz.localize
    start_dt = BERLIN.localize(start_dt_naive, is_dst=None)
    end_dt = BERLIN.localize(end_dt_naive, is_dst=None)

    start_utc = start_dt.astimezone(pytz.UTC)
    end_utc = end_dt.astimezone(pytz.UTC)

    event = Event()
    event.add("uid", f"{uuid.uuid4()}@hfoed")
    event.add("summary", str(fachname).strip())
    event.add("dtstart", start_utc)  # <- UTC => DTSTART:...Z
    event.add("dtend", end_utc)  # <- UTC => DTEND:...Z
    event.add("dtstamp", datetime.datetime.now(datetime.timezone.utc))

    desc = str(lehrkraftname).strip()
    loc = str(raum).strip()
    if desc:
        event.add("description", desc)
    if loc:
        event.add("location", loc)

    cal.add_component(event)


def week_to_ics(data, group, week):
    sheet_name = f"Gruppe{group}_Woche{week}"
    if sheet_name not in data:
        print(f"Keine Daten gefunden für {sheet_name}.")
        return

    df = data[sheet_name]

    cal = Calendar()
    cal.add("prodid", "-//HfoeD Schedule Converter//")
    cal.add("version", "2.0")
    cal.add("calscale", "GREGORIAN")

    for row_idx, row in df.iterrows():
        for col_idx, cell in enumerate(row):
            makeICS(cal, cell, group, week, col_idx, row_idx)

    ics_filename = f"OUTPUT_ICS\\Gruppe{group}_Woche{week}.ics"
    with open(ics_filename, "wb") as f:
        f.write(cal.to_ical())

    print(f"ICS-Datei erstellt: {ics_filename}")


for group in range(1, groups + 1):
    for week in range(1, weeks + 1):
        week_to_ics(data, group, week)
