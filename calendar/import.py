import os
import subprocess
from datetime import datetime, timedelta
from ics import Calendar, Event
import win32com.client

# === Configurare ===
REPO_PATH = r"C:\Users\M67E313\repos\trmnl\trmnl-images\calendar"  # modifică după caz
ICS_FILE = "calendar.ics"
COMMIT_MSG = "Actualizare calendar Outlook"
CALENDAR_DAYS_AHEAD = 1  # câte zile să exporte

# === 1. Conectare la Outlook ===
print("➡️ Conectare la Outlook...")
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
calendar_folder = namespace.GetDefaultFolder(9)  # 9 = Calendar

today = datetime.now()
start_of_week = today - timedelta(days=today.weekday())  # luni
end_of_week = start_of_week + timedelta(days=6, hours=23, minutes=59)

print(f"➡️ Export săptămână: {start_of_week} până la {end_of_week}")

# === 2. Extrage și filtrează evenimente ===
items = calendar_folder.Items
items.IncludeRecurrences = True
items.Sort("[Start]")

# Debug: număr total de evenimente
print(f"📅 Total evenimente inițiale în calendar: {len(items)}")

restriction = "[Start] >= '{}' AND [End] <= '{}'".format(
    start_of_week.strftime("%m/%d/%Y %H:%M %p"),
    end_of_week.strftime("%m/%d/%Y %H:%M %p")
)
print(f"🔍 Restricție aplicată: {restriction}")
restricted_items = items.Restrict(restriction)

print(f"📅 Evenimente găsite după restrict: {len(restricted_items)}")

# === 3. Creează fișier ICS ===
calendar = Calendar()
evenimente_adaugate = 0

for item in restricted_items:
    try:
        e = Event()
        print(item.Subject)
        e.name = item.Subject
        e.begin = item.Start.Format("%Y-%m-%d %H:%M:%S")
        e.end = item.End.Format("%Y-%m-%d %H:%M:%S")
        # e.description = item.Body[:200] if item.Body else ""
        # e.location = item.Location

        calendar.events.add(e)
        evenimente_adaugate += 1

        print(f"✅ Eveniment adăugat: {e.name} ({e.begin} - {e.end})")
    except Exception as ex:
        print(f"⚠️ Eroare la un eveniment: {ex}")

# === 4. Salvează fișierul .ics ===
ics_path = os.path.join(REPO_PATH, ICS_FILE)
with open(ics_path, "w", encoding="utf-8") as f:
    f.writelines(calendar.serialize_iter())

print(f"💾 Fișier salvat: {ics_path}")
print(f"📦 Evenimente incluse: {evenimente_adaugate}")

# === 5. Git push ===
commands = [
    ["git", "add", ICS_FILE],
    ["git", "commit", "-m", COMMIT_MSG],
    ["git", "push"]
]

for cmd in commands:
    subprocess.run(cmd, cwd=REPO_PATH)

print("🚀 Git push finalizat.")