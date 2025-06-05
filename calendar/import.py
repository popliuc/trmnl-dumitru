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

# === 1. Extrage evenimente din Outlook ===
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
calendar_folder = namespace.GetDefaultFolder(9)  # 9 = Calendar

today = datetime.now()
end_day = today + timedelta(days=CALENDAR_DAYS_AHEAD)

items = calendar_folder.Items
items.IncludeRecurrences = True
items.Sort("[Start]")

restriction = "[Start] >= '{}' AND [End] <= '{}'".format(
    today.strftime("%m/%d/%Y %H:%M %p"), end_day.strftime("%m/%d/%Y %H:%M %p")
)
restricted_items = items.Restrict(restriction)

# === 2. Creează fișier ICS ===
calendar = Calendar()

for item in restricted_items:
    try:
        e = Event()
        e.name = item.Subject
        e.begin = item.Start.Format("%Y-%m-%d %H:%M:%S")
        e.end = item.End.Format("%Y-%m-%d %H:%M:%S")
        e.description = item.Body[:200] if item.Body else ""
        e.location = item.Location
        calendar.events.add(e)
    except Exception as ex:
        print(f"Eveniment sărit: {ex}")

# === 3. Salvează fișierul .ics ===
ics_path = os.path.join(REPO_PATH, ICS_FILE)
with open(ics_path, "w", encoding="utf-8") as f:
    f.writelines(calendar.serialize_iter())

# === 4. Git push ===
commands = [
    ["git", "add", ICS_FILE],
    ["git", "commit", "-m", COMMIT_MSG],
    ["git", "push"],
]

for cmd in commands:
    subprocess.run(cmd, cwd=REPO_PATH)
