import win32com.client
import csv

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
messages = inbox.Items

# Load CSV with Delete flags
csv_file = "delete_sender.csv"
delete_senders = set()

with open(csv_file, mode="r", encoding="utf-8-sig") as file:
    reader = csv.DictReader(file)
    for row in reader:
        if row.get("Delete") == "1":
            delete_senders.add(row["SenderName"])

print("Senders marked for deletion:", delete_senders)

# Iterate through messages and delete if sender matches
to_delete = []
for message in list(messages):  # convert to list for safe iteration
    try:
        if message.SenderName in delete_senders:
            to_delete.append(message)
    except Exception:
        continue

print(f"Found {len(to_delete)} messages to delete.")

# Delete messages
for msg in to_delete:
    try:
        msg.Delete()
    except Exception as e:
        print("Error deleting message:", e)

print("Deletion complete.")
