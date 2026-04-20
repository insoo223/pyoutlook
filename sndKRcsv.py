import win32com.client
from collections import Counter
import csv

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access Inbox
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
messages = inbox.Items

# Collect sender names
sender_names = []
for message in messages:
    try:
        sender_names.append(message.SenderName)
    except Exception:
        continue

# Count occurrences
counter = Counter(sender_names)

# Sort by count (descending)
sorted_senders = sorted(counter.items(), key=lambda x: x[1], reverse=True)

# Export to CSV with UTF-8 BOM (Excel-friendly)
csv_file = "outlook_sender_counts.csv"
with open(csv_file, mode="w", newline="", encoding="utf-8-sig") as file:
    writer = csv.writer(file)
    writer.writerow(["SenderName", "MessageCount"])
    for sender, count in sorted_senders:
        writer.writerow([sender, count])

print(f"Export complete! Results saved to {csv_file}")
