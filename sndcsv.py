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
        # Skip items that aren't standard mail messages
        continue

# Count occurrences
counter = Counter(sender_names)

# Sort by count (descending)
sorted_senders = sorted(counter.items(), key=lambda x: x[1], reverse=True)

# Export to CSV
csv_file = "outlook_sender_counts.csv"
with open(csv_file, mode="w", newline="", encoding="utf-8") as file:
    writer = csv.writer(file)
    writer.writerow(["SenderName", "MessageCount"])  # Header row
    for sender, count in sorted_senders:
        writer.writerow([sender, count])

print(f"Export complete! Results saved to {csv_file}")
