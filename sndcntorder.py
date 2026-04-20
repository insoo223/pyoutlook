import win32com.client
from collections import Counter

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
        # Some items may not be mail messages (like meeting requests)
        continue

# Count occurrences
counter = Counter(sender_names)

# Sort by count (descending)
sorted_senders = sorted(counter.items(), key=lambda x: x[1], reverse=True)

# Print results
for sender, count in sorted_senders:
    print(f"{sender}: {count}")
