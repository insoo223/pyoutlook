import win32com.client

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access Inbox
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)  # Sort by newest first

# Read first 5 emails
for i in range(5):
    message = messages[i]
    print("Subject:", message.Subject)
    print("Sender:", message.SenderName)
    print("Body:", message.Body[:50])  # Print first 50 chars
    print("-" * 40)
