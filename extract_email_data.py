import pandas as pd
import win32com.client

# Initialize Outlook application
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.Folders.Item("sjdeepak4@gmail.com").Folders.Item("Test Folder")
messages = inbox.Items

# List to hold email data
email_data = []

for message in messages:
    if message.UnRead:
        email_data.append({
            'SenderName': message.SenderName,
            'SenderEmail': message.SenderEmailAddress,
            'Subject': message.Subject,
            'Date': message.ReceivedTime,
            'Body': message.Body
        })

# Create a DataFrame and save to CSV
df = pd.DataFrame(email_data)
df.to_csv(r'D:\\Code\\Projects\\UiPath Projects\\EASA\\email_data.csv', index=False)
