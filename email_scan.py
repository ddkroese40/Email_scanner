import win32com.client
import os

# setup
print("Setting up Outlook application...")
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # "6" refers to the inbox folder


# messages = inbox.Items
# TODO replace with below if email not recieved to base Inbox to access a subfolder
sub_folder = inbox.Folders["External"]
messages = sub_folder.Items

# your criteria
target_sender = "ddkroese40@gmail.com"
target_subject = "Test email spoof"
save_folder = r"./Sample_file/"

# Convert relative path to absolute path
save_folder = os.path.abspath(save_folder)

if not os.path.exists(save_folder):
    os.makedirs(save_folder)

print("Starting to search through emails...")

# search through emails
for message in messages:
    try:
        if message.Class == 43:
            if message.SenderEmailAddress == target_sender and target_subject.lower() in message.Subject.lower():
                print("Found target email. Checking for attachments...")
                # save attachments
                attachments = message.Attachments
                for attachment in attachments:
                    if attachment.FileName.lower().endswith('.xlsx'):  # specify file types to save
                        save_path = os.path.join(save_folder, attachment.FileName)
                        attachment.SaveAsFile(save_path)
                        print(f"Saved attachment: {attachment.FileName} to {save_path}")
    except Exception as e:
        print(f"Error processing message: {e}")

print("Finished searching through emails.")