import win32com.client
import os
import sys
import argparse
from process_excel import process_excel_file

def email_scan(target_sender, target_subject, location, save_folder, force):
    # setup
    print("Setting up Outlook application...")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # "6" refers to the inbox folder

    # Access the "External" subfolder
    sub_folder = inbox.Folders["External"]
    messages = sub_folder.Items

    # Convert relative path to absolute path
    save_folder = os.path.abspath(save_folder)

    if not os.path.exists(save_folder):
        os.makedirs(save_folder)

    print(f"Saving attachments to: {save_folder}")
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
                            email_date = message.SentOn.strftime("%Y-%m-%d")
                            save_path = os.path.join(save_folder, f"{email_date}_{attachment.FileName}")
                            if os.path.exists(save_path) and not force:
                                print(f"File from {email_date} already exists. Ignoring and ending script.")
                                return
                            attachment.SaveAsFile(save_path)
                            print(f"Saved attachment: {attachment.FileName} to {save_path}")
                            process_excel_file(save_path, location)
                            return
        except Exception as e:
            print(f"Error processing message: {e}")

    print("Finished searching through emails.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Email Scanner")
    parser.add_argument("sender", help="Email sender address")
    parser.add_argument("subject", help="Email subject")
    parser.add_argument("location", help="Location (sheet name) in the Excel file")
    parser.add_argument("--f", action="store_true", help="Force processing all emails and override existing files")
    args = parser.parse_args()

    save_folder = "./Sample_file/"
    email_scan(args.sender, args.subject, args.location, save_folder, args.f)