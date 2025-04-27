import argparse
from email_scan import email_scan

def main():
    parser = argparse.ArgumentParser(description="Email Scanner")
    parser.add_argument("sender", nargs='?', default="ddkroese40@gmail.com", help="Email sender address")
    parser.add_argument("subject", nargs='?', default="Test email spoof", help="Email subject")
    parser.add_argument("location", nargs='?', default="loc3", help="Location (sheet name) in the Excel file")
    parser.add_argument("--f", action="store_true", help="Force processing all emails and override existing files")
    args = parser.parse_args()

    save_folder = "./Sample_file/"

    # Call the email_scan function directly with the provided arguments
    email_scan(args.sender, args.subject, args.location, save_folder, args.f)

if __name__ == "__main__":
    main()