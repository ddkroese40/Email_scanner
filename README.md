# Email_scanner

This app scans emails for a specific sender and subject, downloads all attachments, and processes Excel attachments to report on data from a specified sheet.

## Features

1. Scan emails from a specific sender and subject.
2. Download and save Excel attachments.
3. Process the Excel file to sum up "power used" by month.
4. Save the processed data to an output Excel file.
5. Optional force flag to override existing files.

## Usage


```python main.py --sender "example@example.com" --subject "Custom Subject" --sheet "CustomSheet" --f```
Arguments
sender: Email sender address (default: ddkroese40@gmail.com)
subject: Email subject (default: Test email spoof)
location: Location (sheet name) in the Excel file (default: loc3)
--f: Force processing all emails and override existing files
Example
```python main.py "example@example.com" "Custom Subject" "CustomSheet" --f```

## Changing Code
Change Email Folder: Modify the folder being scanned in email_scan.py: or just change to `messages = inbox.Items` for base Inbox only
```sub_folder = inbox.Folders["External"]```

## Testing
Running Tests
To run the tests, use the following command:

```pytest```