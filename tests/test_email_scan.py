import pytest
from unittest.mock import MagicMock, patch
import sys
import os

# Add the parent directory to sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from email_scan import email_scan

@pytest.fixture
def mock_outlook():
    with patch('win32com.client.Dispatch') as mock_dispatch:
        mock_namespace = MagicMock()
        mock_folder = MagicMock()
        mock_message = MagicMock()
        mock_attachment = MagicMock()

        mock_dispatch.return_value.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_folder
        mock_folder.Folders["External"].Items = [mock_message]

        mock_message.Class = 43
        mock_message.SenderEmailAddress = "ddkroese40@gmail.com"
        mock_message.Subject = "Test email spoof"
        mock_message.Attachments = [mock_attachment]
        mock_message.SentOn.strftime.return_value = "2024-01-01"

        mock_attachment.FileName = "testEmailSpoof.xlsx"
        mock_attachment.SaveAsFile = MagicMock()

        yield mock_dispatch

def test_email_scan(mock_outlook, test_excel_file):
    email_scan("ddkroese40@gmail.com", "Test email spoof", "loc3", "./Sample_file/", False)
    mock_outlook.return_value.GetNamespace.return_value.GetDefaultFolder.return_value.Folders["External"].Items[0].Attachments[0].SaveAsFile.assert_called_once()