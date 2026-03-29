"""
Gmail API Service
Handles authentication, message retrieval, attachment download, and sending.
"""

import os
import pickle
import base64
import mimetypes
from pathlib import Path
from typing import Optional, List, Dict
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from config.settings import (
    GMAIL_SCOPES,
    CREDENTIALS_FILE,
    TOKEN_FILE,
    PERSONAL_EMAIL,
    WORK_EMAIL,
    REPORT_SUBJECT_FILTER,
)
from utils.logger import logger


class GmailService:
    """Gmail API wrapper for reading and sending emails."""

    def __init__(self):
        self.service = None
        self.authenticate()

    def authenticate(self):
        """Authenticate with Gmail API using OAuth2."""
        creds = None

        # Load existing token if available
        if TOKEN_FILE.exists():
            with open(TOKEN_FILE, "rb") as token:
                creds = pickle.load(token)

        # Refresh or create new token
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                logger.info("🔄 Refreshing Gmail OAuth2 token...")
                creds.refresh(Request())
            else:
                logger.info("📱 Creating new Gmail OAuth2 token...")
                flow = InstalledAppFlow.from_client_secrets_file(
                    CREDENTIALS_FILE, GMAIL_SCOPES
                )
                creds = flow.run_local_server(port=0)

            # Save token for next run
            with open(TOKEN_FILE, "wb") as token:
                pickle.dump(creds, token)

        self.service = build("gmail", "v1", credentials=creds)
        logger.info("✅ Gmail API authenticated")

    def get_unread_messages(self, query: str = "is:unread") -> List[Dict]:
        """
        Fetch unread messages matching a query.

        Args:
            query: Gmail search query (e.g., "is:unread from:work@company.com")

        Returns:
            List of message objects with ID, sender, subject
        """
        try:
            results = (
                self.service.users()
                .messages()
                .list(userId="me", q=query, maxResults=10)
                .execute()
            )
            return results.get("messages", [])
        except HttpError as error:
            logger.error(f"Gmail API error: {error}")
            return []

    def get_message_details(self, message_id: str) -> Dict:
        """
        Fetch full message details including headers and parts.

        Args:
            message_id: Gmail message ID

        Returns:
            Message object with full content
        """
        try:
            message = (
                self.service.users()
                .messages()
                .get(userId="me", id=message_id, format="full")
                .execute()
            )
            return message
        except HttpError as error:
            logger.error(f"Could not fetch message {message_id}: {error}")
            return {}

    def download_attachment(
        self,
        message_id: str,
        attachment_id: str,
        filename: str,
        save_dir: Path,
    ) -> Optional[Path]:
        """
        Download attachment from Gmail message.

        Args:
            message_id: Gmail message ID
            attachment_id: Attachment ID from message parts
            filename: Original filename
            save_dir: Directory to save attachment

        Returns:
            Path to saved file or None if failed
        """
        try:
            attachment = (
                self.service.users()
                .messages()
                .attachments()
                .get(userId="me", messageId=message_id, id=attachment_id)
                .execute()
            )

            # Decode base64url data
            file_data = base64.urlsafe_b64decode(attachment["data"])
            file_path = save_dir / filename

            with open(file_path, "wb") as f:
                f.write(file_data)

            logger.info(f"✅ Downloaded attachment: {file_path}")
            return file_path

        except Exception as error:
            logger.error(f"Could not download attachment {filename}: {error}")
            return None

    def extract_attachments(self, message: Dict, save_dir: Path) -> List[Path]:
        """
        Extract all attachments from a message.

        Args:
            message: Message object from get_message_details()
            save_dir: Directory to save attachments

        Returns:
            List of paths to downloaded files
        """
        downloaded_files = []
        save_dir.mkdir(parents=True, exist_ok=True)

        payload = message.get("payload", {})
        parts = payload.get("parts", [])

        for part in parts:
            if part.get("filename"):
                filename = part["filename"]
                attachment_id = part["body"].get("attachmentId")

                if attachment_id:
                    file_path = self.download_attachment(
                        message["id"], attachment_id, filename, save_dir
                    )
                    if file_path:
                        downloaded_files.append(file_path)

        return downloaded_files

    def get_sender_email(self, message: Dict) -> Optional[str]:
        """Extract sender email from message headers."""
        headers = message.get("payload", {}).get("headers", [])
        for header in headers:
            if header["name"] == "From":
                email = header["value"].split("<")[-1].rstrip(">")
                return email
        return None

    def get_message_subject(self, message: Dict) -> Optional[str]:
        """Extract subject from message headers."""
        headers = message.get("payload", {}).get("headers", [])
        for header in headers:
            if header["name"] == "Subject":
                return header["value"]
        return None

    def send_email(
        self,
        to_email: str,
        subject: str,
        body: str,
        attachments: Optional[List[Path]] = None,
    ) -> bool:
        """
        Send email via Gmail.

        Args:
            to_email: Recipient email
            subject: Email subject
            body: Email body (plain text)
            attachments: List of file paths to attach

        Returns:
            True if successful, False otherwise
        """
        try:
            message = MIMEMultipart()
            message["to"] = to_email
            message["subject"] = subject
            message.attach(MIMEText(body, "plain"))

            # Attach files
            if attachments:
                for file_path in attachments:
                    if not file_path.exists():
                        logger.warning(f"Attachment not found: {file_path}")
                        continue

                    with open(file_path, "rb") as attachment:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(attachment.read())

                    encoders.encode_base64(part)
                    part.add_header(
                        "Content-Disposition",
                        f"attachment; filename= {file_path.name}",
                    )
                    message.attach(part)

            # Send via Gmail API
            raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
            send_message = {"raw": raw_message}

            sent = (
                self.service.users()
                .messages()
                .send(userId="me", body=send_message)
                .execute()
            )

            logger.info(f"✅ Email sent to {to_email} (ID: {sent['id']})")
            return True

        except Exception as error:
            logger.error(f"Failed to send email: {error}")
            return False

    def mark_as_read(self, message_id: str) -> bool:
        """Mark message as read."""
        try:
            self.service.users().messages().modify(
                userId="me",
                id=message_id,
                body={"removeLabelIds": ["UNREAD"]},
            ).execute()
            logger.info(f"✅ Marked message {message_id} as read")
            return True
        except Exception as error:
            logger.error(f"Could not mark message as read: {error}")
            return False
