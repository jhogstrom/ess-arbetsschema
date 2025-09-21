import base64
import logging
import mimetypes
import os
from email.message import EmailMessage
from pprint import pprint
from typing import List, Optional

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from googleapi import get_credentials


def _add_attachments(
    *,
    message: EmailMessage,
    attachments: Optional[List[str]] = None,
    logger: logging.Logger,
) -> None:
    """Add file attachments to an EmailMessage.

    This helper function processes a list of file paths and attaches them to
    the provided EmailMessage object. It automatically detects MIME types,
    validates file existence, and handles binary file reading.

    Args:
        message: EmailMessage object to add attachments to
        attachments: Optional list of file paths to attach. If None or empty,
            no attachments will be added
        logger: Logger instance for logging operations and status messages

    Returns:
        None: Function modifies the message object in-place

    Raises:
        FileNotFoundError: If any attachment file doesn't exist at the specified path
        IOError: If file cannot be read (permissions, corruption, etc.)

    Example:
        >>> message = EmailMessage()
        >>> _add_attachments(
        ...     message=message,
        ...     attachments=["report.pdf", "data.xlsx", "image.png"],
        ...     logger=logger
        ... )
        >>> len(message.get_payload())  # Now includes attachments
        4  # Original content + 3 attachments
    """
    if not attachments:
        return
    for attachment_path in attachments:
        if not os.path.exists(attachment_path):
            raise FileNotFoundError(f"Attachment file not found: {attachment_path}")

        # Guess the content type based on the file's extension
        ctype, encoding = mimetypes.guess_type(attachment_path)
        if ctype is None or encoding is not None:
            # No guess could be made, or the file is encoded (compressed)
            ctype = "application/octet-stream"

        maintype, subtype = ctype.split("/", 1)

        # Read and attach the file
        with open(attachment_path, "rb") as attachment_file:
            message.add_attachment(
                attachment_file.read(),
                maintype=maintype,
                subtype=subtype,
                filename=os.path.basename(attachment_path),
            )
        logger.info(f"Attached file: {attachment_path}")


def gmail_send_message(
    *,
    rec_to: List[str],
    rec_cc: Optional[List[str]] = None,
    rec_bcc: Optional[List[str]] = None,
    content: str,
    subject: str,
    attachments: Optional[List[str]] = None,
    logger: logging.Logger,
    dry_run: Optional[bool] = False,
) -> Optional[dict]:
    """Create and send an email message via Gmail API.

    Args:
        receivers: List of recipient email addresses
        content: Email body content as UTF-8 encoded string
        attachments: Optional list of file paths to attach to the email

    Returns:
        Message object with message ID if successful, None if failed

    Raises:
        HttpError: If Gmail API request fails
        FileNotFoundError: If attachment file doesn't exist

    Example:
        >>> gmail_send_message(
        ...     receivers=["user@example.com"],
        ...     content="Hello world",
        ...     attachments=["report.pdf", "data.xlsx"]
        ... )
    """
    creds = get_credentials()
    try:
        service = build("gmail", "v1", credentials=creds)
        message = EmailMessage()

        # Convert plain text to HTML to preserve line formatting
        html_content = content.replace("\n", "<br>")
        html_content = f'<html><body><pre style="font-family: Arial, sans-serif; white-space: pre-wrap;">{html_content}</pre></body></html>'

        # Use HTML content type to avoid automatic line wrapping
        message.set_payload(html_content.encode("utf-8"))
        message.set_charset("utf-8")
        message.set_type("text/html")

        message["To"] = ",".join(rec_to)
        message["Cc"] = ",".join(rec_cc) if rec_cc else ""
        message["Bcc"] = ",".join(rec_bcc) if rec_bcc else ""
        # message["From"] = "varvschef@edsvikensss.se"
        message["Subject"] = subject

        # Add attachments if provided
        _add_attachments(message=message, attachments=attachments, logger=logger)

        # encoded message
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

        create_message = {"raw": encoded_message}
        # pylint: disable=E1101
        if dry_run:
            logger.info("Dry run enabled - not sending email")
            for key in message.keys():
                logger.info(f"{key}: {message[key]}")
            for attachment in message.iter_attachments():
                logger.info(f"Attachment: {attachment.get_filename()}")
            return None
        send_message = (
            service.users().messages().send(userId="me", body=create_message).execute()
        )
        logger.info(f'Message Id: {send_message["id"]}')
    except HttpError as error:
        logger.error(f"An error occurred: {error}")
        send_message = None
    return send_message


if __name__ == "__main__":
    # Set up logging for testing
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)

    with open("templates/email-template.html", encoding="utf-8") as f:
        content = f.read()

    rec_cc = os.getenv("REC_CC", "").split(",")
    rec_bcc = os.getenv("REC_BCC", "").split(",")
    rec_to = os.getenv("REC_TO", "").split(",")

    res = gmail_send_message(
        rec_to=rec_to,
        rec_cc=rec_cc,
        rec_bcc=rec_bcc,
        subject="Nästa upptagning/ESS",
        content=content,
        logger=logger,
        attachments=[
            "stage/Förarschema ESS 2025-09-21.pptx",
            "stage/Förarschema ESS 2025-09-21.xlsx",
        ],
        dry_run=False,
    )
    pprint(res)
