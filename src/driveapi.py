import logging
import mimetypes
import os
from typing import List

from dotenv.main import load_dotenv
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload

from googleapi import get_credentials

load_dotenv()


def _upload_file(file: str, folder_id: str, service, logger: logging.Logger) -> None:
    """
    Upload a file to the specified folder and prints file ID, folder ID
    Args: Id of the folder
    Returns: ID of the file uploaded
    """
    try:
        if not os.path.exists(file):
            logger.error(f"File to upload not found: {file}")
            return
        file_metadata = {"name": os.path.basename(file), "parents": [folder_id]}
        mime_type, _ = mimetypes.guess_type(file)
        if mime_type is None:
            mime_type = "application/octet-stream"
        media = MediaFileUpload(file, mimetype=mime_type, resumable=False)
        # pylint: disable=maybe-no-member
        uploaded_file = (
            service.files()
            .create(
                body=file_metadata,
                media_body=media,
                fields="id",
                supportsAllDrives=True,
            )
            .execute()
        )
        logger.info(f'File ID: "{uploaded_file.get("id")}".')
    except HttpError as error:
        logger.error(f"An error occurred: {error}")
        return


_service = None


def get_service():
    global _service
    if _service is None:
        creds = get_credentials()
        _service = build("drive", "v3", credentials=creds, cache_discovery=False)
    return _service


def upload_to_folder(
    *, folder_id: str, files: List[str], logger: logging.Logger
) -> None:
    """
    Upload files to the specified folder using concurrent futures.

    Args:
        folder_id: ID of the Google Drive folder to upload to
        files: List of file paths to upload
        logger: Logger instance for logging operations

    Returns:
        List[Future]: List of Future objects for the upload tasks that can be
        joined later to ensure all uploads complete

    Example:
        >>> futures = upload_to_folder(folder_id="abc123", files=["file1.txt"], logger=logger)
        >>> # Do other work...
        >>> for future in futures:
        ...     future.result()  # Wait for completion
    """
    service = get_service()

    for file in files:
        _upload_file(file, folder_id, service, logger)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("driveapi")
    folder_id = os.getenv("PARENT_FOLDER_ID", "")
    files_to_upload = [
        "README.md",
        "stage/Förarschema ESS 2025-09-21.pptx",
        "stage/Förarschema ESS 2025-09-21.xlsx",
    ]

    upload_to_folder(folder_id=folder_id, files=files_to_upload, logger=logger)
