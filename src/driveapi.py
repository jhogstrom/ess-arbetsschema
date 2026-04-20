import logging
import mimetypes
import os
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List

from dotenv.main import load_dotenv
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload

from googleapi import get_credentials

load_dotenv()


def _find_file_in_folder(name: str, folder_id: str, service) -> str | None:
    """
    Search for a file by name within a specific Google Drive folder.

    Args:
        name: The filename to search for.
        folder_id: The ID of the folder to search in.
        service: An authenticated Google Drive API service instance.

    Returns:
        The file ID of the first matching file, or None if not found.
    """
    # pylint: disable=maybe-no-member
    query = f"name = '{name}' and '{folder_id}' in parents and trashed = false"
    result = (
        service.files()
        .list(
            q=query,
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        .execute()
    )
    files = result.get("files", [])
    return files[0]["id"] if files else None


def _upload_file(file: str, folder_id: str, logger: logging.Logger) -> None:
    """
    Upsert a file into the specified Google Drive folder.

    If a file with the same name already exists in the folder its content is
    updated in-place (preserving the file ID).  Otherwise a new file is
    created.  A thread-local service instance is used so this function is
    safe to call from multiple threads concurrently.

    Args:
        file: Local path to the file to upload.
        folder_id: ID of the destination Google Drive folder.
        logger: Logger instance for logging operations.

    Returns:
        None
    """
    try:
        if not os.path.exists(file):
            logger.error(f"File to upload not found: {file}")
            return
        service = get_service()
        filename = os.path.basename(file)
        mime_type, _ = mimetypes.guess_type(file)
        if mime_type is None:
            mime_type = "application/octet-stream"
        media = MediaFileUpload(file, mimetype=mime_type, resumable=False)
        # pylint: disable=maybe-no-member
        existing_id = _find_file_in_folder(filename, folder_id, service)
        if existing_id:
            service.files().update(
                fileId=existing_id,
                media_body=media,
                fields="id",
                supportsAllDrives=True,
            ).execute()
            logger.info(f'Updated existing file: "{file}".')
        else:
            file_metadata = {"name": filename, "parents": [folder_id]}
            service.files().create(
                body=file_metadata,
                media_body=media,
                fields="id",
                supportsAllDrives=True,
            ).execute()
            logger.info(f'Uploaded new file: "{file}".')
    except HttpError as error:
        logger.error(f"An error occurred: {error}")
        return


_thread_local = threading.local()


def get_service():
    """Return a thread-local Google Drive service instance.

    httplib2 (used internally by the API client) is not thread-safe, so each
    thread must own its own service object.

    Returns:
        An authenticated Google Drive v3 service instance local to the calling
        thread.
    """
    if not hasattr(_thread_local, "service"):
        creds = get_credentials()
        _thread_local.service = build(
            "drive", "v3", credentials=creds, cache_discovery=False
        )
    return _thread_local.service


def upload_to_folder(
    *, folder_id: str, files: List[str], logger: logging.Logger
) -> None:
    """
    Upload files to the specified folder in parallel.

    Each file is uploaded concurrently using a thread pool.  Existing files
    are updated in-place; new files are created.

    Args:
        folder_id: ID of the Google Drive folder to upload to.
        files: List of local file paths to upload.
        logger: Logger instance for logging operations.

    Returns:
        None

    Example:
        >>> upload_to_folder(folder_id="abc123", files=["file1.txt"], logger=logger)
    """
    with ThreadPoolExecutor() as executor:
        futures = {
            executor.submit(_upload_file, file, folder_id, logger): file
            for file in files
            if not file.endswith(".email.txt")
        }
        for future in as_completed(futures):
            exc = future.exception()
            if exc:
                logger.error(f'Upload failed for "{futures[future]}": {exc}')


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
