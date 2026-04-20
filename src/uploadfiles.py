import json
import os

from dotenv import load_dotenv

from driveapi import upload_to_folder
from platsplanering import setup_logger

load_dotenv()

if __name__ == "__main__":
    logger = setup_logger("uploadfiles", os.getenv("DEBUG_LEVEL", "DEBUG"))

    filedata = json.load(open("stage/generated_files.json", encoding="utf-8"))
    folder_id = filedata.get("parent_folder_id", "")
    files = []
    [files.extend(_) for _ in filedata.get("files", {}).values()]

    upload_to_folder(folder_id=folder_id, files=files, logger=logger)
    logger.info("Upload completed.")
