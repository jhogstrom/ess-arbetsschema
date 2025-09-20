import argparse
from pprint import pprint

import pyperclip

from googleapi import get_google_sheet, get_sheet_titles
from platsplanering import (
    COL_TITLE_memberid,
    FileHelper,
    make_items_integer,
    read_members,
    setup_logger,
)


def parseargs():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        # "--boats", default="Båtmått*.xlsx", help="Excel file with boat information"
        "--members",
        default="Alla_medlemmar_inkl_båtinfo_*.xlsx",
        help="Excel file with boat information.",
    )
    parser.add_argument(
        "--sheetid",
    )
    return parser.parse_args()


args = parseargs()
logger = setup_logger("spots", "INFO")
fh = FileHelper(logger)
members_source = fh.make_filename(args.members, dirs=["boatinfo"])


def read_google_sheet(document: str):
    values = get_google_sheet(document, get_sheet_titles(document)[0])
    headers = values[0] if values else []
    COL_memberid = (
        headers.index(COL_TITLE_memberid) if COL_TITLE_memberid in headers else -1
    )
    result = [
        row[COL_memberid] for row in values[1:] if row and len(row) > COL_memberid
    ]
    pprint(result)
    pprint(len(sorted(result)))
    return sorted(make_items_integer(result))


members = read_members(
    members_source,
    columns=[
        "Medlemsnr",
        "Förnamn",
        "Efternamn",
        "Epost 1",
        "Längd (båt)",
        "Bredd",
    ],
)

if args.sheetid is None:
    members_to_email = [
        # Add list of memberid here to email for.
    ]
else:
    members_to_email = read_google_sheet(args.sheetid)

emails = list(
    set([f"{m['Epost 1']}" for m in members if m["Medlemsnr"] in members_to_email])
)

# Copy emails to clipboard
try:
    pyperclip.copy(",".join(emails))
    logger.info(f"Copied {len(emails)} email addresses to clipboard")
except Exception as e:
    logger.error(f"Failed to copy emails to clipboard: {e}")
    logger.info(f"Email addresses: {emails}")

pprint(emails)
