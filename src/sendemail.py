import argparse
import glob
import json
from datetime import datetime

from gmailapi import gmail_send_message
from helpers import setup_logger

MAILFILE_EXT = ".email.txt"


def parse_args():

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--receiver",
        required=True,
        help="Email address of the recipient (To field)",
    )
    parser.add_argument(
        "--template",
        default="templates/email-template*.html",
        help="Path to the email template file",
    )
    parser.add_argument(
        "--replacement",
        "-r",
        action="append",
        metavar="KEY=VALUE",
        help="Replacement variable in the format KEY=VALUE. Can be repeated.",
    )
    return parser.parse_args()


seasons = {
    "se": {
        1: "vår",
        2: "vår",
        3: "vår",
        4: "vår",
        5: "vår",
        6: "vår",
        7: "vår",
        8: "höst",
        9: "höst",
        10: "höst",
        11: "höst",
        12: "höst",
    },
    "en": {
        1: "spring",
        2: "spring",
        3: "spring",
        4: "spring",
        5: "spring",
        6: "spring",
        7: "spring",
        8: "autumn",
        9: "autumn",
        10: "autumn",
        11: "autumn",
        12: "autumn",
    },
}


def get_email_template(template_pattern: str, date: str) -> str:
    # First get a template for the specific date
    templates = glob.glob(template_pattern.replace("*", "-" + date))
    # Then try to get a season-based template
    if not templates:
        d = datetime.strptime(date, "%Y-%m-%d")
        for lang in ["se", "en"]:
            pattern = template_pattern.replace("*", f"-{seasons[lang][d.month]}")
            templates = glob.glob(pattern)
            if templates:
                return templates[0]

    # If none is found, use a fixed name
    if not templates:
        templates = glob.glob(template_pattern.replace("*", ""))

    # TODO: Consider matching the pattern as is before failing
    if not templates:
        raise FileNotFoundError(f"No template file found matching {template_pattern}")

    if len(templates) > 1:
        raise ValueError(f"Multiple template files found matching {template_pattern}")
    return templates[0]


if __name__ == "__main__":
    args = parse_args()
    logger = setup_logger("mail")

    filedata = json.load(open("stage/generated_files.json", encoding="utf-8"))

    date = [
        _
        for _ in sorted(list(filedata.get("files", {}).keys()))
        if _ >= datetime.now().strftime("%Y-%m-%d")
    ][0]

    files = filedata.get("files", {}).get(date, [])

    emailfile = [_ for _ in files if _.endswith(MAILFILE_EXT)][0]

    logger.info(f"Selected email file for date {date}: {emailfile}")

    with open(emailfile, encoding="utf-8") as f:
        recipients = [_.strip() for _ in f.readlines()]

    logger.info(f"Number of recipients: {len(recipients)}")

    template_file = get_email_template(args.template, date)
    logger.info(f"Using email template file: {template_file}")

    with open(template_file, encoding="utf-8") as f:
        content = f.read()

    replacements = (
        dict([_.split("=", 1) for _ in args.replacement if "=" in _])
        if args.replacement
        else {}
    )
    replacements["date"] = date
    for key, value in replacements.items():
        content = content.replace(f"{{{key}}}", value)

    gmail_send_message(
        rec_to=[args.receiver],
        rec_bcc=recipients,
        content=content,
        subject=f"Nästa upptagning/ESS - {date}",
        attachments=[_ for _ in files if not _.endswith(MAILFILE_EXT)],
        logger=logger,
        dry_run=False,
    )
