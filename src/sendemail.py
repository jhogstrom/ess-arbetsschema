import argparse
import json

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
        default="templates/email-template.html",
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


if __name__ == "__main__":
    args = parse_args()
    logger = setup_logger("mail")

    filedata = json.load(open("stage/generated_files.json", encoding="utf-8"))

    date = sorted(list(filedata.get("files", {}).keys()))[0]

    files = filedata.get("files", {}).get(date, [])

    emailfile = [_ for _ in files if _.endswith(MAILFILE_EXT)][0]

    logger.info(f"Selected email file for date {date}: {emailfile}")

    with open(emailfile, encoding="utf-8") as f:
        recipients = [_.strip() for _ in f.readlines()]

    recipients = ["jspr.hgstrm@gmail.com"]  # For testing only

    logger.info(f"Number of recipients: {len(recipients)}")

    with open(args.template, encoding="utf-8") as f:
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
