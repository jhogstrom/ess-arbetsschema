import argparse
import datetime
import os

import openpyxl
import pandas as pd
from dotenv import load_dotenv
from openpyxl.styles import Border, Side

# Load .env file
load_dotenv()

# Get the defaults
default_file = os.getenv("REPORT_FILE")
default_date = os.getenv("REPORT_DATE")
default_template = os.getenv("TEMPLATE")
default_outdir = os.getenv("OUTDIR")

# Parse command line arguments
parser = argparse.ArgumentParser()
parser.add_argument("-f", "--file", default=default_file, help="Excel file to read")
parser.add_argument(
    "-d", "--date", default=default_date, help="Date to generate report for"
)
parser.add_argument(
    "-t", "--template", default=default_template, help="Template file to fill in"
)
parser.add_argument(
    "-o", "--outdir", default=default_outdir, help="Filename to write the output to"
)
args = parser.parse_args()


# Iterate over the lines
def boats_filter(row, report_date):
    # Perform your test here. This is just an example.
    if "Arbetspass" in row["Schema"]:
        return False
    if row["Datum"] != report_date:
        return False
    if pd.isna(row["Medlem (fullt namn)"]):
        return False
    return True


def work_filter(row, report_date):
    # Perform your test here. This is just an example.
    if "Arbetspass" not in row["Schema"]:
        return False
    if row["Datum"] != report_date:
        return False
    if pd.isna(row["Medlem (fullt namn)"]):
        return False
    return True


def get_dates(df: pd.DataFrame) -> list:
    year = datetime.datetime.now().year
    return list(
        set(
            [
                row["Datum"]
                for _, row in df.iterrows()
                if datetime.datetime.strptime(row["Datum"], "%Y-%m-%d").year == year
                and "SJÖSÄTTNING" in row["Schema"].upper()
            ]
        )
    )


def make_report(date: str, df: pd.DataFrame, outdir: str, template: str):
    print(f"Generating report for {date}")

    # Read the Excel file
    boatrows = sorted(
        [_ for i, _ in df.iterrows() if boats_filter(_, date)],
        key=lambda x: x["Pass tid"],
    )

    work_rows = sorted(
        [_ for i, _ in df.iterrows() if work_filter(_, date)],
        key=lambda x: x["Pass tid"],
    )

    # Load the template Excel file
    wb = openpyxl.load_workbook(template)

    # Select the sheet where you want to add the matchrows
    sheet = wb["Sheet1"]  # Replace 'Sheet1' with the name of your sheet

    # Specify the starting row and column
    start_row = 5

    # Define the border
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    def add_cell(sheet, row, col, value):
        cell = sheet.cell(row=row, column=col, value=value)
        cell.border = thin_border

    # Write the matchrows to the sheet
    for i, row in enumerate(boatrows, start=start_row):
        sheet.insert_rows(i)
        add_cell(sheet, i, 1, row["Pass tid"])
        namn = row["Medlem (fullt namn)"].split("(")
        add_cell(sheet, i, 2, namn[1][:-1])
        add_cell(sheet, i, 3, namn[0].strip())
        add_cell(sheet, i, 4, str(int(row["Mobil"])))
        add_cell(sheet, i, 5, row["Plats"])
        add_cell(sheet, i, 6, row["Modell"])
        kommentar = (
            row["Kommentar medlem"] if not pd.isna(row["Kommentar medlem"]) else ""
        )
        add_cell(sheet, i, 7, kommentar)
        esk = (
            "ESK: " + row["inställningESK"]
            if not pd.isna(row["inställningESK"])
            else None
        )
        dusk1 = (
            "DUSK1: " + row["inställningDUSK"]
            if not pd.isna(row["inställningDUSK"])
            else None
        )
        dusk2 = (
            "DUSK2: " + row["InställningDUSK2"]
            if not pd.isna(row["InställningDUSK2"])
            else None
        )
        settings = ", ".join(_ for _ in [esk, dusk1, dusk2] if _ is not None)
        add_cell(sheet, i, 8, settings)

    for i, row in enumerate(work_rows, start=start_row + len(boatrows) + 4):
        sheet.insert_rows(i)
        add_cell(sheet, i, 1, row["Pass tid"])
        namn = row["Medlem (fullt namn)"].split("(")
        add_cell(sheet, i, 2, namn[1][:-1])
        add_cell(sheet, i, 3, namn[0].strip())
        add_cell(sheet, i, 4, str(int(row["Mobil"])))

    sheet.cell(1, 1, f"Sjösättning ESS {date}")
    sheet.cell(
        row=1, column=7, value=datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    )

    # Save the workbook
    filename = os.path.join(outdir, "Förarschema ESS " + date + ".xlsx")
    if not os.path.exists(os.path.dirname(filename)):
        os.makedirs(os.path.dirname(filename))
    wb.save(filename)
    print(f"\tSjösättningar: {len(boatrows)}")
    print(f"\tArbetspass: {len(work_rows)}")
    print(f"Report written to '{filename}'")


df = pd.read_excel(args.file)
dates = get_dates(df)
for d in dates:
    make_report(d, df, args.outdir, template=args.template)
