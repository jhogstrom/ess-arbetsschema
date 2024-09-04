import argparse
import datetime
import glob
import logging
import os
from typing import List, Optional

import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Cm, Inches, Pt

logger = logging.getLogger("spaceplan")
logger.setLevel(logging.INFO)
formatter = logging.Formatter("%(name)s: %(asctime)s - %(levelname)s - %(message)s")
ch = logging.StreamHandler()
ch.setFormatter(formatter)
logger.addHandler(ch)


def parseargs():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--file", default="*karta*.pptx", help="PowerPoint file to read"
    )
    parser.add_argument(
        "--requests",
        default="Anmälningar 2024.xlsx",
        help="Excel file with requests for spots",
    )
    parser.add_argument(
        "--boats", default="Båtmått*.xlsx", help="Excel file with boat information"
    )
    parser.add_argument(
        "--outfile",
        default="stage/karta.pptx",
        help="Filename for output PowerPoint file",
    )
    return parser.parse_args()


def make_filename(filename: str, *, dirs: Optional[List[str]] = None) -> str:
    if not filename:
        raise ValueError("No filename provided")
    if os.path.exists(filename):
        return filename
    dirs = [] or dirs
    for d in dirs:
        f = os.path.join(d, filename)
        if os.path.exists(f):
            return f

    matches = []
    for d in dirs:
        logger.debug(f"Searching for {filename} in {os.path.join(d, filename)}")
        matches.extend(glob.glob(os.path.join(d, filename)))
    if not matches:
        raise FileNotFoundError(
            f"File {filename} not found in {dirs} using pattern {filename}"
        )
    return max(matches, key=os.path.getmtime)


def read_file(filename: str) -> Presentation:
    """
    Read the PowerPoint file and return the Presentation object.

    Args:
        filename (str): Filename of the PowerPoint file to read

    Raises:
        ValueError: Raised if no filename is provided
        FileNotFoundError: Raised if the file is not found

    Returns:
        Presentation: PowerPoint Presentation object
    """
    logger.info(f"Reading file {filename}")

    # Open the PowerPoint file and return it
    return Presentation(filename)


def make_items_integer(thelist: list) -> List[int]:
    """
    Make all items in the list integers. If the item is not an integer, try to convert it to an integer by removing all non-numeric characters.

    Args:
        thelist (list): List of items to convert to integers

    Returns:
        List[int]: Returns a list of integers
    """
    for i, m in enumerate(thelist):
        if not isinstance(m, int):
            logger.warning(f"Item '{m}' is not an integer")
            for c in m:
                if not c.isdigit():
                    m = m.replace(c, "")
            thelist[i] = int(m)
    return thelist


def get_boats(request_filename: str, boats_filename: str) -> list:
    # TODO: Parametrize the NO_SPOT_OPTION string
    # TODO: Parametrize the column names
    NO_SPOT_OPTION = (
        "Jag vill INTE ta upp min båt i år och vill INTE ha nån vinterplats hos ESS"
    )

    logger.info(f"Reading requests file {request_filename}")
    # boats_filename = "boatinfo/Båtmått_20240904_1709.xlsx"
    request_data = pd.read_excel(request_filename)
    # Filter the DataFrame where 'Upptagning' is not equal to NO_SPOT_OPTION
    # space_requested_list = make_items_integer(requests.loc[requests['Upptagning'] != NO_SPOT_OPTION, 'Medlemsnummer'].tolist())
    no_spot_requested = make_items_integer(
        request_data.loc[
            request_data["Upptagning"] == NO_SPOT_OPTION, "Medlemsnummer"
        ].tolist()
    )

    all_requests = make_items_integer(request_data["Medlemsnummer"].tolist())

    boats = pd.read_excel(boats_filename)
    boats = boats[boats["Medlemsnr"].isin(all_requests)]
    boats = boats[
        [
            "Medlemsnr",
            "Längd (båt)",
            "Bredd (båt)",
            "Förnamn",
            "Efternamn",
            "Plats",
            "Modell",
        ]
    ].to_dict(orient="records")
    for boat in boats:
        boat["member"] = int(boat.pop("Medlemsnr"))
        boat["length"] = boat.pop("Längd (båt)") + 1
        boat["width"] = boat.pop("Bredd (båt)") + 1
        # boat['name'] = f"{boat.pop('Förnamn')[0]} {boat.pop('Efternamn')}\n({boat.pop('Modell')})"
        boat["name"] = f"{boat.pop('Efternamn')}"
        boat["requested"] = boat["member"] not in no_spot_requested

    # Duplicates in request file?
    if len(all_requests) != len(set(all_requests)):
        logger.warning(
            f"{len(all_requests) - len(set(all_requests))} Duplicate memberships number in request list found of {len(all_requests)}"
        )
    # Make boats unique by member id
    boats = {boat["member"]: boat for boat in boats}
    return boats.values()


def get_shape_by_name(slide, name: str):
    """
    Gets the shape by name from the slide object or None if not found

    Args:
        slide: Slide in pptx
        name (str): name of the shape to find

    Returns:
        Shape: Shape, if found by name, otherwise None
    """
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


def get_shape_by_member(slide, member_id: str, member_name: str):
    """
    Get the shape by member id or member name by looking at
    the text in the shape. If the id or name is in the text,
    return the shape.

    Args:
        slide: Slide in pptx
        member_id (str): Member id to search for
        member_name (str): Member name to search for

    Returns:
        Shape: Shape, if found by member id or member name, otherwise None
    """
    for shape in slide.shapes:
        if str(member_id) in shape.text_frame.text:
            logger.debug(f"Found shape with member {member_id}")
            return shape
        if member_name.upper() in shape.text_frame.text.upper():
            logger.debug(f"Found shape with member name {member_name}")
            return shape
    return None


def set_shape_text(shape, text: str) -> None:
    """
    Set the text of the shape to the text provided and format it.

    Args:
        shape: Shape in pptx
        text (str): Text to set in the shape
    """
    TEXT_COLOR = RGBColor(0, 0, 0)  # Set the text color to black
    TEXT_MARGIN = Cm(0.1)
    FONT_SIZE = 8
    text_frame = shape.text_frame
    text_frame.margin_left = TEXT_MARGIN
    text_frame.margin_right = TEXT_MARGIN
    text_frame.margin_top = TEXT_MARGIN
    text_frame.margin_bottom = TEXT_MARGIN
    text_frame.clear()  # Clear the existing text
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # Make the text
    p = text_frame.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(FONT_SIZE)
    run.font.color.rgb = TEXT_COLOR
    run.font.bold = False
    # Enable text autofit
    text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE


def add_boats_to_map(prs: Presentation, boats: list):
    # TODO: Parametrize the SCALE_LENGTH and SCALE_WIDTH
    # TODO: Parametrize the FILL_COLOR and FILL_COLOR_NOSPOT
    slide = prs.slides[0]
    SCALE_LENGTH = 5
    SCALE_WIDTH = 5
    FILL_COLOR = RGBColor(26, 255, 26)
    FILL_COLOR_NOSPOT = RGBColor(255, 0, 0)
    LINE_COLOR = RGBColor(0, 0, 0)  # Black color
    LINE_WIDTH = Inches(0.01)  # Thin outline
    # If creating a new shape, set the left and top position to top left on the slide.
    left = Pt(1)
    top = Pt(1)
    request_count = 0
    yield_count = 0
    fills = {True: FILL_COLOR, False: FILL_COLOR_NOSPOT}
    # Iterate over the list and create or reuse shapes
    for boat in boats:
        member = boat["member"]
        shape_name = f"Member: {member}"
        logger.debug(f"Adding boat {shape_name}")

        shape = get_shape_by_name(slide, shape_name)

        if shape:
            logger.debug(f"Shape {shape_name} already exists")
            continue

        shape = get_shape_by_member(slide, str(member), boat["name"])
        if shape:
            logger.debug(f"Reusing shape for member {member}")
        else:
            width = Pt(boat["length"] * SCALE_WIDTH)
            length = Pt(boat["width"] * SCALE_LENGTH)

            shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, left, top, width, length
            )
        shape.name = f"Member: {member}"
        # Set the shape fill color to yellow
        fill = shape.fill
        fill.solid()
        fill.fore_color.rgb = fills[boat["requested"]]
        if boat["requested"]:
            request_count += 1
        else:
            yield_count += 1

        # Set the shape outline to thin black
        line = shape.line
        line.color.rgb = LINE_COLOR
        line.width = LINE_WIDTH
        set_shape_text(shape, f"{member} {boat['name']}")
    logger.info(f"Request count: {request_count}")
    logger.info(f"Yield count: {yield_count}")


def remove_shape_by_name(slide, name: str) -> bool:
    """
    Remove the shape by name from the slide object and return True if removed, otherwise False.

    Args:
        slide: Slide in pptx
        name (str): Name of the shape to remove

    Returns:
        bool: True if shape removed, otherwise False
    """
    shape = get_shape_by_name(slide, name)
    if shape:
        sp = shape.element
        sp.getparent().remove(sp)
        return True
    return False


def update_revision(shape, revision: str = "1", boats: list = None):
    """
    Update the revision shape with the current date and time.

    Args:
        shape: Shape in pptx
    """
    text = [
        f"Revision {revision}",
        f"Båtar: {len(boats)}",
        datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
    ]
    shape.text = "\n".join(text)


if __name__ == "__main__":
    args = parseargs()

    ppt = read_file(make_filename(args.file, dirs=["templates"]))
    boats = get_boats(
        make_filename(args.requests, dirs=["boatinfo"]),
        make_filename(args.boats, dirs=["boatinfo"]),
    )
    add_boats_to_map(ppt, boats)
    slide1 = ppt.slides[0]
    update_revision(get_shape_by_name(slide1, "Revision"), revision="1", boats=boats)
    ppt.save(args.outfile)
