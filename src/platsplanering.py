import argparse
import datetime
import glob
import json
import logging
import math
import os
import re
from typing import Any, Dict, Hashable, List, Optional

import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.slide import Slide
from pptx.util import Cm, Inches, Pt

logger = logging.getLogger("spaceplan")
logger.setLevel(logging.INFO)
formatter = logging.Formatter("%(name)s: %(asctime)s - %(levelname)s - %(message)s")
ch = logging.StreamHandler()
ch.setFormatter(formatter)
logger.addHandler(ch)

FILL_COLOR = RGBColor(214, 245, 214)
FILL_COLOR_NOSPOT = RGBColor(255, 230, 230)
FILL_COLOR_ON_LAND = RGBColor(230, 230, 255)
FILL_COLOR_MEMBER_LEFT = RGBColor(255, 153, 255)


def define_colors(filename: Optional[str]) -> Dict[str, RGBColor]:
    result = {
        "reserved": RGBColor(214, 245, 214),
        "declined": RGBColor(255, 230, 230),
        "member_left": RGBColor(255, 153, 255),
        "on_land": RGBColor(230, 230, 255),
        "unknown": RGBColor(255, 255, 255),
    }

    if filename and os.path.exists(filename):
        try:
            with open(filename) as f:
                user_colors = json.load(f)
            try:
                for key, value in user_colors.items():
                    result[key] = RGBColor(*value)
            except ValueError:
                logger.error(f"Could not read colors from file {filename}")
        except json.JSONDecodeError:
            logger.error(f"Could not read colors from file {filename}")
    return result


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
        # "--boats", default="Båtmått*.xlsx", help="Excel file with boat information"
        "--members",
        default="Alla_medlemmar_inkl_båtinfo_*.xlsx",
        help="Excel file with boat information",
    )
    parser.add_argument(
        "--outfile",
        default=f"stage/varvskarta {datetime.datetime.now().year}.pptx",
        help="Filename for output PowerPoint file",
    )
    parser.add_argument(
        "--exmembers",
        default="boatinfo/ex-members.txt",
        help="Filename with ex-members",
    )
    parser.add_argument(
        "--onland",
        default="boatinfo/sommarliggare.xlsx",
        help="Excel file with members already on land",
    )
    parser.add_argument(
        "--scheduled",
        default="boatinfo/torrsättning.xlsx",
        help="Excel file with members already on land",
    )
    return parser.parse_args()


def make_filename(filename: str, *, dirs: Optional[List[str]] = None) -> str:
    if not filename:
        raise ValueError("No filename provided")
    if os.path.exists(filename):
        logger.debug(f"File {filename} => {filename}")
        return filename
    dirs = [] or dirs
    for d in dirs:
        f = os.path.join(d, filename)
        if os.path.exists(f):
            logger.debug(f"File {filename} => {f}")
            return f

    matches = []
    for d in dirs:
        logger.debug(f"Searching for {filename} in {os.path.join(d, filename)}")
        matches.extend(glob.glob(os.path.join(d, filename)))
    if not matches:
        raise FileNotFoundError(
            f"File {filename} not found in {dirs} using pattern {filename}"
        )
    result = max([_ for _ in matches if "~" not in _], key=os.path.getmtime)
    logger.debug(f"File {filename} => {result}")
    return result


def read_pptx_file(filename: str) -> Presentation:
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
        org = m
        if not isinstance(m, int):
            logger.warning(f"Item '{m}' is not an integer")
            for c in m:
                if not c.isdigit():
                    m = m.replace(c, "")
            try:
                thelist[i] = int(m)
            except ValueError:
                logger.error(f"Could not convert '{org}' to integer")
    return thelist


def members_on_land(filename: str) -> list[int]:
    logger.info(f"Reading who's on land from file {filename}")
    members = pd.read_excel(filename)
    year = datetime.datetime.now().year
    return members.loc[members["År"] == year, "Medlemsnr"].tolist()


def member_left_club(filename: str) -> list[int]:
    with open(filename) as f:
        lines = [_ for _ in f.readlines() if not _.startswith("#")]

    # Regular expression to find the first number in each line
    number_pattern = re.compile(r"\d+")

    # Extract the first number from each line
    return [
        int(number_pattern.search(line).group())
        for line in lines
        if number_pattern.search(line)
    ]


def get_scheduled(filename: str) -> list[int]:
    schedule = pd.read_excel(filename)
    scheduled = [
        int(_) for _ in set(schedule["Medlemsnr"].tolist()) if not math.isnan(_)
    ]
    return scheduled


def get_no_spot_requested(filename: str) -> list[int]:
    # TODO: Parametrize the NO_SPOT_OPTION string
    request_data = pd.read_excel(filename)
    NO_SPOT_OPTION = (
        "Jag vill INTE ta upp min båt i år och vill INTE ha nån vinterplats hos ESS"
    )
    result = make_items_integer(
        request_data.loc[
            request_data["Upptagning"] == NO_SPOT_OPTION, "Medlemsnummer"
        ].tolist()
    )
    return result


def get_requests(filename: str) -> list[int]:
    logger.info(f"Reading requests file {filename}")
    request_data = pd.read_excel(filename)
    result = make_items_integer(request_data["Medlemsnummer"].tolist())
    logger.debug(f"Read {len(result)} requests")
    result = set(result)
    logger.debug(f"Read {len(result)} unique requests")
    return result


def read_members(filename: str) -> List[Dict[Hashable, Any]]:
    logger.info(f"Reading member file {filename}")
    boats = pd.read_excel(filename)
    # Filter out the boats that are not in the requests
    # boats = boats[boats["Medlemsnr"].isin(all_requests)]
    # Filter out the columns we are interested in
    boats = boats[
        [
            "Medlemsnr",
            "Längd (båt)",
            "Bredd",
            "Förnamn",
            "Efternamn",
            "Plats",
            # "Modell",
        ]
    ].to_dict(orient="records")
    logger.info(f"Read {len(boats)} boats from member file that are in the requests.")
    return boats


def get_boats(
    *,
    members: List[Dict[Hashable, Any]],
    already_there: List[int],
    scheduled: List[int],
    no_spot_requested: List[int],
    all_requests,
) -> list:
    # TODO: Parametrize the column names

    extra = 0
    for id in scheduled:
        if id not in all_requests:
            logger.warning(f"Member {id} not in requests, but booked for dry dock")
            all_requests.add(id)
            extra += 1
    logger.debug(f"Added {extra} scheduled boats to requests")

    for id in already_there:
        if id not in all_requests:
            logger.warning(f"Member {id} not in requests, but already on land")
            all_requests.add(id)

    for r in all_requests:
        if r not in [b["Medlemsnr"] for b in members]:
            logger.warning(f"Member {r} not found in boat file")

    members = [b for b in members if b["Medlemsnr"] in all_requests]
    for member in members:
        member["member"] = int(member.pop("Medlemsnr"))
        member["length"] = float(member.pop("Längd (båt)")) + 1
        member["width"] = float(member.pop("Bredd")) + 1
        # boat['name'] = f"{boat.pop('Förnamn')[0]} {boat.pop('Efternamn')}\n({boat.pop('Modell')})"
        member["name"] = f"{member.pop('Efternamn')}"
        member["requested"] = member["member"] not in no_spot_requested

    # Make boats unique by member id
    members = {boat["member"]: boat for boat in members}
    result = members.values()
    logger.info(f"After deduplication: {len(result)} unique boats")

    return result


def make_shape_name(member_id: int) -> str:
    return f"Member: {member_id}"


def get_shape(slide, shape_name: str):
    """
    Get the shape by member id or member name by looking at
    the text in the shape. If the id or name is in the text,
    return the shape.

    Args:
        slide: Slide in pptx
        member_id (str): Member id to search for

    Returns:
        Shape: Shape, if found by member id or member name, otherwise None
    """
    for shape in slide.shapes:
        if shape.name == shape_name:
            logger.debug(f"Found shape '{shape_name}'.")
            return shape
    for shape in slide.shapes:
        if any(_ in shape.text for _ in shape_name.split()):
            logger.debug(f"Found shape {shape.name} matching {shape_name}")
            return shape
    logger.warning(f"Did not find shape with name '{shape_name}'")
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


def color_boats(slide, members: List[int], color: RGBColor, logmsg: str):
    for member in members:
        shape = get_shape(slide, make_shape_name(member))
        if shape:
            shape.fill.fore_color.rgb = color
            text = shape.text.replace("\n", "--")
            logger.info(f"Member {member} ('{text}') {logmsg}")
            shape.name = f"Member: {member}"
        else:
            logger.warning(f"Member {member} {logmsg}, but not found on map")


def format_shape(shape, fill_color: RGBColor) -> None:
    """
    Format the shape with the fill color provided.

    Args:
        shape: Shape in pptx
        fill_color (RGBColor): Fill color to set in the shape
    """
    LINE_COLOR = RGBColor(0, 0, 0)  # Black color
    LINE_WIDTH = Inches(0.01)  # Thin outline
    # Set the shape fill color to the color provided
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = fill_color

    # Set the shape outline to thin black
    line = shape.line
    line.color.rgb = LINE_COLOR
    line.width = LINE_WIDTH


def ensured_shape(slide, shape_name: str, boat: Dict[str, Any]):
    # TODO: Parametrize the SCALE_LENGTH and SCALE_WIDTH
    SCALE_LENGTH = 5
    SCALE_WIDTH = 5
    # If creating a new shape, set the left and top position to top left on the slide.
    left = Pt(1)
    top = Pt(1)

    shape = get_shape(slide, shape_name)
    if not shape:
        width = Pt(boat["length"] * SCALE_WIDTH)
        length = Pt(boat["width"] * SCALE_LENGTH)

        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, length)
    shape.name = shape_name
    return shape


def add_boats_to_map(
    slide: Slide, boats: list, *, already_there: List[int], ex_members: List[int]
):
    fills = {True: colors["reserved"], False: colors["declined"]}

    logger.info(f"Adding {len(boats)} boats to map")
    counts = {True: 0, False: 0}
    # Iterate over the list and create or reuse shapes
    for boat in boats:
        member = boat["member"]
        shape_name = f"Member: {member}"
        logger.debug(f"Adding boat {shape_name}")
        shape = ensured_shape(slide, shape_name, boat)

        format_shape(shape, fills[boat["requested"]])
        counts[boat["requested"]] += 1

        expansions = {
            "length": boat["length"],
            "width": boat["width"],
            "name": boat["name"],
            "member": member,
        }
        NAME_ONLY = "{member} {name}"  # noqa: F841
        SIZE_ONLY = "{member} {length:.1f}x{width:.1f}"  # noqa: F841
        FULL = "{member} {name}\n{length:.1f}x{width:.1f}"  # noqa: F841
        caption = FULL.format(**expansions)

        set_shape_text(shape, caption)
    logger.info(f"Spots count: {counts[True]}")
    logger.info(f"Yield count: {counts[False]}")
    color_boats(slide, ex_members, colors["member_left"], "has left the club")
    color_boats(slide, already_there, colors["on_land"], "is already on land")

    for s in slide.shapes:
        if "Rectangle" in s.name:
            logger.info(f"'{s.text}' has not requested a spot")
            s.fill.fore_color.rgb = colors["unknown"]
    for boat in boats:
        if not boat["requested"]:
            logger.info(f"{boat['member']} {boat['name']} has declined a spot")


def remove_shape_by_name(slide, shape_name: str) -> bool:
    """
    Remove the shape by name from the slide object and return True if removed, otherwise False.

    Args:
        slide: Slide in pptx
        shape_name (str): Name of the shape to remove

    Returns:
        bool: True if shape removed, otherwise False
    """
    shape = get_shape(slide, shape_name)
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


def update_legend(slide, colors):
    """
    Update the legend boxes with the current colors.

    Args:
        slide: Slide in pptx
        colors: Dictionary with colors for the legend boxes
    """
    for key, color in colors.items():
        shape_name = f"Legend: {key}"
        shape = get_shape(slide, shape_name)
        if not shape:
            logger.warning(f"Could not find shape '{shape_name}'")
        else:
            logger.debug(f"Setting color {color} for shape '{shape_name}'")
            shape.fill.fore_color.rgb = color


if __name__ == "__main__":
    args = parseargs()
    colors = define_colors("templates/colors.json")

    ex_members = member_left_club(make_filename(args.exmembers, dirs=["boatinfo"]))
    already_there = members_on_land(make_filename(args.onland, dirs=["boatinfo"]))
    all_requests = get_requests(make_filename(args.requests, dirs=["boatinfo"]))
    no_spot_requested = get_no_spot_requested(
        make_filename(args.requests, dirs=["boatinfo"])
    )
    members = read_members(make_filename(args.members, dirs=["boatinfo"]))
    scheduled = get_scheduled(make_filename(args.scheduled, dirs=["boatinfo"]))
    boats = get_boats(
        members=members,
        already_there=already_there,
        scheduled=scheduled,
        no_spot_requested=no_spot_requested,
        all_requests=all_requests,
    )

    ppt = read_pptx_file(make_filename(args.file, dirs=["templates"]))
    map_slide = ppt.slides[0]
    add_boats_to_map(
        slide=map_slide, boats=boats, ex_members=ex_members, already_there=already_there
    )
    update_revision(get_shape(map_slide, "Revision"), revision="1", boats=boats)
    update_legend(map_slide, colors)
    try:
        ppt.save(args.outfile)
    except PermissionError:
        logger.error(f"Could not save file {args.outfile}")
        logger.error("File is open in another application")
        exit(1)
