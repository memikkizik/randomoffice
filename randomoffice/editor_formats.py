"""Explicit page and cell formatting supported by the local editor."""

from copy import copy
from colorsys import hls_to_rgb, rgb_to_hls
import math
import re
from typing import Any


PAGE_MARGINS = {
    "normal": (1, 1, 1, 1),
    "narrow": (0.5, 0.5, 0.5, 0.5),
    "wide": (1, 1, 2, 2),
}


def validate_page_layout(layout: Any) -> None:
    if not isinstance(layout, dict):
        raise ValueError("page_layout must be an object")
    for field, values in (("orientation", ("portrait", "landscape")),
                          ("paper", ("A4", "Letter")),
                          ("margin", tuple(PAGE_MARGINS))):
        if field in layout and layout[field] not in values:
            raise ValueError(f"Invalid page_layout {field}")


def apply_page_layout(document, layout: dict[str, Any]) -> None:
    """Set real DOCX section dimensions, including HTML-based exports."""
    from docx.enum.section import WD_ORIENT
    from docx.shared import Inches, Mm

    validate_page_layout(layout)
    paper = layout.get("paper", "A4")
    landscape = layout.get("orientation", "portrait") == "landscape"
    width, height = (Mm(210), Mm(297)) if paper == "A4" else (Inches(8.5), Inches(11))
    top, bottom, left, right = PAGE_MARGINS[layout.get("margin", "normal")]
    for section in document.sections:
        section.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
        section.page_width, section.page_height = (height, width) if landscape else (width, height)
        section.top_margin, section.bottom_margin = Inches(top), Inches(bottom)
        section.left_margin, section.right_margin = Inches(left), Inches(right)


def read_page_layout(document) -> dict[str, str]:
    """Map the first section to the editor's supported paper/margin presets."""
    from docx.shared import Inches, Mm

    section = document.sections[0]
    width, height = section.page_width, section.page_height
    landscape = width > height
    short, long = sorted((width, height))
    papers = {"A4": (Mm(210), Mm(297)), "Letter": (Inches(8.5), Inches(11))}
    paper = min(papers, key=lambda key: abs(short-papers[key][0]) + abs(long-papers[key][1]))
    actual = (section.top_margin, section.bottom_margin, section.left_margin, section.right_margin)
    margin = min(PAGE_MARGINS, key=lambda key: sum(abs(value-Inches(expected))
                 for value, expected in zip(actual, PAGE_MARGINS[key]) if value is not None))
    return {"orientation": "landscape" if landscape else "portrait", "paper": paper, "margin": margin}


def cell_address(key: str) -> tuple[int, int]:
    if not isinstance(key, str) or not re.fullmatch(r"(?:0|[1-9]\d*):(?:0|[1-9]\d*)", key):
        raise ValueError("Cell formatting addresses must use zero-based row:column")
    row, column = map(int, key.split(":"))
    if row >= 1_048_576 or column >= 16_384:
        raise ValueError("Cell formatting address exceeds Excel worksheet limits")
    return row + 1, column + 1


def validate_cell_styles(styles: Any) -> None:
    if not isinstance(styles, dict) or len(styles) > 200_000:
        raise ValueError("cell_styles must be an object with at most 200,000 entries")
    for key, style in styles.items():
        cell_address(key)
        if not isinstance(style, dict):
            raise ValueError("Each cell style must be an object")
        for field in ("bold", "italic", "underline", "strike"):
            if field in style and not isinstance(style[field], bool):
                raise ValueError(f"Cell style {field} must be boolean")
        for field in ("color", "background"):
            if field in style and (not isinstance(style[field], str) or not re.fullmatch(r"#[0-9a-fA-F]{6}", style[field])):
                raise ValueError(f"Cell style {field} must use #RRGGBB")
        if "font_name" in style and (not isinstance(style["font_name"], str) or not 1 <= len(style["font_name"]) <= 200):
            raise ValueError("Cell style font_name must be 1-200 characters")
        if "font_size" in style:
            size = style["font_size"]
            if isinstance(size, bool) or not isinstance(size, (int, float)) or not math.isfinite(size) or not 1 <= size <= 400:
                raise ValueError("Cell style font_size must be between 1 and 400 points")
        if "alignment" in style and style["alignment"] not in ("left", "center", "right", "justify"):
            raise ValueError("Cell style alignment must be left, center, right, or justify")


def apply_cell_style(cell, style: dict[str, Any]) -> None:
    from openpyxl.styles import PatternFill

    font = copy(cell.font)
    for key, field in (("bold", "bold"), ("italic", "italic"), ("strike", "strike"),
                       ("font_name", "name"), ("font_size", "size")):
        if key in style:
            setattr(font, field, style[key])
    if "underline" in style:
        font.underline = "single" if style["underline"] else None
    if "color" in style:
        font.color = style["color"][1:]
    cell.font = font
    if "background" in style:
        cell.fill = PatternFill(fill_type="solid", fgColor=style["background"][1:])
    if "alignment" in style:
        alignment = copy(cell.alignment)
        alignment.horizontal = style["alignment"]
        cell.alignment = alignment


def workbook_theme_colors(workbook) -> list[str]:
    """Resolve common theme colors without loading any external resources."""
    from lxml import etree

    if not workbook.loaded_theme:
        return []
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    try:
        root = etree.fromstring(workbook.loaded_theme, parser=parser)
        namespaces = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        scheme = root.find(".//a:clrScheme", namespaces)
        if scheme is None:
            return []
        # Spreadsheet theme indexes use lt1/dk1/lt2/dk2, rather than XML order.
        colors = {}
        for entry in scheme:
            if len(entry):
                element = entry[0]
                colors[etree.QName(entry).localname] = element.get("lastClr") or element.get("val", "")
        return [colors.get(name, "") for name in ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2",
                "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink")]
    except (ValueError, TypeError, etree.XMLSyntaxError):
        return []


def _cell_color(color, theme: list[str]) -> str | None:
    from openpyxl.styles.colors import COLOR_INDEX

    if color is None:
        return None
    value = None
    if color.type == "rgb":
        value = color.rgb[-6:]
    elif color.type == "indexed" and 0 <= color.indexed < len(COLOR_INDEX):
        value = COLOR_INDEX[color.indexed][-6:]
    elif color.type == "theme" and 0 <= color.theme < len(theme):
        value = theme[color.theme]
    if not isinstance(value, str) or not re.fullmatch(r"[0-9a-fA-F]{6}", value):
        return None
    # DrawingML tint modifies HLS luminance, retaining the theme color's hue.
    tint = color.tint or 0
    channels = [int(value[index:index+2], 16) for index in (0, 2, 4)]
    if tint:
        hue, luminance, saturation = rgb_to_hls(*(component/255 for component in channels))
        luminance = luminance*(1+tint) if tint < 0 else luminance*(1-tint)+tint
        channels = [round(component*255) for component in hls_to_rgb(hue, luminance, saturation)]
    return "#" + "".join(f"{max(0, min(255, component)):02X}" for component in channels)


def read_cell_style(cell, theme: list[str]) -> dict[str, Any]:
    if not cell.has_style:
        return {}
    font = cell.font
    style: dict[str, Any] = {"bold": bool(font.bold), "italic": bool(font.italic),
                            "underline": bool(font.underline), "strike": bool(font.strike)}
    if font.name:
        style["font_name"] = font.name
    if font.size is not None:
        style["font_size"] = font.size
    color = _cell_color(font.color, theme)
    if color:
        style["color"] = color
    if cell.fill.patternType == "solid":
        background = _cell_color(cell.fill.fgColor, theme)
        if background:
            style["background"] = background
    if cell.alignment.horizontal in ("left", "center", "right", "justify"):
        style["alignment"] = cell.alignment.horizontal
    return style
