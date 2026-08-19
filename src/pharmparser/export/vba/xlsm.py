"""Turning an ``.xlsx`` into a macro-enabled ``.xlsm``, without Excel.

openpyxl writes the report but cannot draw form controls, and the old path drew them
by driving Excel over COM. This module does the same job by rewriting the package:
it adds the compiled VBA project, a legacy VML drawing per sheet holding the buttons,
and the relationships and content types that tie them together.

Everything here is ordinary OOXML plumbing, so it runs on any platform and is covered
by ``tests/unit/test_xlsm.py`` and ``tests/integration/test_export_xlsm.py``.
"""

from __future__ import annotations

import logging
import re
import shutil
import zipfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path

logger = logging.getLogger(__name__)

CONTENT_TYPES = "[Content_Types].xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
VBA_PART = "xl/vbaProject.bin"

MACRO_ENABLED_WORKBOOK = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
SHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
VBA_CONTENT_TYPE = "application/vnd.ms-office.vbaProject"
VBA_RELATIONSHIP = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
VML_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

BUTTON_WIDTH_POINTS = 64
BUTTON_HEIGHT_POINTS = 18

_CELL = re.compile(r"^([A-Z]+)(\d+)$")


@dataclass(frozen=True, slots=True)
class ButtonSpec:
    """A form-control button anchored at ``cell`` that runs ``macro`` when clicked."""

    cell: str
    caption: str
    macro: str

    def position(self) -> tuple[int, int]:
        """0-based (column, row) of the anchor cell."""
        match = _CELL.match(self.cell.upper())
        if match is None:
            raise ValueError(f"Not a cell reference: {self.cell!r}")
        letters, digits = match.groups()
        column = 0
        for letter in letters:
            column = column * 26 + (ord(letter) - ord("A") + 1)
        return column - 1, int(digits) - 1


def _escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _vml(buttons: Sequence[ButtonSpec], shape_id_base: int) -> str:
    """A legacy VML drawing holding one Excel form button per spec."""
    shapes = []
    for index, button in enumerate(buttons):
        column, row = button.position()
        shapes.append(
            f"""
  <v:shape id="_x0000_s{shape_id_base + index}" type="#_x0000_t201"
   style='position:absolute;margin-left:0pt;margin-top:0pt;
   width:{BUTTON_WIDTH_POINTS}pt;height:{BUTTON_HEIGHT_POINTS}pt;z-index:{index + 1};
   mso-wrap-style:tight' o:button="t" fillcolor="buttonFace [67]"
   strokecolor="windowText [64]" o:insetmode="auto">
   <v:fill color2="buttonFace [67]" o:detectmouseclick="t"/>
   <o:lock v:ext="edit" rotation="t"/>
   <v:textbox style='mso-direction-alt:auto' o:singleclick="f">
    <div style='text-align:center'>
     <font face="Calibri" size="220" color="#000000">{_escape(button.caption)}</font>
    </div>
   </v:textbox>
   <x:ClientData ObjectType="Button">
    <x:Anchor>{column}, 0, {row}, 0, {column + 1}, 0, {row + 1}, 0</x:Anchor>
    <x:PrintObject>False</x:PrintObject>
    <x:AutoFill>False</x:AutoFill>
    <x:FmlaMacro>[0]!{_escape(button.macro)}</x:FmlaMacro>
    <x:TextHAlign>Center</x:TextHAlign>
    <x:TextVAlign>Center</x:TextVAlign>
   </x:ClientData>
  </v:shape>"""
        )
    return f"""<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
 <v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201"
  path="m,l,21600r21600,l21600,xe">
  <v:stroke joinstyle="miter"/>
  <v:path shadowok="f" o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
  <o:lock v:ext="edit" shapetype="t"/>
 </v:shapetype>{"".join(shapes)}
</xml>
"""


def _sheet_parts(names: Sequence[str]) -> dict[str, str]:
    """Map ``xl/worksheets/sheetN.xml`` to the sheet name the workbook gives it."""
    return {f"xl/worksheets/sheet{index}.xml": name for index, name in enumerate(names, start=1)}


def _patch_content_types(xml: str) -> str:
    """Mark the workbook macro-enabled and declare the .bin and .vml parts."""
    xml = xml.replace(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        MACRO_ENABLED_WORKBOOK,
    )
    defaults = ""
    if 'Extension="bin"' not in xml:
        defaults += f'<Default Extension="bin" ContentType="{VBA_CONTENT_TYPE}"/>'
    if 'Extension="vml"' not in xml:
        defaults += '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
    if defaults:
        opening = xml.index(">", xml.index("<Types")) + 1
        xml = xml[:opening] + defaults + xml[opening:]
    return xml


def _patch_workbook_rels(xml: str) -> str:
    if VBA_RELATIONSHIP in xml:
        return xml
    identifier = _next_relationship_id(xml)
    entry = f'<Relationship Id="{identifier}" Type="{VBA_RELATIONSHIP}" Target="vbaProject.bin"/>'
    return xml.replace("</Relationships>", f"{entry}</Relationships>")


def _next_relationship_id(xml: str) -> str:
    used = {int(value) for value in re.findall(r'Id="rId(\d+)"', xml)}
    return f"rId{max(used, default=0) + 1}"


def _sheet_rels(target: str, existing: str | None = None) -> tuple[str, str]:
    """Add a VML relationship without discarding table/drawing relationships."""
    if existing is None:
        identifier = "rId1"
        xml = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{RELATIONSHIPS_NS}"></Relationships>'
        )
    else:
        identifier = _next_relationship_id(existing)
        xml = existing
    relationship = f'<Relationship Id="{identifier}" Type="{VML_RELATIONSHIP}" Target="{target}"/>'
    return xml.replace("</Relationships>", f"{relationship}</Relationships>"), identifier


def _add_legacy_drawing(sheet_xml: str, relationship_id: str) -> str:
    """Point the sheet at its VML drawing.

    openpyxl does not declare the relationship namespace on ``<worksheet>``, so the
    ``r:id`` attribute has to bring it along; without that the part is not
    well-formed XML and Excel rejects the workbook.
    """
    if "<legacyDrawing" in sheet_xml:
        return sheet_xml

    sheet_xml = _ensure_relationship_namespace(sheet_xml)
    tag = '<legacyDrawing r:id="' + relationship_id + '"/>'
    # legacyDrawing must be the last child of worksheet.
    return sheet_xml.replace("</worksheet>", tag + "</worksheet>")


def _ensure_relationship_namespace(sheet_xml: str) -> str:
    opening = sheet_xml.index("<worksheet")
    end = sheet_xml.index(">", opening)
    if "xmlns:r=" in sheet_xml[opening:end]:
        return sheet_xml
    declaration = ' xmlns:r="' + OFFICE_RELATIONSHIPS_NS + '"'
    return sheet_xml[:end] + declaration + sheet_xml[end:]


def package(
    source: Path,
    target: Path,
    vba_project: bytes,
    buttons: Mapping[str, Sequence[ButtonSpec]],
    sheet_names: Sequence[str],
) -> Path:
    """Write ``source`` back out as ``target``, macro-enabled and with buttons drawn."""
    parts = _sheet_parts(sheet_names)
    with_buttons = {part: name for part, name in parts.items() if buttons.get(name)}

    with zipfile.ZipFile(source) as original:
        names = original.namelist()
        payloads = {name: original.read(name) for name in names}

    payloads[CONTENT_TYPES] = _patch_content_types(payloads[CONTENT_TYPES].decode()).encode()
    payloads[WORKBOOK_RELS] = _patch_workbook_rels(payloads[WORKBOOK_RELS].decode()).encode()
    payloads[VBA_PART] = vba_project

    shape_id = 1025
    for index, (part, sheet_name) in enumerate(sorted(with_buttons.items()), start=1):
        specs = list(buttons[sheet_name])
        vml_part = f"xl/drawings/vmlDrawing{index}.vml"
        payloads[vml_part] = _vml(specs, shape_id).encode()
        shape_id += len(specs)

        rels_part = f"xl/worksheets/_rels/{Path(part).name}.rels"
        existing_rels = payloads.get(rels_part)
        rels, relationship_id = _sheet_rels(
            f"../drawings/vmlDrawing{index}.vml",
            existing_rels.decode() if existing_rels is not None else None,
        )
        payloads[rels_part] = rels.encode()
        payloads[part] = _add_legacy_drawing(payloads[part].decode(), relationship_id).encode()

    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, payload in payloads.items():
            archive.writestr(name, payload)

    logger.info(
        "Packaged %s with %d button(s) across %d sheet(s)",
        target.name,
        sum(len(specs) for specs in buttons.values()),
        len(with_buttons),
    )
    return target


def copy_without_macros(source: Path, target: Path) -> Path:
    """Fallback when no VBA project is available: ship the plain workbook."""
    shutil.copyfile(source, target)
    return target
