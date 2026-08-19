"""Building ``vbaProject.bin`` from VBA source, on any platform.

The OLE container and its ``dir``/``PROJECT`` streams come from the ``ms-ovba``
package; the module text is compressed by :mod:`.compression` instead of that
package's own encoder, which corrupts anything past ~2 KB.

Module text must stay ASCII. Streams are written in the project's code page, and a
Cyrillic ``Sub`` name would round-trip through cp1252 as mojibake — see
:func:`~pharmparser.export.vba.macros.macro_identifier`, which is why generated
macro names are transliterated.
"""

from __future__ import annotations

import logging
import os
import uuid
from pathlib import Path
from tempfile import TemporaryDirectory

from .compression import compress

logger = logging.getLogger(__name__)

PROJECT_ID = "{9E394C0B-697E-4AEE-9FA6-446F51FB30DC}"
SHEET_GUID = uuid.UUID("0002082000000000C000000000000046")
WORKBOOK_GUID = uuid.UUID("0002081900000000C000000000000046")
OLE_AUTOMATION_GUID = uuid.UUID("0002043000000000C000000000000046")
OFFICE_GUID = uuid.UUID("2DF8D04C5BFA101BBDE500AA0044DE52")


class VbaBuildError(RuntimeError):
    """Raised when the VBA project could not be assembled."""


def _patched_compression() -> None:
    """Point ms-ovba's module writer at our encoder.

    ms-ovba offers no seam for this, and its own encoder corrupts anything past
    ~2 KB, so the method is replaced outright. ``test_ovba_compression`` and
    ``test_macro_export_without_excel`` both fail loudly if this stops taking
    effect after an upgrade.
    """
    from ms_ovba_compression.ms_ovba import MsOvba

    MsOvba.compress = lambda self, data: compress(bytes(data))  # type: ignore[method-assign]


def build(modules: dict[str, str], sheet_names: list[str] | None = None) -> bytes:
    """Compile ``{module name: VBA source}`` into a ``vbaProject.bin``.

    ``sheet_names`` is unused by the macros themselves but keeps the project's
    document modules aligned with the workbook Excel will open it against.
    """
    if not modules:
        raise VbaBuildError("A VBA project needs at least one module.")

    for name, source in modules.items():
        if not source.isascii():
            offending = sorted({c for c in source if not c.isascii()})
            raise VbaBuildError(
                f"Module {name!r} contains non-ASCII characters {offending!r}; "
                "VBA module streams are written in the project code page."
            )

    try:
        from ms_ovba.Models.Entities.doc_module import DocModule
        from ms_ovba.Models.Entities.reference import Reference
        from ms_ovba.Models.Entities.reference_registered import ReferenceRegistered
        from ms_ovba.Models.Entities.std_module import StdModule
        from ms_ovba.Models.Fields.libid_reference import LibidReference
        from ms_ovba.vbaProject import VbaProject
        from ms_ovba.Views.project_ole_file import ProjectOleFile
    except ImportError as e:  # pragma: no cover - dependency is declared
        raise VbaBuildError(f"ms-ovba is not installed: {e}") from e

    _patched_compression()

    project = VbaProject()
    project.project_id = PROJECT_ID

    with TemporaryDirectory() as directory:
        workspace = Path(directory)

        # Excel expects a document module per worksheet plus one for the workbook.
        for index, _ in enumerate(sheet_names or ["Sheet1"], start=1):
            project.add_module(_document_module(DocModule, workspace, f"Sheet{index}", SHEET_GUID))
        project.add_module(_document_module(DocModule, workspace, "ThisWorkbook", WORKBOOK_GUID))

        for name, source in modules.items():
            path = workspace / f"{name}.bas"
            path.write_text(_with_name_attribute(name, source), encoding="ascii", newline="\n")
            module = StdModule(name)
            module.add_file(str(path))
            module.normalize_file()
            project.add_module(module)

        for guid, version, library, description, alias in (
            (OLE_AUTOMATION_GUID, "2.0", r"C:\Windows\System32\stdole2.tlb", "OLE Automation", "stdole"),
            (
                OFFICE_GUID,
                "2.0",
                r"C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL",
                "Microsoft Office 16.0 Object Library",
                "Office",
            ),
        ):
            reference = ReferenceRegistered(LibidReference(guid, version, "0", library, description))
            project.add_reference(Reference(reference, alias))

        # ProjectOleFile writes vbaProject.bin into the current directory.
        previous = Path.cwd()
        try:
            os.chdir(workspace)
            ProjectOleFile.write_file(project)
        finally:
            os.chdir(previous)

        built = workspace / "vbaProject.bin"
        if not built.is_file():
            raise VbaBuildError("ms-ovba did not produce a vbaProject.bin")
        payload = built.read_bytes()

    logger.debug("Built vbaProject.bin (%d bytes) from %d module(s)", len(payload), len(modules))
    return payload


def _document_module(doc_module_class, workspace: Path, name: str, guid: uuid.UUID):
    """An empty ``ThisWorkbook``/``SheetN`` class module."""
    path = workspace / f"{name}.cls"
    path.write_text(
        f'VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\nAttribute VB_Name = "{name}"\n',
        encoding="ascii",
        newline="\n",
    )
    module = doc_module_class(name)
    module.add_file(str(path))
    module.add_guid(guid)
    module.normalize_file()
    return module


def _with_name_attribute(name: str, source: str) -> str:
    if source.lstrip().startswith("Attribute VB_Name"):
        return source
    return f'Attribute VB_Name = "{name}"\n{source}'
