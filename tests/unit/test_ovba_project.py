"""Building vbaProject.bin, and the dependency bugs it has to work around.

ms-ovba is 0.0.1. Two of its defects would each make the .xlsm unbuildable, and
both are patched at the seam rather than worked around at the call site, so these
tests pin the patches: an upgrade that fixes them upstream should make these fail
loudly rather than silently re-break the export.
"""

from __future__ import annotations

from datetime import datetime

import pytest

from pharmparser.export.vba.ovba.project import (
    VbaBuildError,
    _patched_compression,
    _patched_filetime,
    build,
)

MODULE = """Option Explicit

Public Sub ApplyFilters()
    ActiveSheet.AutoFilterMode = False
End Sub
"""


def test_the_msfiletime_zero_date_does_not_go_through_a_timestamp() -> None:
    """The Windows failure, reproduced by its cause rather than its platform.

    ms-ovba builds its default date with ``datetime.timestamp()`` on 1601-01-01.
    Windows raises OSError [Errno 22] for any pre-epoch date there, so every
    VbaProject() call died before this patch — the export worked on Linux and not
    on the one platform the app ships to.
    """
    from ms_dtyp.filetime import Filetime

    _patched_filetime()
    source = Filetime.from_msfiletime.__func__  # type: ignore[attr-defined]
    assert "timestamp" not in source.__code__.co_names


@pytest.mark.parametrize("filetime", [0, 132537600000000000])
def test_filetime_round_trips_exactly(filetime: int) -> None:
    from ms_dtyp.filetime import Filetime

    _patched_filetime()
    assert Filetime.from_msfiletime(filetime).to_msfiletime() == filetime


def test_the_zero_filetime_is_the_epoch_the_format_defines() -> None:
    from ms_dtyp.filetime import Filetime

    _patched_filetime()
    assert Filetime.from_msfiletime(0) == datetime(1601, 1, 1)


def test_compression_is_replaced_with_ours() -> None:
    """ms-ovba's own encoder corrupts module text past ~2 KB."""
    from ms_ovba_compression.ms_ovba import MsOvba

    from pharmparser.export.vba.ovba.compression import compress

    _patched_compression()
    payload = (MODULE * 200).encode()
    assert MsOvba().compress(payload) == compress(payload)


def test_a_project_is_built_with_a_module_stream_per_sheet() -> None:
    olefile = pytest.importorskip("olefile")

    blob = build({"PharmParser": MODULE}, ["Данные", "Проценты", "Анализ"])
    with olefile.OleFileIO(blob if isinstance(blob, str) else __import__("io").BytesIO(blob)) as ole:
        streams = {"/".join(entry) for entry in ole.listdir()}

    assert "VBA/PharmParser" in streams
    assert "VBA/dir" in streams
    assert "PROJECT" in streams
    # one document module per worksheet, plus the workbook's
    assert {"VBA/Sheet1", "VBA/Sheet2", "VBA/Sheet3", "VBA/ThisWorkbook"} <= streams


def test_the_module_source_survives_the_build() -> None:
    olevba = pytest.importorskip("oletools.olevba")

    blob = build({"PharmParser": MODULE}, ["Данные"])
    parser = olevba.VBA_Parser("vbaProject.bin", data=blob)
    sources = {
        name: code for _, _, name, code in parser.extract_macros() if "PharmParser" in name
    }
    assert sources, "no PharmParser module came back out"
    assert "Public Sub ApplyFilters()" in next(iter(sources.values()))


def test_non_ascii_module_text_is_refused() -> None:
    """Module streams live in the project code page, so this would be mojibake."""
    with pytest.raises(VbaBuildError, match="non-ASCII"):
        build({"PharmParser": "Sub Данные()\nEnd Sub"}, ["Данные"])


def test_an_empty_project_is_refused() -> None:
    with pytest.raises(VbaBuildError, match="at least one module"):
        build({}, ["Данные"])
