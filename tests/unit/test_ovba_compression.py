"""MS-OVBA compression, checked against an independent decoder.

Every case round-trips through oletools' ``decompress_stream``, which is a separate
implementation of the same specification — so these tests verify the encoder rather
than merely agreeing with themselves.

The published pure-Python encoder (``ms-ovba`` 0.0.1) passes at small sizes and then
silently corrupts everything past ~2 KB, at the CopyToken bit-split boundary at
decompressed offset 2048. The sweep below is aimed squarely at that.
"""

from __future__ import annotations

import os
import random

import pytest
from oletools.olevba import decompress_stream

from pharmparser.export.vba.ovba.compression import CHUNK, SAFE_PARTIAL, compress

VBA = """Sub SortASCENDINGD_dannye()
    Application.ScreenUpdating = False
    ActiveSheet.Range("A3:F100000").Sort Key1:=ActiveSheet.Columns("D"), Order1:=xlAscending
    Application.ScreenUpdating = True
End Sub
"""


def roundtrip(data: bytes) -> bytes:
    return bytes(decompress_stream(bytearray(compress(data))))


@pytest.mark.parametrize("data", [b"", b"A", b"Option Explicit\r\n", VBA.encode()])
def test_small_inputs(data: bytes) -> None:
    assert roundtrip(data) == data


@pytest.mark.parametrize("size", [1, 255, 2047, 2048, 2049, 4095, 4096, 4097, 8192, 20000])
def test_repetitive_input_of_every_interesting_size(size: int) -> None:
    data = (VBA.encode() * (size // len(VBA) + 2))[:size]
    assert roundtrip(data) == data


@pytest.mark.parametrize("size", [2047, 2048, 2049, 3639, 3640, 3641, 4095, 4096, 4097, 9000])
def test_incompressible_input_of_every_interesting_size(size: int) -> None:
    """A short final chunk may not use the raw form; it would be padded to 4096."""
    data = os.urandom(size)
    assert roundtrip(data) == data


def test_the_boundary_that_breaks_ms_ovba() -> None:
    """A module larger than 2 KB is exactly what this project generates."""
    data = (VBA * 40).encode()
    assert len(data) > 4000
    assert roundtrip(data) == data


def test_a_full_incompressible_chunk_is_stored_verbatim() -> None:
    data = os.urandom(CHUNK)
    blob = compress(data)
    assert blob[1:3] == (0x3000 | (CHUNK - 1)).to_bytes(2, "little")
    assert roundtrip(data) == data


def test_a_short_incompressible_chunk_is_split_rather_than_padded() -> None:
    data = os.urandom(SAFE_PARTIAL + 200)
    assert roundtrip(data) == data
    assert len(roundtrip(data)) == len(data), "padding would have made it longer"


def test_fuzz_across_the_copytoken_boundary() -> None:
    rng = random.Random(20260819)
    alphabets = [b"AB", b"Option Explicit Sub End \r\n", bytes(range(256))]
    for _ in range(200):
        size = rng.randrange(1, 6000)
        alphabet = rng.choice(alphabets)
        data = bytes(alphabet[rng.randrange(len(alphabet))] for _ in range(size))
        assert roundtrip(data) == data, f"size={size} alphabet={len(alphabet)}"
