"""MS-OVBA 2.4.1 run-length compression, encoder side.

VBA module source inside ``vbaProject.bin`` is stored in this format. Writing it
ourselves is what lets the ``.xlsm`` be produced without Excel.

The only published pure-Python encoder, ``ms-ovba`` 0.0.1, silently corrupts module
text past roughly 2 KB — its ``CopyToken`` bit split is wrong once the decompressed
position crosses 2048, which is the boundary where the length/offset allocation
changes. Every macro module this project generates is larger than that, so the
encoder lives here instead. ``tests/unit/test_ovba_compression.py`` round-trips the
output through oletools' independent decoder, including a fuzz sweep across that
boundary.
"""

from __future__ import annotations

CHUNK = 4096
"""A CompressedChunk holds at most this many decompressed bytes (2.4.1.1.4)."""

MAX_CANDIDATES = 64
"""Cap on how far back a match is searched; fewer matches only costs a little size."""

SAFE_PARTIAL = 3640
"""Largest short chunk whose worst case (n literals + ceil(n/8) flags) still fits
the 12-bit CompressedChunkSize field."""


def _bit_count(difference: int) -> int:
    """CopyToken Help (2.4.1.3.19.1): bits given to the offset at this position."""
    if difference <= 1:
        return 4
    return max((difference - 1).bit_length(), 4)


def _max_length(difference: int) -> int:
    """Longest run a single CopyToken can encode at this position."""
    return (0xFFFF >> _bit_count(difference)) + 3


def _compress_chunk(data: bytes, start: int, end: int) -> bytes:
    """Encode ``data[start:end]`` as a sequence of flag bytes and tokens."""
    out = bytearray()
    tokens = bytearray()
    flags = 0
    flag_count = 0
    position = start
    index: dict[bytes, list[int]] = {}

    while position < end:
        best_offset = best_length = 0
        if position + 3 <= end:
            limit = min(end, position + _max_length(position - start))
            for candidate in reversed(index.get(data[position : position + 3], ())[-MAX_CANDIDATES:]):
                length = 0
                # Overlapping matches are legal: the source may run into the region
                # being written, which repeats the pattern.
                while position + length < limit and data[candidate + length] == data[position + length]:
                    length += 1
                if length > best_length:
                    best_offset, best_length = position - candidate, length
                    if position + best_length >= limit:
                        break

        if best_length >= 3:
            bits = _bit_count(position - start)
            token = ((best_offset - 1) << (16 - bits)) | (best_length - 3)
            tokens += token.to_bytes(2, "little")
            flags |= 1 << flag_count
            step = best_length
        else:
            tokens.append(data[position])
            step = 1

        for offset in range(step):
            here = position + offset
            if here + 3 <= end:
                index.setdefault(data[here : here + 3], []).append(here)
        position += step

        flag_count += 1
        if flag_count == 8:
            out.append(flags)
            out += tokens
            tokens.clear()
            flags = flag_count = 0

    if flag_count:
        out.append(flags)
        out += tokens
    return bytes(out)


def compress(data: bytes) -> bytes:
    """Compress ``data`` into an MS-OVBA CompressedContainer."""
    out = bytearray(b"\x01")  # SignatureByte
    position = 0
    while position < len(data):
        end = min(position + CHUNK, len(data))
        body = _compress_chunk(data, position, end)

        if len(body) < CHUNK:
            header = 0xB000 | (len(body) - 1)
        elif end - position == CHUNK:
            # A full chunk that did not compress is stored verbatim, exactly 4096 bytes.
            body = data[position:end]
            header = 0x3000 | (CHUNK - 1)
        else:
            # A short final chunk may not use the raw form: that always emits 4096
            # bytes and would append padding to the module. Shrink it instead and
            # let the remainder start a new chunk.
            end = position + SAFE_PARTIAL
            body = _compress_chunk(data, position, end)
            header = 0xB000 | (len(body) - 1)

        out += header.to_bytes(2, "little") + body
        position = end
    return bytes(out)
