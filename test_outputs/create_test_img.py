"""Create a minimal test PNG image."""

import struct
import zlib


def create_png(path: str, width: int = 50, height: int = 30) -> None:
    def chunk(chunk_type: bytes, data: bytes) -> bytes:
        c = chunk_type + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

    ihdr_data = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    raw = b""
    for y in range(height):
        raw += b"\x00"  # filter none
        for x in range(width):
            raw += bytes([70, 130, 180])  # steel blue RGB

    idat_data = zlib.compress(raw)

    with open(path, "wb") as f:
        f.write(b"\x89PNG\r\n\x1a\n")
        f.write(chunk(b"IHDR", ihdr_data))
        f.write(chunk(b"IDAT", idat_data))
        f.write(chunk(b"IEND", b""))


create_png("/home/maruf/Development/Personal/AI/excel-mcp/test_outputs/test_img.png")
print("Created test_img.png")
