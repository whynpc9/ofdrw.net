#!/usr/bin/env python3
"""Generate the self-authored, deterministic 64x56 1 bpp PBM validation input."""

from pathlib import Path
import sys

WIDTH = 64
HEIGHT = 56


def black(x: int, y: int) -> bool:
    return ((x // 8 + y // 8) % 2 == 0) ^ (x == y)


def main() -> None:
    if len(sys.argv) != 2:
        raise SystemExit("usage: generate_owned_pattern.py OUTPUT.pbm")
    rows = bytes(
        sum(int(black(x, y)) << (7 - x % 8) for x in range(col, col + 8))
        for y in range(HEIGHT)
        for col in range(0, WIDTH, 8)
    )
    Path(sys.argv[1]).write_bytes(f"P4\n{WIDTH} {HEIGHT}\n".encode("ascii") + rows)


if __name__ == "__main__":
    main()
