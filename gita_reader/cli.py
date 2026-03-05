#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from gita_reader.pipeline import TodoBuilder, write_outputs


def main() -> None:
    parser = argparse.ArgumentParser(description="Build personal to-do lists from competition workbook")
    parser.add_argument("xlsx", type=Path, help="Path to source .xlsx file")
    parser.add_argument("--out", type=Path, default=Path("output"), help="Output directory")
    args = parser.parse_args()

    builder = TodoBuilder(args.xlsx)
    data = builder.build()
    write_outputs(data, args.out)

    people_count = len(data.get("people", {}))
    board_tasks = len(data.get("shared_tasks", {}).get("board_tasks", []))
    liaison_tasks = len(data.get("shared_tasks", {}).get("liaison_tasks", []))
    print(f"Generated to-dos for {people_count} people")
    print(f"Shared board tasks: {board_tasks}")
    print(f"Shared liaison tasks: {liaison_tasks}")
    print(f"Output directory: {args.out}")


if __name__ == "__main__":
    main()
