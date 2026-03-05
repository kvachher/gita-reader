#!/usr/bin/env python3
from __future__ import annotations

import sys
import trace
import unittest
from pathlib import Path


def main() -> int:
    repo_root = Path(__file__).resolve().parents[1]
    tests_dir = repo_root / "tests"
    package_dir = repo_root / "gita_reader"
    import os

    os.chdir(repo_root)
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))

    loader = unittest.defaultTestLoader
    suite = loader.discover("tests", pattern="test_*.py")
    runner = unittest.TextTestRunner(verbosity=1)

    ignoredirs = [sys.prefix, sys.exec_prefix]
    tracer = trace.Trace(count=True, trace=False, ignoredirs=ignoredirs)
    test_result = tracer.runfunc(runner.run, suite)

    counts = tracer.results().counts
    files = sorted(package_dir.rglob("*.py"))

    print("\nCoverage (gita_reader):")
    print("  File                               Covered/Total   Percent")
    print("  ----------------------------------------------------------")

    total_covered = 0
    total_executable = 0
    for file_path in files:
        executable = trace._find_executable_linenos(str(file_path))
        if not executable:
            continue
        covered = sum(1 for line in executable if counts.get((str(file_path), line), 0) > 0)
        total = len(executable)
        percent = (covered / total * 100.0) if total else 100.0
        total_covered += covered
        total_executable += total
        rel = file_path.relative_to(repo_root).as_posix()
        print(f"  {rel:<34} {covered:>4}/{total:<7} {percent:>7.2f}%")

    overall = (total_covered / total_executable * 100.0) if total_executable else 100.0
    print("  ----------------------------------------------------------")
    print(f"  TOTAL                              {total_covered:>4}/{total_executable:<7} {overall:>7.2f}%")

    return 0 if test_result.wasSuccessful() else 1


if __name__ == "__main__":
    raise SystemExit(main())
