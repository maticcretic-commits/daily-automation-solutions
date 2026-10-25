#!/usr/bin/env python3
"""Run the full build: synthetic data -> workbook."""
import generate
import build_workbook

if __name__ == "__main__":
    generate.main()
    build_workbook.build()
