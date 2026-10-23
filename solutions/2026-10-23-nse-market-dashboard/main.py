#!/usr/bin/env python3
"""NSEPulse pipeline: generate synthetic snapshots -> build workbook."""
import generate
import build_workbook


def main():
    generate.main()
    build_workbook.build()
    print("NSEPulse build complete.")


if __name__ == "__main__":
    main()
