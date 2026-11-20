#!/usr/bin/env python3
"""Build the DataOps Automation System end to end.

1. make_data.py  -> data/*.csv (deterministic fictional source data)
2. pipeline.py   -> SQLite aggregates, star-schema exports, workbook
"""
import make_data
import pipeline

if __name__ == "__main__":
    make_data.make_data()
    pipeline.run_pipeline()
    print("Built dataops-automation-system.xlsx")
