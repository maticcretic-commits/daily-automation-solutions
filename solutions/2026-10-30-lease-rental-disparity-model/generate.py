#!/usr/bin/env python3
"""Generate deterministic synthetic lease data for the LeaseDisparity demo.

Fictional portfolio: 12 commercial leases in the Los Angeles area, built
around a Freelancer brief asking for an Excel workbook comparing contracted
rent with current market rent for up to 12 leases and valuing the disparity
over each remaining term at a single yield rate. All names, rents and terms
are invented for the demo; no real client data. Rents are USD per month.
Seed 20261030 (data are an explicit fixed list, so the seed is documentary).
"""
import csv
import os

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")

# LeaseID, Property, Tenant, Area, ContractRentMonthly, MarketRentMonthly, RemainingMonths
LEASES = [
    ("L01", "Harbor View Retail", "Marina Books LLC", "Long Beach", 4200, 5800, 48),
    ("L02", "Arts District Loft Office", "Studio North Inc", "Los Angeles", 6800, 7400, 36),
    ("L03", "Valley Warehouse Bay 4", "Sunset Logistics Co", "Van Nuys", 9500, 8200, 24),
    ("L04", "Pasadena Cafe Corner", "Rose City Roasters", "Pasadena", 3100, 3900, 18),
    ("L05", "Downtown Flex Suite 210", "Beacon Consulting", "Los Angeles", 5400, 5400, 30),
    ("L06", "Burbank Sound Stage Office", "Echo Post LLC", "Burbank", 11200, 12900, 60),
    ("L07", "Santa Monica Clinic Suite", "Pacific Wellness Group", "Santa Monica", 7600, 7100, 12),
    ("L08", "Glendale Showroom", "Foothill Auto Gallery", "Glendale", 8900, 9800, 42),
    ("L09", "Torrance Light Industrial", "Harbor Parts Supply", "Torrance", 6300, 5900, 9),
    ("L10", "Hollywood Production Office", "Silverline Media", "Hollywood", 4800, 6100, 27),
    ("L11", "Culver City Creative Studio", "Framehouse Films", "Culver City", 5700, 6300, 15),
    ("L12", "Inglewood Retail Pad", "Forum Foods Inc", "Inglewood", 2600, 3400, 54),
]


def main():
    os.makedirs(DATA, exist_ok=True)
    with open(os.path.join(DATA, "leases.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["LeaseID", "Property", "Tenant", "Area",
                    "ContractRentMonthly", "MarketRentMonthly", "RemainingMonths"])
        w.writerows(LEASES)
    print(f"leases={len(LEASES)}")


if __name__ == "__main__":
    main()
