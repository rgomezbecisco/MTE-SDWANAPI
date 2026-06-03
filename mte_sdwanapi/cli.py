"""CLI utilities for the SD-WAN report script."""

import argparse


def parse_args():
    """Return parsed command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Generate SD-WAN device and template Excel report."
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print collected API data during execution.",
    )
    return parser.parse_args()
