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
    parser.add_argument(
        "--find_default_feature_templates",
        action="store_true",
        help="Find factory default feature templates.",
    )
    return parser.parse_args()
