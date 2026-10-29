"""Headless command line parser (argparse as-is, P3 decision D4).

Argument names, defaults, help texts and usage are frozen (C3, ADR 0003 section 6).
Unknown arguments (e.g. --auto-backup) are ignored via parse_known_args.
"""
import argparse


def build_parser():
    """Return the Headless ArgumentParser (prog defaults to sys.argv[0] at call time)."""
    parser = argparse.ArgumentParser()
    parser.add_argument("--schedule-config", help="Path to schedule configuration file")
    parser.add_argument("--user", help="User name to back up")
    parser.add_argument("--dest", help="Destination folder for backup")
    parser.add_argument("--include-hidden", action="store_true")
    parser.add_argument("--include-system", action="store_true")
    parser.add_argument("--retention", type=int, default=2, help="Number of backups to keep")
    parser.add_argument("--log-retention", type=int, default=30, help="Number of days to keep logs")
    parser.add_argument("--log-folder", default="logs")
    parser.add_argument("--filters", type=str, help="JSON-encoded filter configuration")
    return parser


def parse(argv=None):
    """Parse argv (default sys.argv[1:]) and return the namespace; unknown args are ignored.

    Invalid values (e.g. --retention abc) print usage to stderr and raise SystemExit(2).
    """
    args, _unknown = build_parser().parse_known_args(argv)
    return args
