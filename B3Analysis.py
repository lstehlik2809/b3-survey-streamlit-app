# -*- coding: utf-8 -*-
"""Compatibility entry point for the local B3 report workflow."""

from src.b3_analysis import generate_default_report


if __name__ == "__main__":
    result = generate_default_report()
    print(f"Report saved to: {result.report_path}")
