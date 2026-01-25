# -*- coding: utf-8 -*-
"""Lightweight logging helpers (stderr)."""

import sys


def warn(msg: str) -> None:
    print(f"[WARN] {msg}", file=sys.stderr)


def info(msg: str) -> None:
    print(f"[INFO] {msg}", file=sys.stderr)
