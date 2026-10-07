# -*- coding: utf-8 -*-
"""PyInstaller entry point for the XP bundle.

PyInstaller runs the entry script as a top-level module (__name__ == "__main__" with no
parent package), so a package-internal module such as ort_xp/__main__.py cannot use its
relative imports. This thin launcher lives OUTSIDE the package and imports it normally.

ASCII-only on purpose: PyInstaller 3.3.1 is encoding-picky on this toolchain.

Any unhandled exception is written to <app>\Logs\fatal_*.log and shown in a native
message box (the frozen app is a windowed subsystem, so there is no console).

`python -m ort_xp` keeps working through ort_xp/__main__.py; this file is only for the
frozen bundle.
"""

import sys


def _run():
    from ort_xp.__main__ import main

    return main()


if __name__ == "__main__":
    try:
        sys.exit(_run())
    except SystemExit:
        raise
    except Exception:
        from ort_xp import fatal

        sys.exit(fatal.handle(*sys.exc_info()))
