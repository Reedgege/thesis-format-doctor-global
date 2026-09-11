"""Thesis Format Doctor Global - packaging entry point.

Nuitka compiles this file into the executable stub; the real logic lives in
`src.main`. Keeping a thin root launcher lets `src` stay a normal package so
its intra-package relative imports (`from .engine import ...`) keep working
after compilation.
"""

import sys

from src.main import main

if __name__ == "__main__":
    sys.exit(main())
