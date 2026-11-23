"""PyInstaller entry script (absolute import so the package resolves when frozen)."""
import sys

from notice_generator.main import main

if __name__ == "__main__":
    sys.exit(main())
