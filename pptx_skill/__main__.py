"""Allow ``python -m pptx_skill <command>`` to run the CLI."""
import sys

from pptx_skill.cli import main

if __name__ == "__main__":
    sys.exit(main())
