"""Launch the member collector restored from the shared-folder version."""
from pathlib import Path
import sys

if not getattr(sys, 'frozen', False):
    sys.path.insert(0, str(Path(__file__).resolve().parent / 'capture_reference'))

from desktop_app import launch

if __name__ == '__main__':
    raise SystemExit(launch('capture'))
