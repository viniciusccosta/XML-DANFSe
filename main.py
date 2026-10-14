import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from xml_danfse.application import run

if __name__ == "__main__":
    run(save_htmlfile=True)
