"""
RUN BOTH NSE SCANNERS
---------------------
1) NSE_Live_V2_V3_Scanner.py
   - Live daily V2 + V3 stock-selection scanner.

2) nse_intraday_picks_pro.py
   - NSE Intraday PRO stock-picks generator.

Both source files stay separate. This file simply runs them one after another.
"""

import runpy
from pathlib import Path

BASE = Path(__file__).resolve().parent

LIVE_V2_V3 = BASE / "NSE_Live_V2_V3_Scanner.py"
INTRADAY_PRO = BASE / "nse_intraday_picks_pro.py"


def run_file(path: Path, label: str) -> None:
    print("\n" + "=" * 80)
    print(f"STARTING: {label}")
    print(f"FILE: {path.name}")
    print("=" * 80)

    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    runpy.run_path(str(path), run_name="__main__")

    print("\n" + "=" * 80)
    print(f"FINISHED: {label}")
    print("=" * 80)


def main() -> None:
    run_file(LIVE_V2_V3, "LIVE V2 + V3 SCANNER")
    run_file(INTRADAY_PRO, "NSE INTRADAY PRO")


if __name__ == "__main__":
    main()
