"""
MSCIN workbook emitter.

Loads mscin_model.json (built by build_mscin_model.py) and emits the full
PD Outputs workbook with:
  - Program Map (with PO IRMA mapping populated; CCT columns present but blank)
  - Course Alignment
  - Standard Path
  - Comps to Skills  (NEW)
  - Skills to Comps  (NEW)
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

# Make local module imports work when run from anywhere
sys.path.insert(0, str(Path(__file__).parent))

from xlsx_writer import write_program_map  # noqa: E402


def main():
    here = Path(__file__).parent
    model_path = here / "data" / "live" / "mscin_model.json"
    template = here / "template_BSSCOM.xlsx"
    owl = here / "owl.png"
    output_dir = here.parent / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / "MSCIN_PD_Outputs_Workbook.xlsx"

    if not model_path.exists():
        sys.exit(f"No model at {model_path}. Run build_mscin_model.py first.")

    with model_path.open() as f:
        model = json.load(f)

    result = write_program_map(
        model=model,
        template_path=template,
        output_path=output,
        owl_image_path=owl,
        program_code="MSCIN",
    )
    print(f"Wrote {result}")


if __name__ == "__main__":
    main()
