"""
Modal endpoint: build a Program Map xlsx from a posted JSON bundle.

Deploys the existing pdow_kit/ Python code verbatim as a Modal FastAPI
endpoint. The Render proxy POSTs the JSON bundle returned by
/api/pdow-data to this endpoint and streams back the xlsx bytes.

Deploy:
    cd modal_builder
    modal deploy app.py

First deploy prompts for Modal auth (free account, https://modal.com).
After success, Modal prints the endpoint URL — copy it into Render as
the MODAL_BUILDER_URL environment variable. See README.md.
"""
from __future__ import annotations
from pathlib import Path
import modal

# The kit lives one directory up so Modal and the in-repo Python runner
# share the exact same merger / writer / template / image.
KIT_DIR = Path(__file__).resolve().parent.parent / "pdow_kit"

app = modal.App("pdow-builder")

image = (
    modal.Image.debian_slim(python_version="3.11")
    .pip_install("openpyxl==3.1.5", "Pillow==12.2.0", "fastapi[standard]")
    # Mount the whole kit (merger.py, xlsx_writer.py, template_BSSCOM.xlsx,
    # owl.png, etc.) under /kit in the container. data.json in the kit is
    # ignored here — we feed a fresh bundle via the POST body.
    .add_local_dir(str(KIT_DIR), "/kit", copy=True)
)


@app.function(image=image, timeout=120)
@modal.fastapi_endpoint(method="POST", docs=False)
def build(data: dict):
    """POST JSON → xlsx bytes.

    Body shape: the `data.json` the kit consumes —
      { program_code, model, alignments }.
    Response: xlsx binary with a Content-Disposition filename matching
    <PROGRAM>_Program_Map.xlsx.
    """
    import sys
    import tempfile
    from fastapi.responses import Response

    sys.path.insert(0, "/kit")
    from merger import merge                          # noqa: E402
    from xlsx_writer import write_program_map         # noqa: E402

    program_code = (data.get("program_code") or "PROGRAM").strip() or "PROGRAM"
    model = merge(data)

    # Write to a tmpfile inside the container (ephemeral, fine for one-shot
    # builds). The writer's LibreOffice-based formula-recalc step is a no-op
    # when LibreOffice isn't present — Excel recalculates on open, which is
    # what end users do anyway.
    with tempfile.TemporaryDirectory() as tmpdir:
        out_path = Path(tmpdir) / f"{program_code}_Program_Map.xlsx"
        write_program_map(
            model=model,
            template_path=Path("/kit/template_BSSCOM.xlsx"),
            output_path=out_path,
            owl_image_path=Path("/kit/owl.png"),
            tab_name="Program Map",
            program_code=program_code,
        )
        xlsx_bytes = out_path.read_bytes()

    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{program_code}_Program_Map.xlsx"',
        },
    )
