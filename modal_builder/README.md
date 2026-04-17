# Modal builder — PDOW xlsx endpoint

Deploys the existing `pdow_kit/` Python code as a Modal FastAPI endpoint so the browser can get a finished xlsx directly (no "download zip and run Python" for end users). The Render Node proxy forwards `/api/pdow-xlsx` requests here; the browser never talks to Modal directly.

## Why Modal

- The exact Python kit runs as-is — no rewrite, no WASM, no Pyodide pageweight.
- Single deploy command; no Dockerfile.
- Free tier ($30/mo credit) easily covers thousands of xlsx builds.
- Sub-second warm, ~2–4 s cold start.

## One-time setup

```bash
# 1. Install Modal CLI and authenticate (opens browser to create/link account)
pip install modal
modal token new

# 2. Deploy this app
cd modal_builder
modal deploy app.py
```

On a successful deploy, Modal prints a line like:

```
✓ Created web endpoint for build => https://yourorg--pdow-builder-build.modal.run
```

That URL is the **builder endpoint**. Copy it.

## Wire it into Render

The Render Node proxy reads the builder URL from an env var. In the Render dashboard:

1. Open the `coda-exporter` service → **Environment** tab.
2. Add a new environment variable:
   - **Key:** `MODAL_BUILDER_URL`
   - **Value:** the full URL from `modal deploy` (e.g. `https://yourorg--pdow-builder-build.modal.run`).
3. Save. Render auto-redeploys.

From that point `/api/pdow-xlsx?programAbbr=MSCSIA` on the Render proxy will:
1. Fetch the JSON bundle from Coda (same logic as `/api/pdow-data`).
2. POST the bundle to this Modal endpoint.
3. Stream the returned xlsx back to the browser.

## Redeploying the Modal app

Any time the kit changes (`pdow_kit/xlsx_writer.py`, `merger.py`, template, etc.), re-run:

```bash
cd modal_builder
modal deploy app.py
```

The URL stays the same across deploys, so no Render env update is needed.

## Notes

- The app wires `add_local_dir(KIT_DIR, "/kit", copy=True)` at image-build time, so deploys need the `pdow_kit/` directory present locally. Run `modal deploy` from a checkout of this repo.
- The openpyxl / Pillow versions in `app.py` track what `pdow_kit/README.md` specifies — bump together if you upgrade one.
- Timeout is 120s, which is well above observed build times (~2s per program).
