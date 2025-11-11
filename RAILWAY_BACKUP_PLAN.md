# Railway Deployment - Backup Plan

## If Cairo/SVG Dependencies Keep Failing

**Current Problem:** `pycairo` requires Cairo graphics library, which is hard to install in Railway's environment.

**Why It's Needed:** SVG logo conversion (`Logos/*.svg` → PNG for PowerPoint)

---

### Option A: Remove SVG Dependency (FASTEST)

**Step 1:** Remove SVG dependencies from `requirements.txt`

```diff
- svglib>=1.6.0
- reportlab>=4.0.0
```

**Step 2:** Update `src/pptx_generator.py`

Remove or comment out the SVG conversion function (line ~1308):
```python
def _convert_svg_to_png(svg_path: str, width_px: int = 200, height_px: int = 200) -> str:
    # Disabled for Railway deployment
    return None
```

**Step 3:** Update logo matching logic in `src/ai_selector.py` (line ~204):

```python
# Skip SVG logos on Railway
if matched_logo and not matched_logo.endswith('.svg'):
    placeholders[logo_key] = f"Logos/{matched_logo}"
else:
    placeholders[logo_key] = ""  # Skip logos
```

**Impact:** Presentations work but without logos (not critical for core functionality)

---

### Option B: Pre-Convert All SVGs to PNGs

**Step 1:** Locally convert all SVGs:
```powershell
# Run locally (requires Inkscape or similar)
cd Logos/
foreach ($svg in Get-ChildItem *.svg) {
    $png = $svg.Name -replace '\.svg$', '.png'
    inkscape --export-filename=$png --export-width=200 $svg
}
```

**Step 2:** Update code to use PNG files:
```python
# In src/ai_selector.py, change:
placeholders[logo_key] = f"Logos/{matched_logo}.png"  # Force PNG
```

**Step 3:** Remove SVG dependencies from `requirements.txt`

**Impact:** Logos work perfectly, no Cairo needed

---

### Option C: Use Pre-Built Wheel (ADVANCED)

Force pip to use a pre-compiled wheel for pycairo:

```bash
pip install --only-binary :all: pycairo
```

Add to `nixpacks.toml`:
```toml
[phases.install]
cmds = ["pip install --only-binary :all: -r requirements.txt"]
```

**Note:** May not have wheels for all Python versions/platforms

---

## Recommendation

1. **Try current APT fix first** (pushed as `49df52a`)
2. If it fails again → **Option B** (pre-convert SVGs locally)
3. Last resort → **Option A** (disable logos entirely)

Railway build logs will tell you if APT packages worked.
