---
name: editable-visio-recreation
description: Recreate diagram images, scientific figures, heatmaps, architecture charts, topology plots, and other raster references as editable Microsoft Visio .vsdx files. Use when Codex is asked to convert a PNG/JPG/screenshot/PDF figure into an editable Visio version, preserve visual layout, rebuild text/shapes/arrows/color scales as separate Visio objects, generate a preview image, or verify that the output is vector/editable rather than an embedded bitmap.
---

# Editable Visio Recreation

## Goal

Rebuild the visible diagram as editable Visio objects. Prefer separate rectangles, lines, arrows, text blocks, circles, polygons, heatmap cells, colorbar segments, and labels over embedding the source image.

Deliver both:

- `.vsdx` editable Visio file
- `_preview.png` exported from the actual VSDX for visual QA; a separately drawn mockup is not proof of the file's appearance

## Workflow

1. Inspect the workspace for the source image and related generators.
   - Use `Get-ChildItem` or `rg --files` when available.
   - Open the source with `view_image` before coding.
   - Get source dimensions so the Visio page matches the raster aspect ratio.

2. Classify the figure.
   - Heatmap: grid cells, sampled colors, axis labels, colorbar, optional cell text.
   - Architecture/block diagram: boxes, rounded rectangles, arrows, dashed groups, plus nodes, labels.
   - Scatter/topology/map: axes, points, markers, legends, paths, labels.
   - Mixed figure: split into visible regions and rebuild each region with simple primitives.

3. Choose generation strategy.
   - Direct XML is suitable for primitive-heavy figures when a known-good package template is available. Keep relationships, content types, page entries, and masters consistent; changing two XML files alone is not a complete package strategy.
   - Prefer native Visio connectors through COM for diagrams whose boxes will be moved. Drawn arrow polygons are editable but do not automatically follow a box.
   - COM requires desktop Visio and an interactive Windows session. Use a private invisible instance, not an existing user document; do not enable macros or refresh external data.

4. Rebuild primitives.
   - Use page units derived from pixels, commonly `SCALE = 90.0` pixels per inch.
   - Preserve source coordinates as constants: plot bounds, colorbar bounds, group box bounds, node coordinates.
   - Keep all text editable. Use text inside shapes for cell values or labels; use transparent rectangles for standalone text.
   - Use actual shape fills and strokes. Do not use the original image as a background unless the user explicitly requests a trace layer.

5. Generate an actual-file preview.
   - Run `scripts/render_visio.ps1` against the output VSDX using a new output directory. It exports every page via desktop Visio and records the source SHA256 in `render.json`.
   - Open all exported pages with `view_image` and compare against the source. Check text clipping, missing glyphs, line attachment, color, and page edges.
   - PIL may help estimate layout, but label its output as a mockup, not a Visio render. If Visio is unavailable, state that opening/rendering remains unverified.

6. Validate editability.
   - Run `scripts/validate_vsdx.py` for ZIP, XML, relationship targets, every page, text, geometry warnings, and media checks. Pass the selected profile. `static_ok` is not an application-opening certificate.
   - For a new generator or changed geometry strategy, open a disposable copy in Visio, resize a representative block and move a connected block. Save/flush recalculation before checking the visible geometry bounds and connector endpoints. Never perform this destructive test on the only deliverable.
   - Record static results, actual render results, and interaction-test results separately. A large shape count or absence of images alone does not prove useful editability.
   - Report any intentional approximations, such as unreadable source labels or missing labels in the source image.

For coordinate conversion, heatmap sampling, and native connector details, read `references/implementation-patterns.md`. Commands and profile selection are in `references/verification.md`.

## Figure-Specific Guidance

### Heatmaps

- Detect the actual visible grid dimensions. Do not assume title text like "24 nodes" guarantees 24 visible columns.
- Preserve source-visible labels. If an axis label is missing in the raster, do not silently invent it; note the source discrepancy.
- Sample each cell color robustly by taking median RGB values from the cell interior and ignoring near-white grid lines and near-black text.
- For annotated heatmaps, store values in a `DATA` matrix and render each value as editable text.
- Use readable text color based on fill luminance; white text on dark red/blue cells and black text elsewhere.
- Sample the colorbar into many thin editable rectangles, then add editable tick labels.

### Architecture Diagrams

- Recreate dashed group boxes, filled rounded blocks, arrows, plus nodes, and internal flow separately.
- Use transparent text boxes for titles and section labels.
- Use polygons for diagonal arrows and arrowheads when direct Visio connectors are unnecessary.
- Keep mathematical labels editable as plain text. Preserve superscripts in a simple readable form if true rich text is not necessary.

### Scatter/Topology Figures

- Reconstruct axes, tick labels, legends, markers, and highlighted points as separate objects.
- When points are numerous, extract or sample points from the image but keep final points as individual editable shapes.
- Use stable plot bounds and data-to-pixel conversions so axis labels and markers align.

## Quality Bar

- The output opens as a valid `.vsdx`.
- The preview matches the source layout closely enough for paper/report editing.
- Text is readable and editable.
- Repeated items are built as repeated shapes, not one screenshot.
- The final response links the `.vsdx`, preview image, and any generator script.
- Mention verification results: static checks, actual Visio export, and representative resize/connector checks (or clearly state what could not be verified).

## Local Lessons From This Project

- Some environments may not have `git`, and `rg` may be blocked. Fall back to PowerShell file listing and `Select-String`.
- Direct XML is deterministic, but package consistency and actual application opening still need checking.
- Existing generated `.vsdx` files can serve as templates for package structure.
- PIL mockups can catch initial alignment errors but can disagree with Visio's text layout and geometry.
- Source images may contain inconsistencies. Preserve visible truth and call out discrepancies rather than normalizing them silently.
- For user-editable Visio output, shape geometry must scale when the user resizes the visible selection box. Prefer `RelMoveTo` / `RelLineTo` rows for rectangles and blocks, or use `Width` / `Height` formulas on geometry cells. Do not use fixed absolute `MoveTo` / `LineTo` geometry without formulas for resizable shapes.
- Lines and arrows should use Visio-style `BeginX` / `BeginY` / `EndX` / `EndY` cells with `PinX=(BeginX+EndX)/2`, `PinY=(BeginY+EndY)/2`, `Width=SQRT((EndX-BeginX)^2+(EndY-BeginY)^2)`, and `Angle=ATAN2(EndY-BeginY,EndX-BeginX)` so line resizing behaves like native Visio geometry.
- If the current workspace has `visio_reference_standard/`, inspect the user's reference and README. Do not assume that folder exists in other projects.
- Use the `source` profile by default; it does not impose a page size or font. The optional `journal-small` profile in `references/profiles.json` uses a maximum 80 by 60 mm page, 6 pt text, SimSun for Chinese, and Times New Roman for Western text. Explicit user/template requirements override presets. For other numerical constraints, use a task-local profiles JSON rather than changing installed defaults. Split mixed Chinese/Western labels when rich text runs are not implemented.
- Preserve the original source image. Create a clearly named task output folder for the generated `.vsdx`, preview PNG, generator, and verification reports. Do not move or overwrite source files unless the user asks.
