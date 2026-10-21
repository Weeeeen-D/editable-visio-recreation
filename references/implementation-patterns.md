# Implementation Patterns

## Direct `.vsdx` Package Pattern

Use a known-good `.vsdx` as a template:

1. Open template and output with `zipfile`.
2. Copy only the parts needed for the output. Remove obsolete relationships/content-type entries when deleting thumbnails, media, or pages.
3. Update page XML, page index, content types, and relationship targets together. Preserve referenced masters and styles, or remove their references when replacing them with self-contained primitives.
4. Derive page dimensions from the selected preset and source aspect ratio. Use one uniform scale factor; do not distort the source to fill the maximum page size.

Use the reusable static validator and actual-file exporter described in [verification.md](verification.md). ZIP integrity alone is insufficient.

## Coordinate Helpers

Use pixel-first constants and convert to Visio page units:

```python
SCALE = 90.0

def px_to_page(px, py, image_h):
    return px / SCALE, (image_h - py) / SCALE

def px_rect_to_page(box, image_h):
    x1, y1, x2, y2 = box
    return x1 / SCALE, (image_h - y2) / SCALE, x2 / SCALE, (image_h - y1) / SCALE
```

Build every Visio element from the same pixel constants used by the preview.

## Shape Builder Pattern

Maintain a small `ShapeBuilder` with monotonically increasing IDs. Implement primitives:

- `rect(...)` for rectangles, rounded blocks, transparent text boxes, heatmap cells
- `polygon(...)` for arrows, arrowheads, stars, approximate circles
- thin filled rectangles or polygons for lines

Important cells for rectangles:

- `PinX`, `PinY`, `Width`, `Height`, `LocPinX`, `LocPinY`
- `FillPattern`, `FillForegnd`
- `LinePattern`, `LineColor`, `LineWeight`
- margins and `VerticalAlign`
- character and paragraph sections when text is present

## Heatmap Sampling

Sample cell colors from the source image interior:

```python
import statistics

def robust_cell_color(image, box):
    image = image.convert("RGB")
    pixels = image.load()
    x1, y1, x2, y2 = box
    samples = []
    for y in range(round(y1 + 5), round(y2 - 5)):
        for x in range(round(x1 + 5), round(x2 - 5)):
            r, g, b = pixels[x, y]
            if r > 246 and g > 246 and b > 246:
                continue
            if r < 55 and g < 55 and b < 55:
                continue
            samples.append((r, g, b))
    if not samples:
        return pixels[round((x1 + x2) / 2), round((y1 + y2) / 2)]
    ranked = sorted(samples, key=lambda c: c[0] + c[1] + c[2])
    trim = len(ranked) // 10
    trimmed = ranked[trim:len(ranked) - trim] if trim else ranked
    return tuple(round(statistics.median(ch)) for ch in zip(*trimmed))
```

For annotated heatmaps, keep the numeric `DATA` matrix explicit so labels are editable and auditable.

## Preview Pattern

For a fast preliminary layout mockup only, draw from the same constants and data. Give it a `_mockup.png` suffix. The delivered `_preview.png` must be an actual-file render:

- Use `Image.new("RGB", source.size, WHITE)`.
- Draw cells/blocks/arrows/text with PIL.
- Use Windows fonts when available:
  - `C:\Windows\Fonts\simhei.ttf`
  - `C:\Windows\Fonts\simsun.ttc`
  - `C:\Windows\Fonts\msyh.ttc`
  - `C:\Windows\Fonts\times.ttf`
  - `C:\Windows\Fonts\arial.ttf`
- Open the preview with `view_image` before final response.

## Reporting

In the final response, include:

- Absolute links to `.vsdx` and `_preview.png`.
- Link to the generator script if created.
- Validation summary separating static checks, actual-file rendering, and representative interaction tests. Do not treat a mockup as an application render.
- Any source-image discrepancies or intentional approximations.

## Native connectors and interaction tests

With a private Visio COM instance, drop `app.ConnectorToolDataObject` onto the page, then glue the connector's `BeginX` and `EndX` cells to the source and destination `PinX` cells. This creates dynamic attachment. Merely setting endpoint coordinates does not create glue.

On a disposable test copy, record connector endpoint coordinates, move the target block, save to flush deferred routing, then compare. For a rectangle, record `BoundingBox(4, ...)`, change `Width`, and verify the visible geometry bounds change, not only the selection box. Reopen/render the final saved output separately. Never change the real deliverable during these tests.

Primary API references: [GlueTo](https://learn.microsoft.com/en-us/office/vba/api/visio.cell.glueto), [BoundingBox](https://learn.microsoft.com/en-us/office/vba/api/visio.shape.boundingbox), [OpenEx](https://learn.microsoft.com/en-us/office/vba/api/visio.documents.openex), [Page.Export](https://learn.microsoft.com/en-us/office/vba/api/visio.page.export).
