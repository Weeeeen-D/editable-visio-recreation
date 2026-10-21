# Verification and presets

Resolve scripts relative to this skill's installed directory, not the project working directory. `validate_vsdx.py` uses Python 3's standard library only. `render_visio.ps1` requires Windows desktop Visio and an interactive login session.

```powershell
& '<python.exe>' '<skill-dir>\scripts\validate_vsdx.py' '<output.vsdx>' --profile source --report '<check.json>'
& '<skill-dir>\scripts\render_visio.ps1' -InputPath '<output.vsdx>' -OutputDirectory '<new-preview-folder>' -Dpi 300
```

- `journal-small` is an optional 80 by 60 mm maximum page, 6 pt, SimSun and Times New Roman preset. Preserve aspect ratio within that maximum.
- `source` disables the preset size/font checks when matching a different source or publisher template. It does not disable package checks.
- For custom numerical constraints, create a task-local JSON using the same keys as `profiles.json`, then pass `--profiles <path> --profile <name>`.
- `--allow-media` permits images/embedded parts only when the user explicitly accepts a mixed or trace-layer deliverable. Report which elements are not editable primitives.
- The static validator reads cached ShapeSheet values. It checks direct character formatting, but does not fully resolve masters, style inheritance, multi-run fonts, Visio formulas, or font substitutions. Warnings need review; `--strict` makes warnings fail the static check.
- `static_ok=true` means only the implemented static checks passed. The report always lists unverified application and interaction behavior.
- The renderer opens a private invisible Visio instance read-only, disables macros/events/external data refresh, exports every page, restores its previous raster resolution, and never saves the source. Compare `source_sha256` in `render.json` with the final VSDX hash before claiming it was rendered.
- Use a new output directory for each render. A failed export can leave an incomplete directory; it is not a successful verification unless `render.json` exists and all PNGs have been visually inspected.
- If COM reports that the login session does not exist, the execution sandbox may not have desktop access. Use the normal approval flow to rerun the bounded render in the user's session. Do not attach to or terminate the user's existing Visio process.
- Inspect exported PNGs for blank content, missing characters, clipped labels, overlapping elements, detached arrows and page edges. Report actual Visio export and visual review separately.

For new or changed generators, also perform the disposable-copy resize and connector-following test in [implementation-patterns.md](implementation-patterns.md). Representative tests provide evidence for those shapes, not every possible edit in every diagram.
