# Editable Visio Recreation

Recreate diagram screenshots as editable Microsoft Visio objects with an AI coding agent, then inspect the resulting file with repeatable validation tools.

[中文说明](README.zh-CN.md) · [Skill instructions](SKILL.md) · [Verification guide](references/verification.md) · [Demo](examples/README.md)

## What this project does

This repository combines an agent skill, implementation patterns, a Python VSDX static checker, and a Windows Visio exporter. It helps an agent rebuild text, blocks, lines, heatmap cells, and labels as separate objects rather than a single embedded screenshot.

**It is an agent-assisted workflow, not a standalone one-command image converter.** The agent must inspect the source, reconstruct its geometry, and create the VSDX. The bundled demo creates a synthetic diagram to demonstrate editable output; it does not perform general image recognition.

## Requirements

| Component | Requirement |
| --- | --- |
| Skill instructions | A coding agent that supports SKILL.md |
| Static checker and tests | Python 3.10+; standard library only |
| Actual-file PNG export | Windows, desktop Microsoft Visio, PowerShell, and an interactive login session |
| Optional raster sampling | An image library such as Pillow, only if the chosen reconstruction uses it |
| Fonts | Fonts used by the diagram must be installed; substitution can change appearance |

Python checks can run on Windows, Linux, or macOS. The included COM exporter and demo require Windows and desktop Visio. This project does not supply or license Visio or fonts.

## Install as a Codex skill

Download or clone the repository, open a terminal in its root, and run:

```powershell
python scripts/install_skill.py
```

The installer copies the skill to `$CODEX_HOME/skills/editable-visio-recreation`, or `~/.codex/skills/editable-visio-recreation` when CODEX_HOME is unset. It refuses to overwrite an existing installation. Reopen the agent session if the skill is not discovered.

For another agent, copy SKILL.md and LICENSE together with agents/, scripts/, and references/ into that agent's supported skill directory. Other agents have not been compatibility-tested.

Example prompt:

> Use $editable-visio-recreation to recreate the attached diagram as an editable .vsdx. Keep labels and shapes editable, preserve the source, and deliver an actual Visio-exported preview plus separate static and interaction-check results.

## Check an existing VSDX

Run from the repository root:

```powershell
python scripts/validate_vsdx.py diagram.vsdx --profile source --report check.json
```

The default `source` profile checks the package without imposing a page size or font preset. For a small journal diagram, opt into `--profile journal-small`: maximum 80 × 60 mm, 6 pt, SimSun for Chinese, and Times New Roman for Western text. Custom JSON profiles are supported.

The checker inspects ZIP integrity, XML, relationship targets, pages, shapes, connections, selected cached typography values, and embedded media. It reports fixed-geometry and inheritance warnings. `--strict` makes warnings fail the command.

**A successful static check does not certify that Visio opens the file, the drawing matches the source, or every object edits correctly.** The checker is not a complete OOXML schema validator and does not fully resolve masters, styles, formulas, fonts, or media editability. `--allow-media` relaxes media rejection; it does not prove useful editable content.

## Export the actual file

On Windows with desktop Visio:

```powershell
.\scripts\render_visio.ps1 -InputPath .\diagram.vsdx -OutputDirectory .\new-preview-folder -Dpi 300
```

The output directory must not already exist. The script uses a private invisible Visio instance, opens the file read-only with macros/events/data refresh disabled, exports every page, and records hashes in render.json. Review the PNGs yourself; opening and exporting alone do not establish visual fidelity.

Reports can contain local paths. Review and redact them before publishing reports from private projects.

## Reproducible example

See [examples/README.md](examples/README.md) for a synthetic native-shape diagram, an actual-file PNG export, and representative resize and connector-following checks.

![Synthetic demo exported from Microsoft Visio](examples/render/page-001_preview.png)

## Tests

```powershell
python -m unittest discover -s scripts -p "test_*.py" -v
```

The validator's regression fixtures are synthetic OPC packages. They exercise implemented static checks and are not evidence of application-valid Visio documents. The separate demo provides application evidence when generated on a machine with Visio.

The CI configuration runs static regression tests on Windows and Linux. Hosted CI does not install or run desktop Visio.

## Repository contents

- SKILL.md: agent workflow and delivery requirements.
- scripts/validate_vsdx.py: dependency-free static inspector.
- scripts/render_visio.ps1: actual-file exporter using desktop Visio.
- scripts/create_demo.ps1: reproducible editable-output example.
- scripts/install_skill.py: non-overwriting skill installer.
- references/: implementation patterns, profiles, and verification limits.
- examples/: synthetic example and its verification evidence.

## Contributing

Useful contributions include a minimal reproducible VSDX, a failing checker test, clearer reconstruction patterns, or a diagram example that you own and can redistribute. Describe whether a report concerns static validation, rendering, or interactive editing. Avoid uploading private diagrams, customer files, credentials, or licensed source images.

## License

[MIT](LICENSE). Example code and synthetic diagrams created for this repository are covered by the same license. Microsoft Visio and third-party fonts remain subject to their own licenses.

