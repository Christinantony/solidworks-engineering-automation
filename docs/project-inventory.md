# Source inventory

All 16 supplied macro containers were inspected. No byte-identical macro containers were found. Similar filenames identify development families, not proven chronological versions.

| Original macro | Published location | Status | Source evidence |
| :--- | :--- | :--- | :--- |
| `Assembly Batch rename.swp` | [src/assembly-renaming](../src/assembly-renaming) | Implementation; runtime unverified | Tree traversal, unique-part filtering, preview/conflict checks, RenameDocument events, drawing Save As |
| `Batch Rename.swp` | [experiments/renaming/batch-save-as](../experiments/renaming/batch-save-as) | Earlier alternative | Per-component Save As numbering; distinct approach, not a byte duplicate |
| `Create Drawing.swp` | [experiments/drawings/create-drawing](../experiments/drawings/create-drawing) | Recorded prototype | Create a drawing, apply sheet format, insert one front view |
| `Create holes cd.swp` | [experiments/hole-wizard/create-holes-cd](../experiments/hole-wizard/create-holes-cd) | Recorded alternative | Different HoleWizard5 parameters |
| `Create holes from sketch points.swp` | [src/holes-from-sketch-points](../src/holes-from-sketch-points) | Interactive prototype | Collect selected sketch points and dimensions; map drill code; call HoleWizard5 |
| `Create holes new.swp` | [experiments/hole-wizard/create-holes-new](../experiments/hole-wizard/create-holes-new) | Recorded alternative | Different HoleWizard5 parameters |
| `Create holes.swp` | [experiments/hole-wizard/create-holes](../experiments/hole-wizard/create-holes) | Recorded experiment | Hard-coded HoleWizard5 and sketch-point calls |
| `Drawing template change.swp` | [experiments/drawings/template-stub](../experiments/drawings/template-stub) | Stub | Only connects to Application.SldWorks |
| `Forming Chamber Hole.swp` | [src/excel-point-importer](../src/excel-point-importer) | Implementation; runtime unverified | Excel first-sheet X/Y input, inch-to-meter conversion, origin offsets, coordinate grouping, horizontal/vertical relations |
| `Name the circles drawn.swp` | [experiments/annotations/circle-label](../experiments/annotations/circle-label) | Prototype | Attempt sketch text near selected arc/circle |
| `Name the circles selected in Drawing.swp` | [experiments/annotations/drawing-label](../experiments/annotations/drawing-label) | Interactive prototype | Sequential notes at selected entity positions; unbounded loop |
| `Name the circles selected.swp` | [experiments/annotations/selected-sketch-label](../experiments/annotations/selected-sketch-label) | Interactive prototype | Sequential sketch-text labels; unbounded loop |
| `Rename Parts.swp` | [experiments/renaming/single-part](../experiments/renaming/single-part) | Prototype | Single-part suffix rename and reference event handler |
| `Test Macro for hole.swp` | [experiments/annotations/note-sequence](../experiments/annotations/note-sequence) | Test utility | Sequential note placement; does not create holes |
| `test macro for rename.swp` | [experiments/renaming/recorded-selection](../experiments/renaming/recorded-selection) | Test recording | Hard-coded component selection |
| `Test Macro.swp` | [experiments/annotations/recorded-note](../experiments/annotations/recorded-note) | Test recording | Recorded selection and note operations |
