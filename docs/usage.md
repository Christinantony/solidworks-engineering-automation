# Usage and engineering assumptions

## Assembly renaming

Open a saved, resolved top-level assembly. Run `Assembly_Batch_rename1.main`. The code targets part filenames ending in `_01`, skips suppressed/virtual/unloaded components, and prompts for a first number such as `HC-01010`. Numbering advances in traversal order. Inspect the temporary preview, resolve destination conflicts, and only then accept the macro's confirmation. Compare the final report and reopen affected drawings to verify references. Save failures do not constitute a rollback.

## Excel point importer

Use the first worksheet of an Excel workbook. Column A contains X and column B contains Y, both in **inches**. Start at row 1 with numeric data and no header. A blank X stops reading; blank or nonnumeric Y is rejected. Copy the supplied CSV values into a workbook before using the picker.

With a part and 2D sketch active, run `Main.Main`, select the workbook, and enter origin X/Y in inches. Conversion is `meters = inches × 0.0254`. The point creator uses `x = originX + inputX`, `y = originY + inputY`, `z = 0`. Grouping uses a coordinate tolerance of `1e-9` meters. Equal X groups generate vertical relations; equal Y groups generate horizontal relations. Review constraint state afterward.

`OrdinateDimensions`, `Utilities`, and `clsPointGroup` are empty stubs. `OriginManager` is an alternative origin-selection module and is not called by `Main`.

## Hole prototype

Review `Create_holes_from_sketch_.main` in the VBA editor before running. It waits for face/point selections, requests a drill angle in degrees and dimensions in inches, converts angle to radians and dimensions to meters, and maps diameter to the nearest table entry. Verify that the desired drill code agrees with the supplied diameter and that the selected face is supplied correctly to Hole Wizard. No geometric output has been independently verified here.

## Drawing and annotation experiments

The drawing-creation recording requires local template, sheet-format, and part paths. It inserts one front view at recorded coordinates. `Drawing_template_change1.main` only initializes the application: it is not an implemented template-switching tool. Annotation macros need explicit termination and selection checks before routine use.
