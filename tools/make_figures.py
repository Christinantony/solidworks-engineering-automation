"""Draw the README illustrations for this repository.

SolidWorks cannot run here, so none of these are CAD screenshots. Each
figure is computed from the macro sources themselves:

  assets/point_importer.png   examples/coordinates.csv pushed through the
                              importer's arithmetic (inches × 0.0254, origin
                              offset, 1e-9 m grouping, vertical/horizontal
                              relations) re-implemented in Python
  assets/rename_workflow.png  the nine numbered steps of
                              Assembly_Batch_rename1.main, with a preview in
                              the exact text format the macro writes
  assets/drill_codes.png      the ANSI drill table parsed out of
                              Create_holes_from_sketch_.bas and the
                              nearest-size rule it applies

    python tools/make_figures.py
"""
import csv
import re
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "assets"
OUT.mkdir(exist_ok=True)
NOTE = "Illustration computed from the macro source; not a SolidWorks screenshot."

# ---------------------------------------------------------------- point importer
INCH = 0.0254                 # SketchTools / ExcelIO: meters = inches × 0.0254
COORD_TOLERANCE = 1e-9        # Relations.bas
origin_x, origin_y = 0.0, 0.0 # InputManager prompt, in inches, converted the same way

pts = []
with open(ROOT / "examples" / "coordinates.csv") as f:
    for i, (x, y) in enumerate(csv.reader(f), start=1):
        pts.append({"id": i, "x": float(x) * INCH, "y": float(y) * INCH, "xg": 0, "yg": 0})
for p in pts:
    p["sx"], p["sy"] = origin_x * INCH + p["x"], origin_y * INCH + p["y"]

def build_groups(key, tag):      # BuildXGroups / BuildYGroups, anchor-first grouping
    g = 0
    for a in pts:
        if a[tag] == 0:
            g += 1
            a[tag] = g
            for c in pts:
                if c[tag] == 0 and abs(c[key] - a[key]) < COORD_TOLERANCE:
                    c[tag] = g
build_groups("x", "xg")
build_groups("y", "yg")

def relations(tag):              # CreateVerticalRelations / CreateHorizontalRelations: anchor to every other member
    rels = []
    for g in range(1, max(p[tag] for p in pts) + 1):
        members = [p for p in pts if p[tag] == g]
        for c in members[1:]:
            rels.append((members[0], c))
    return rels
vertical, horizontal = relations("xg"), relations("yg")

fig, ax = plt.subplots(figsize=(9, 6.2), dpi=120)
for a, b in vertical:
    ax.plot([a["sx"], b["sx"]], [a["sy"], b["sy"]], "--", color="#2563eb", lw=1.6, label="sgVERTICALPOINTS2D" if (a, b) == vertical[0] else None)
for a, b in horizontal:
    ax.plot([a["sx"], b["sx"]], [a["sy"], b["sy"]], "--", color="#16a34a", lw=1.6, label="sgHORIZONTALPOINTS2D" if (a, b) == horizontal[0] else None)
ax.scatter([p["sx"] for p in pts], [p["sy"] for p in pts], s=90, color="#111", zorder=3, label="CreatePoint(x, y, 0)")
offsets = {1: (-14, -14, "right", "top"), 2: (14, -14, "left", "top"), 3: (14, 10, "left", "bottom"), 4: (-14, 10, "right", "bottom")}
for p in pts:
    dx, dy, ha, va = offsets[p["id"]]
    ax.annotate(f"point {p['id']}\n({p['sx']:.4f}, {p['sy']:.4f}) m\nX group {p['xg']} · Y group {p['yg']}",
                (p["sx"], p["sy"]), textcoords="offset points", xytext=(dx, dy), fontsize=8.5, ha=ha, va=va)
ax.plot(0, 0, "+", color="#dc2626", ms=14, mew=2, label="sketch origin (offset 0, 0)")
ax.set_aspect("equal")
ax.set_xlim(-0.03, 0.085)
ax.set_ylim(-0.022, 0.075)
ax.set_xlabel("x (m)")
ax.set_ylabel("y (m)")
ax.grid(color="#e5e7eb")
ax.set_axisbelow(True)
ax.set_title("Excel point importer: examples/coordinates.csv (inches) → sketch points and relations")
ax.legend(loc="upper right", fontsize=8.5)
rows = "\n".join(f"row {p['id']}: {p['x']/INCH:g}, {p['y']/INCH:g} in" for p in pts)
ax.text(0.98, 0.5, f"Input (column A, column B):\n{rows}\n\n{len(vertical)} vertical + {len(horizontal)} horizontal relations",
        transform=ax.transAxes, fontsize=8.5, va="center", ha="right", family="monospace",
        bbox=dict(boxstyle="round", fc="#f8fafc", ec="#cbd5e1"))
fig.text(0.01, 0.005, NOTE, fontsize=7.5, color="#555")
fig.tight_layout(rect=(0, 0.02, 1, 1))
fig.savefig(OUT / "point_importer.png")

# ---------------------------------------------------------------- rename workflow
steps = [
    ("1  Collect", "Traverse() the active configuration depth-first;\nkeep parts ending in _01, skip suppressed, virtual\nand unloaded components, one entry per file path"),
    ("2  Number", "Ask for the first part number (e.g. HC-01010);\nSplitPartNumber keeps prefix and digit width,\nthe rest count +1 in tree order"),
    ("3  Preview", "Write rename_preview.txt, find each part's .SLDDRW,\nflag ** CONFLICT when a target already exists;\nany conflict aborts before anything changes"),
    ("4  Open drawings", "OpenDoc6 every matching drawing so its\nreferences follow the rename; re-activate the\ntop assembly or stop"),
    ("5  Hook events", "AttachAsm / AttachPart handlers force\n\"Update where used\" on save"),
    ("6  Rename", "Select each component and call\nExtension.RenameDocument(newName);\nfailures go to the report with RenameErrName"),
    ("7  Save", "Parts first, then sub-assemblies\n(deepest first), then the top assembly"),
    ("8  Drawings", "SaveAs3 each drawing under the new name,\nclose it; old drawing deletion is disabled\n(DELETE_OLD_DRAWINGS = False)"),
    ("9  Report", "Parts renamed / saved, drawings renamed;\nopen rename_report.txt if anything needs attention"),
]
fig, ax = plt.subplots(figsize=(14, 7.2), dpi=120)
ax.axis("off")
cols, w, h = 3, 4.1, 1.55
for i, (title, body) in enumerate(steps):
    r, c = divmod(i, cols)
    x, y = c * 4.6, -r * 2.2
    face = "#fee2e2" if i == 2 else "#eff6ff"
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.08", fc=face, ec="#1e3a8a", lw=1.2))
    ax.text(x + 0.15, y + h - 0.18, title, fontsize=11, weight="bold", va="top")
    ax.text(x + 0.15, y + h - 0.55, body, fontsize=8.2, va="top", linespacing=1.35)
    if i < len(steps) - 1:
        nr, nc = divmod(i + 1, cols)
        if nr == r:
            ax.add_patch(FancyArrowPatch((x + w + 0.1, y + h / 2), (x + 4.6 - 0.1, y + h / 2), arrowstyle="-|>", mutation_scale=14, color="#1e3a8a"))
        else:
            ax.add_patch(FancyArrowPatch((x + w / 2, y - 0.1), (nc * 4.6 + w / 2, y - 2.2 + h + 0.1), arrowstyle="-|>", mutation_scale=14, color="#1e3a8a",
                                         connectionstyle="arc3,rad=0.0"))
ax.text(2 * 4.6 + w / 2, -2 * 2.2 - 0.55, "stops here on any conflict or a 'No' at the confirmation", ha="center", fontsize=8, color="#991b1b")
preview = ("PREVIEW - nothing has been changed yet\n" + "-" * 70 + "\n"
           "01  BRACKET_01.SLDPRT  ->  HC-01010.SLDPRT   | drawing -> HC-01010.SLDDRW\n"
           "02  SHAFT_01.SLDPRT  ->  HC-01011.SLDPRT   | drawing: none found\n"
           "03  COVER_01.SLDPRT  ->  HC-01012.SLDPRT   | drawing -> HC-01012.SLDDRW   ** CONFLICT: HC-01012.SLDPRT already exists\n"
           "\nSkipped:\n  Virtual component: Pin_01^Housing\n")
ax.text(0, -2 * 2.2 - 1.05, "rename_preview.txt, in the exact format the macro writes (example file names):", fontsize=8.5, va="top", weight="bold")
ax.text(0, -2 * 2.2 - 1.4, preview, fontsize=7.6, va="top", family="monospace",
        bbox=dict(boxstyle="round", fc="#fffbeb", ec="#f59e0b"))
ax.set_xlim(-0.2, 3 * 4.6)
ax.set_ylim(-2 * 2.2 - 3.6, h + 0.3)
ax.set_title("Assembly renaming: the nine stages of Assembly_Batch_rename1.main", fontsize=12)
fig.text(0.01, 0.005, NOTE, fontsize=7.5, color="#555")
fig.tight_layout(rect=(0, 0.02, 1, 1))
fig.savefig(OUT / "rename_workflow.png")

# ---------------------------------------------------------------- drill codes (parsed from the VBA source)
src = (ROOT / "src" / "holes-from-sketch-points" / "Create_holes_from_sketch_.bas").read_text()
table = [(float(s), c) for s, c in re.findall(r'Array\(([0-9.]+),\s*"([^"]+)"\)', src)]
assert len(table) > 100, "drill table not found"
def closest(dia):                # GetClosestDrillSizeCode: first entry wins ties, strict < comparison
    best = table[0]
    for size, code in table:
        if abs(size - dia) < abs(best[0] - dia):
            best = (size, code)
    return best
examples = [0.2, 0.3125, 0.45, 0.8]
fig, ax = plt.subplots(figsize=(12, 5.6), dpi=120)
sizes = sorted(s for s, _ in table)
ax.vlines(sizes, 0, 1, color="#94a3b8", lw=0.6)
ax.set_ylim(0, 1)
ax.set_yticks([])
ax.set_xlim(0, 1.02)
ax.set_xlabel("requested hole diameter (inches, as typed into the InputBox)")
for k, d in enumerate(examples):
    size, code = closest(d)
    ax.axvline(d, color="#dc2626", lw=1.2)
    ax.annotate(f'input {d}" → nearest {size}" = code "{code}"\n→ HoleWizard5(2, 0, 18, "{code}", 0, {d*INCH:.5f} m, depth, …)',
                (d, 0.86 - k * 0.2), xytext=(8, 0), textcoords="offset points", fontsize=8.5, va="center",
                bbox=dict(boxstyle="round", fc="white", ec="#dc2626"))
numbered = [s for s, c in table if c.startswith("#")]
letters = [s for s, c in table if c.isalpha() and len(c) == 1]
fractions = [s for s, c in table if "/" in c or c == "1"]
for ys, data, lab, col in [(0.06, numbered, "wire gauge #80…#1", "#2563eb"), (0.12, letters, "letter sizes A…Z", "#16a34a"), (0.18, fractions, "fractional 1/32…1", "#9333ea")]:
    ax.scatter(data, [ys] * len(data), s=14, color=col, label=f"{lab} ({len(data)} entries)", zorder=3)
ax.legend(loc="lower right", fontsize=8.5)
ax.set_title(f"Hole prototype: the {len(table)}-entry ANSI drill table in the macro and its nearest-size rule")
fig.text(0.01, 0.005, NOTE + " Table parsed from the .bas file at run time.", fontsize=7.5, color="#555")
fig.tight_layout(rect=(0, 0.02, 1, 1))
fig.savefig(OUT / "drill_codes.png")
print("wrote", sorted(p.name for p in OUT.glob("*.png")))
