# Verification record and acceptance checks

Completed during portfolio preparation: extraction of VBA source from all 16 OLE macro containers; classification of 29 non-host source modules; source-based capability and limitation review; hashes of supplied binaries; scan of published text for personal local paths; local Markdown-link verification.

Not performed: VBA compilation, SolidWorks or Excel execution, CAD geometry validation, reference integrity after saving, version compatibility testing, timing or productivity measurement.

For a local acceptance pass:

| Scenario | Expected check |
| :--- | :--- |
| Missing/wrong active document | Supported entry points reject unsuitable context without changes |
| Duplicate part instances | Featured rename processes one candidate per unique filepath |
| Existing destination | Rename preview reports conflict and aborts before changes |
| Referenced drawing | Drawing opens with the intended renamed model after save/reopen |
| Coordinate input `(1, 2)` with origin `(0, 0)` | Sketch point is `(0.0254, 0.0508, 0)` meters |
| Two points sharing X | Intended vertical relation appears without overconstraint |
| Invalid Excel cell | Import reports the row and releases Excel objects |
| Hole creation | Feature count, diameter, depth, angle and face match requested input |

Check reports and actual documents: success messages alone are insufficient to establish correct CAD output.
