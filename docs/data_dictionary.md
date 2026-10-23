# Data dictionary: tsv-to-sysmlv2

## Input: 33-column TSV from `templates/SysML17_Migration_Export_w_Diagrams.vm`

Tab-separated, UTF-8, one header row, one row per element or diagram.
Empty cells mean the column does not apply to that element type.

| # | Column | Meaning |
|---|---|---|
| 1 | Phase | Export phase number assigned by the template (ordering aid only). |
| 2 | SysML 1.7 Element Type | Metaclass or stereotype label of the element, for example Block, Requirement, Diagram. |
| 3 | Name | Element name. Line breaks inside names are exported as the token `<NL>`. |
| 4 | Owner Qualified Name | Qualified name of the owning element (`::`-separated). |
| 5 | Applied Stereotype | Applied stereotype names, if any. |
| 6 | Qualified Name | Full qualified name of the element. |
| 7 | Type Qualified Name | Qualified name of the element's type (properties, ports, parameters). |
| 8 | Multiplicity | Multiplicity string, for example `1`, `0..*`. |
| 9 | Direction | Direction of a port, flow property or parameter (in, out, inout). |
| 10 | Source/Client Qualified Name | Source end (connectors, flows) or client (dependencies). |
| 11 | Src End Part With Port QN | Part-with-port context of the source connector end. |
| 12 | Target/Supplier Qualified Name | Target end (connectors, flows) or supplier (dependencies). |
| 13 | Tgt End Part With Port QN | Part-with-port context of the target connector end. |
| 14 | Item Conveyed Qualified Name | Item conveyed by an item flow. |
| 15 | Called Behavior Qualified Name | Behavior invoked by a call behavior action. |
| 16 | Source Node Qualified Name | Source of an activity edge (control or object flow). |
| 17 | Target Node Qualified Name | Target of an activity edge. |
| 18 | Trigger | Transition trigger text. |
| 19 | Guard Expression | Transition or edge guard text. |
| 20 | Signal Qualified Name | Signal referenced by a trigger, accept event or send action. |
| 21 | Element ID | MagicDraw element ID. Matches the Element ID column of the CATIA migration report. |
| 22 | Diagram Type | Diagram kind for Diagram rows, for example SysML Block Definition Diagram. |
| 23 | General/Parent QN | Generalization target (parent classifier). |
| 24 | CE0 Role QN | Role of connector end 0. |
| 25 | CE1 Role QN | Role of connector end 1. |
| 26 | State Entry Behavior QN | Entry behavior of a state. |
| 27 | State Exit Behavior QN | Exit behavior of a state. |
| 28 | State Do Activity QN | Do activity of a state. |
| 29 | Transition Effect QN | Effect behavior of a transition. |
| 30 | Pseudostate Kind | Kind of a pseudostate, for example initial, choice, junction. |
| 31 | Partition Represents QN | Element represented by an activity partition (swimlane). |
| 32 | Association End 0 | First association end (name and type). |
| 33 | Association End 1 | Second association end (name and type). |

## Output 1: SysML v2 text (`.sysml`)

Plain UTF-8 text in SysML v2 textual notation. The first three lines are comments naming
the generating script, the dialect and the source TSV. `// TODO-MIGRATE` comments mark
constructs that need a human pass; `// DIAGRAM (not migrated)` comments mark where each
diagram lived. Both kinds also appear as rows in the report workbook.

## Output 2: migration report workbook (`*_MigrationReport.xlsx`)

Same sheet layout as the CATIA Magic migration plugin's `MigrationReportOf<project>.xlsx`.
`Element Url` is a synthesized `mdel://<Element ID>` link; `Reason` says why the item
could not be migrated automatically.

### Sheet `Not Migrated Elements`

Columns: `Element MetaType` | `Element Name` | `Qualified Name` | `Element Url` | `Element ID` | `Reason`

### Sheet `Not Migrated Properties`

Columns: `Element Name` | `Element MetaType` | `Property Name` | `Qualified Name` | `Element Url` | `Reason`

### Sheet `Not Migrated Diagrams`

Columns: `Diagram Type` | `Element Name` | `Qualified Name` | `Element Url` | `Element ID` | `Reason`
