# Architecture Overview

## Main Flow

1. `ThisAddIn` creates or returns the Ribbon entry point.
2. `Ribbon` resolves the current Visio application and forwards the action.
3. `MermaidFlowchartService` shows the input form, validates Mermaid, and coordinates generation.
4. `MermaidParser` converts Mermaid text into `FlowchartData`.
5. `VisioFlowchartGenerator` orchestrates shape creation, layout, and connection routing.

## Core Components

### `MermaidFlowchartService`

- Owns the user-facing execution flow.
- Reuses parsed form data when available to avoid duplicate parsing.
- Centralizes success and error dialogs.

### `MermaidParser`

- Parses the supported Mermaid flowchart subset.
- Normalizes nodes, inline node declarations, and connections.
- Deduplicates connections before generation starts.

### `VisioFlowchartGenerator`

- Keeps orchestration only.
- Ensures an active document/page exists.
- Delegates shape creation, layout, and connector routing to dedicated collaborators.

### `VisioFlowchartShapeFactory`

- Creates Visio shapes for each parsed node.
- Applies node sizing and visual styling.

### `VisioFlowchartLayoutEngine`

- Configures the page.
- Applies the manual layered layout used by the add-in.
- Runs the final page-level layout adjustment.

### `VisioFlowchartConnectionRouter`

- Assigns connection sides.
- Applies connector styling.
- Handles self-loop drawing separately from normal connectors.

## Current Routing Defaults

- Incoming edges prefer the top side when it is still free.
- Outgoing edges from normal nodes prefer the bottom side.
- Decision nodes with a single outgoing edge prefer the bottom side.
- Decision nodes with multiple outgoing edges prefer the left and right sides first.
- Self-loops are drawn separately and do not participate in normal side-allocation rules.
- When no primary side is free, the router falls back to the nearest available side.

## Current Maintenance Rules

- UI entry logic should stay in `ThisAddIn`, `Ribbon`, and `MermaidFlowchartService`.
- Mermaid syntax support should be added in `MermaidParser` first.
- Shape appearance changes belong in `VisioFlowchartShapeFactory`.
- Spacing and node placement changes belong in `VisioFlowchartLayoutEngine`.
- Connection-side rules, labels, and loop behavior belong in `VisioFlowchartConnectionRouter`.

## Manual Regression

Use [manual-regression-samples.md](/E:/项目/Viso流程图生成插件/VisioFlowchartExtractor/VisioAddIn1/docs/manual-regression-samples.md) after routing or layout changes.
