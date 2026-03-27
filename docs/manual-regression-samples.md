# Manual Regression Samples

Use these Mermaid snippets after routing/layout changes to catch regressions quickly.

## Main Flow + Loop

```mermaid
flowchart TD
    A(开始) --> B{用户登录}
    B -->|Y| C[用户语音指令]
    B -->|N| B
    C --> D[ASRPro指令解析]
    D --> E[语音识别]
```

Check:
- Main flow exits from bottom and enters from top.
- Self-loop stays black and does not affect branch counting.
- `N` sits above the loop line.

## Decision Branching

```mermaid
flowchart TD
    A[输入] --> B{简单指令}
    B -->|Y| C[执行指令]
    B -->|N| D[DeepSeek指令解析]
    D --> E{处于指令表}
    E -->|Y| C
    E -->|N| F[生成回答]
```

Check:
- Decision nodes prefer left/right exits when branching.
- Incoming edges still prefer the top side when available.
- No node side is reused while another side is free.

## Merge Into End Node

```mermaid
flowchart TD
    A[控制硬件操作] --> D(结束)
    B[执行软件操作] --> D
    C[扬声器] --> D
```

Check:
- Incoming edges pick the top side first, then nearest free side.
- Paths remain visually short after layout changes.

## Single-Exit Decision

```mermaid
flowchart TD
    A[输入] --> B{判断}
    B --> C[处理]
    C --> D(结束)
```

Check:
- The decision node uses the bottom side when it has only one outgoing edge.
- The target node still prefers the top side for incoming flow.

## Mixed Incoming Sides

```mermaid
flowchart TD
    A[上游1] --> D[汇聚节点]
    B[上游2] --> D
    C[上游3] --> D
```

Check:
- One incoming edge should claim the top side first.
- Remaining incoming edges should fall back to the nearest free side without reusing a free side late.
