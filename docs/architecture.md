# Architecture Flowchart: Workflow of the App

> **Cursor preview shows blank boxes?** That is normal. Cursor's built-in markdown preview
> does **not** render Mermaid diagrams, even with the Matt Bierner extension installed.
>
> **Use one of these instead (all show colorful diagrams):**

| Method | How |
|---|---|
| **Recommended** | Double-click `docs/architecture-viewer.html` — opens in Chrome/Edge with all 7 diagrams |
| **From terminal** | Run `OpenArchitectureDiagrams.bat` in the project root |
| **On GitHub** | Push the repo and open `PROJECT-GUIDE.md` in your browser on GitHub |

---

## How to open the diagram viewer

1. In the file explorer, go to the `docs` folder
2. Double-click **`architecture-viewer.html`**
3. Your browser opens with colorful, scrollable architecture diagrams

Or run from the project root:

```
OpenArchitectureDiagrams.bat
```

---

## What the diagrams show

| # | Diagram | What it explains |
|---|---|---|
| 1 | Application startup | JAR → Spring Boot → Tomcat → browser opens |
| 2 | Generate sequence | POST /generate → validate → Excel → Outlook draft |
| 3 | System context | You, browser, Tomcat, files, PowerShell, Outlook, mail server |
| 4 | Layered architecture | Presentation → Web → Model → Business → Output → Infrastructure |
| 5 | Bean wiring | Spring singletons vs objects created by Java code |
| 6 | Class relationships | ReportController, ReportService, writers, models |
| 7 | Data transformation | Browser input → TaskEntry → Excel file + email HTML |

---

## Why Cursor preview is blank

The tab titled **"Preview architecture.md"** with an **[Edit Markdown file]** button is
**Cursor's own preview panel**. It is not the same as VS Code's markdown preview (`Ctrl+Shift+V`).

| Preview type | Renders Mermaid? |
|---|---|
| Cursor built-in preview (your screenshot) | **No** — blank box |
| VS Code `Ctrl+Shift+V` + Matt Bierner extension | Sometimes yes |
| `architecture-viewer.html` in browser | **Yes — always** |
| GitHub in browser | **Yes — always** |

The Matt Bierner extension only hooks into the **standard** markdown preview command, not Cursor's
custom preview panel.

---

## Full written guide

For class-by-class explanations, see **`PROJECT-GUIDE.md`** in the project root.
The HTML viewer is the visual companion to sections 5 and 14 of that guide.
