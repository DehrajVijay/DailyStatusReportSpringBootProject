# Daily Status Report Web

A local Spring Boot web application that automates daily status report preparation: enter tasks in a browser, generate a formatted Excel file, and open a pre-filled Outlook draft in one click.

Built to replace a manual routine that took 15-20 minutes each night (copying the prior day's workbook, reformatting, cropping into Outlook, retyping recipients and subject).

## What it does

1. Open the app in your browser (defaults to `http://localhost:8081/`)
2. Fill in tasks on a pre-loaded table (recurring rows like SOD and Break included)
3. Click **Generate**
4. App creates `DailyStatus_YYYY-MM-DD.xlsx` and opens an Outlook draft with an inline HTML table
5. Review and click **Send** in Outlook (the app never sends automatically)

Output folder: `Documents\DailyStatusReports\`

## Tech stack

| Technology | Version | Role |
|---|---|---|
| Java | 17 | Language |
| Spring Boot | 3.4.5 | Web framework, dependency injection |
| Spring MVC | — | REST controller, form binding |
| Thymeleaf | — | Server-side HTML templates |
| Apache POI | 5.2.5 | Excel (.xlsx) generation |
| Maven | wrapper | Build |
| PowerShell + Outlook COM | — | Outlook draft automation (Windows) |

## Architecture

Layered MVC with a thin controller and a web-agnostic service layer:

```
Browser -> ReportController -> ReportService -> TaskParser
                                            -> StatusReportExcelWriter (POI)
                                            -> OutlookDraftService (PowerShell)
```

- **17 Java classes** across controller, service, validation, output, and config layers
- Constructor-based dependency injection throughout
- Two-layer properties config (JAR defaults + user overrides in Documents)
- HTML output escaped to prevent XSS; `ProcessBuilder` uses argument lists (no shell injection)

For full architecture diagrams, class-by-class walkthrough, and interview prep, see [PROJECT-GUIDE.md](PROJECT-GUIDE.md).

## Prerequisites

- Java 17
- Windows with **desktop Outlook** installed
- `mail.to`,`mail.cc`,`mail.subject.prefix`,`mail.greeting`,`mail.signature` configured in `Documents\DailyStatusReports\outlook-config.properties`

## How to run

**Option A — batch file**

```
Double-click WebDailyStatus.bat
```

**Option B — command line**

```bash
mvnw.cmd package -DskipTests
java -jar target\daily-status-report-web-1.0.0.jar
```

**Option C — IDE**

Run `DailyStatusReportWebApplication.main()`.

On first run, edit the copy of `outlook-config.properties` in `Documents\DailyStatusReports\` and set `mail.to`.

## Project structure

```
src/main/java/com/statusreport/
  ReportController.java          # Web layer (GET /, POST /generate)
  ReportService.java             # Orchestrates validate -> Excel -> Outlook
  TaskParser.java                # Validation
  StatusReportExcelWriter.java   # Apache POI
  OutlookDraftService.java       # PowerShell bridge
  ...
src/main/resources/
  templates/index.html           # Thymeleaf UI
  send-outlook-draft.ps1         # Outlook COM script
```

## Author

**Vijay Dehraj** — Java backend developer  
[LinkedIn](https://linkedin.com/in/vijay-dehraj-developer) · [GitHub](https://github.com/DehrajVijay)
