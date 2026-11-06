# Access Review MCP server

An MCP server that lets Claude review the forms of a Microsoft Access front end without using your screen. It starts its own Access instance over COM, keeps every Access window just past the edge of the desktop, and captures forms with `PrintWindow`, which asks Access to draw a window into a bitmap. Nothing comes to the front, nothing takes focus, and no mouse or keyboard input is sent, so you can keep working on the same PC while a review runs.

It was built for visual review of apps made with [access-modern-ui](https://github.com/stephen3303/access-modern-ui), but the core tools work with any Access front end. It also reads, compares and captures pages of local OneNote notebooks, for apps that publish pages to OneNote; see [OneNote pages](#onenote-pages). And it drives the other Office apps (Excel, Word, PowerPoint, Visio, Project, Publisher, and Outlook when you enable it) through their COM object models, with typed tools for common jobs and a generic tool that reaches any method or property; see [Office apps](#office-apps).

```
AccessReview.ps1      the review functions (usable from plain PowerShell too)
AccessReviewMcp.ps1   the MCP server (stdio) that exposes them as tools
FormSnapshot.ps1      reads and compares form snapshots (no Access needed)
FormSnapshot.Tests.ps1  checks for FormSnapshot.ps1
FormDesign.ps1        writes form snapshots as JSON designs (no Access needed)
FormDesign.Tests.ps1  checks for FormDesign.ps1
ReviewSentry.ps1      optional reporting of findings to Sentry (off unless a DSN is set)
ReviewSentry.Tests.ps1  checks for ReviewSentry.ps1 (no Access or network needed)
OneNoteReview.ps1     the OneNote tools (OneNote desktop over COM)
OneNotePage.ps1       reads and compares OneNote page XML (no OneNote needed)
OneNotePage.Tests.ps1   checks for OneNotePage.ps1
OfficeReview.ps1      the Office tools (Excel, Word, PowerPoint, Outlook... over COM)
OfficeCom.ps1         member paths, safety rules and value conversion (no Office needed)
OfficeCom.Tests.ps1   checks for OfficeCom.ps1
OfficeMember.Tests.ps1  checks the office_invoke member walk with stand-in objects (no Office needed)
OfficeSmoke.ps1       end-to-end check of the Office tools on a review machine with Office
manifest.json         desktop extension manifest (for the .mcpb package)
```

## Requirements

- Windows 10 or 11, or Windows Server with a desktop session.
- Microsoft Access, desktop version. Tested with Access in Microsoft 365; the free Access Runtime is untested.
- Windows PowerShell 5.1 and .NET Framework 4.x, which come with Windows. Nothing else is installed.
- A front end you trust, in a trusted location. Under automation Access runs startup code (AutoExec, startup form) without the Enable Content bar, and `run` can call any public procedure. So `open_app` refuses files downloaded from the internet (Mark of the Web, `Zone.Identifier` zone 3 or 4) and, when **Front-end folder** / `ACCESS_REVIEW_ALLOWED_FOLDERS` is set, any file outside it. The extension requires that setting; set it in a config file too.

## Install

### Claude Desktop: extension file

1. Get `access-review-<version>.mcpb` (or build it, see [Build the package](#build-the-package)).
2. In Claude Desktop open **Settings > Extensions**, then drag the file in (or use **Advanced settings > Install extension**).
3. Fill in the optional settings (see [Settings](#settings)) and turn it on.

### Claude Desktop: config file

In Claude Desktop open **Settings > Developer > Edit Config**. That opens `claude_desktop_config.json`. Add an `access-review` entry under `mcpServers`, using the folder you put the scripts in, then restart Claude Desktop:

```json
{
  "mcpServers": {
    "access-review": {
      "command": "powershell.exe",
      "args": ["-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
               "-File", "C:\\Tools\\access-review\\AccessReviewMcp.ps1"]
    }
  }
}
```

If the file already has other servers, add the `access-review` entry next to them inside the same `mcpServers` object. Backslashes in paths are doubled in JSON.

### Claude Code

Add the same entry to a project's `.mcp.json`, or run:

```
claude mcp add access-review -- powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File C:\Tools\access-review\AccessReviewMcp.ps1
```

`-ExecutionPolicy Bypass` applies only to this one PowerShell process. If your organisation enforces an execution policy through Group Policy, that policy still wins; ask your administrator.

## Settings

All optional. In the extension they are fields on its settings page; with a config file they go in an `env` object on the server entry.

| Setting | Environment variable | Default | What it does |
|---|---|---|---|
| Capture folder | `ACCESS_REVIEW_OUTPUT` | `%LOCALAPPDATA%\AccessReview` | Full path. Where PNGs are saved, in a new time-stamped folder per session. `open_app` can also take `output_dir`. |
| Front-end folder | `ACCESS_REVIEW_ALLOWED_FOLDERS` | none (extension: required) | Folders `open_app` and `office_open` may open files from, including subfolders, separated by `;`. Full paths. When unset, `open_app` can open any existing file, and the Office tools use only the capture folder. |
| Shared lock files | `ACCESS_REVIEW_LOCKS` | none | Full paths of lock files, separated by `;`, for a PC several reviewers take turns on. `review_lock` takes all or none; with none configured it does nothing. |
| Palettes | `ACCESS_REVIEW_PALETTES` | `Navy=1;Midnight=2;Teal Navy=6` | Palettes `set_palette` offers, as `name=number`. The number is passed to the palette procedure. The first one is the default: `close_app` switches back to it. A malformed value is ignored with a warning (in the server log, its instructions and each `set_palette` result) and the defaults are used. |
| Palette procedure | `ACCESS_REVIEW_PALETTE_PROCEDURE` | `SetPalette` | Public VBA `Sub Name(ByVal i As Long)` in your front end that switches palettes. |
| Sentry DSN | `ACCESS_REVIEW_SENTRY_DSN` | none (off) | Reports findings to a Sentry project; see [Reporting to Sentry](#reporting-to-sentry). |
| Office apps | `ACCESS_REVIEW_OFFICE_APPS` | `excel;word;powerpoint;visio;project;publisher;access` | Apps the `office_*` tools may drive. Outlook is left out: add `outlook` only if Claude may read this PC's Outlook mail (read only). |
| Attach to running Office apps | `ACCESS_REVIEW_OFFICE_ATTACH` | none | PowerPoint, Outlook, Project and Publisher keep one instance per user. While you have one open, `office_start` refuses rather than take it over, unless it is listed here. An attached app is never quit, captured or given documents. |
| (config file only) | `ACCESS_REVIEW_SENTRY_ENVIRONMENT` | `review` | Environment name on each Sentry report, for example the machine's role. |

Example `env` for a config file:

```json
"env": {
  "ACCESS_REVIEW_ALLOWED_FOLDERS": "D:\\Review\\FrontEnds",
  "ACCESS_REVIEW_OUTPUT": "D:\\Reviews",
  "ACCESS_REVIEW_PALETTES": "Light=1;Dark=2",
  "ACCESS_REVIEW_PALETTE_PROCEDURE": "ApplyTheme"
}
```

## Tools

| Tool | What it does |
|---|---|
| `open_app` | Opens an existing front end in a new Access instance (never creates a file). Refuses downloaded files and files outside the front-end folder. `mode`: `offscreen` (default) or `onscreen` for a dedicated review machine. |
| `close_app` | Resets the palette to the default, quits without saving design changes and confirms the process is gone. |
| `open_form` | Opens a form and sizes it to its design size (form width, visible section heights), or to `width_px` x `height_px` at 96 dpi. Reports whether the Access window clips it. |
| `resize_form` | Design size, a given size or maximized (Access window at the main display's work area, as a real maximized window gets). A shell form that fills the Access window is sized by resizing the Access window until its inside size matches; `SizeMatches` reports whether it did. |
| `capture` | PNG of a form, or a zoomed crop, taken once two frames in a row match (so a form still drawing isn't caught mid-draw; `Stable` reports it). Returned as an image and saved in the capture folder, named form-palette-label. |
| `controls` | Each control's Left/Top/Width/Height (96 dpi px), section, caption and colours, for measuring gaps and alignment. |
| `snapshot` | Saves the definition of each form (`SaveAsText`) under a label, in `snapshots\<label>` in the capture folder's base. Without `forms`, every form, replacing that label's earlier snapshot. Changes nothing in the front end. |
| `compare_snapshots` | Compares two labels property by property, without Access: forms and controls added or removed, and each changed property before and after. See [Comparing two builds](#comparing-two-builds). |
| `export_design` | Writes a label's forms as JSON designs, without Access, for rebuilding them in another app or as HTML. See [Exporting form designs](#exporting-form-designs). |
| `dialogs` | Lists and captures open message boxes and popups; `close=true` closes them. Also lists dialogs seen during earlier calls. |
| `error_log` | `mark=true` before testing, then `mark=false` to list new rows in an `error_log` table and new text in the file a public `LogFilePath()` function returns. Says so when the front end has neither. |
| `compile` | Compiles and saves all modules. |
| `run` | Runs a public procedure with `Application.Run`, for example a sign-in. Whole numbers are passed as Long (Double when too large) and output arguments come back. |
| `set_palette` | Runs the palette procedure, then reopens the open forms at their last size (a form left open keeps part of the old drawing). A form with a public `Reopen` method reopens itself on its current page (read from its `PageKey`), so a signed-in user stays signed in; if that fails the call returns an error and never closes the form. Other forms are closed and reopened with their OpenArgs. |
| `review_lock` / `review_unlock` | Takes and releases the shared lock files, re-reading after 3 s to catch a race. Releases only its own. |
| `onenote_notebooks` / `onenote_open_notebook` | Lists the local notebooks open in OneNote that the tools may use; opens an existing local notebook. See [OneNote pages](#onenote-pages). |
| `onenote_pages` / `onenote_page` | Lists sections and pages (with each page's record ID given `meta_name`); reads one page as a summary and saves its XML. |
| `onenote_capture` | PNG of a page, drawn in a new OneNote window kept off screen, then closed. |
| `onenote_snapshot` / `onenote_compare` | Saves every page's XML under a label; compares two labels page by page. |
| `onenote_check_ids` | Pages without a record ID, IDs on more than one page, and expected IDs missing or unexpected. |
| `office_apps` / `office_start` / `office_quit` | Office apps installed and enabled here; start one in its own off-screen instance; close its documents without saving and quit it. See [Office apps](#office-apps). |
| `office_open` / `office_new` / `office_close` | Open an existing Excel, Word, PowerPoint or Visio file from the allowed folder (read-only, macros off by default); make a blank one; close without saving. |
| `office_save_as` / `office_export_pdf` / `office_capture` | Save a copy in the allowed or capture folder; write a PDF; PNG of an app or document window. |
| `office_invoke` / `office_members` / `office_release` | Get, set or call any member path of an object model (objects come back as handles); list what an object offers from its type library; forget handles. |
| `office_run_macro` | Runs a macro in a document opened with `enable_macros=true`. |
| `excel_range` / `word_text` / `powerpoint_slides` / `outlook_items` | Read or write cells; a document's paragraphs and styles; slides with shapes, text and tags; mail items in a folder (read only). |

Every Access call runs under a guard. A window that appears during a call is parked off screen, captured and described. VBA and Office message boxes are closed so the call returns. A Yes/No box ignores Close, so the guard presses No, or the default button when there is no No. Another window (a dialog form opened modally) is closed if it is still blocking the call after 4 s, except while `open_app` starts the front end, so a splash or sign-in form stays open. A call that runs past its timeout ends Access. Each tool's result lists what the guard saw.

In off-screen mode a watcher keeps every Access window, including popups the app opens or re-centres on its own, past the desktop edge for the whole session; a newly shown popup can still appear for up to about 50 ms before it is moved.

### What your front end needs

- Nothing, for opening, sizing, capturing, controls, dialogs, compile and run.
- For `error_log`: a table named `error_log` with columns `id`, `logged_at`, `error_number`, `description`, `source`, `active_form`, `active_control`, and/or a public `LogFilePath()` function returning a text log's path. [access-modern-ui's `modErrors.bas`](https://github.com/stephen3303/access-modern-ui/blob/main/examples/modules/modErrors.bas) provides both.
- For `set_palette`: a public palette procedure (see [Settings](#settings)). Apps built on access-modern-ui have `SetPalette`.

### Reporting to Sentry

Off unless **Sentry DSN** / `ACCESS_REVIEW_SENTRY_DSN` is set. Then each new `error_log` row or `errors.log` entry (reported once, however often `error_log` is read) and each VBA or Office message box, hang or Access exit that the guard sees becomes a Sentry event, so findings from every review run collect in one place with a history across runs. Kit dialogs and popup forms are not reported.

Only these fields are sent: the front end's file name, the error number, the procedure with its line, the form and control names, the kind of finding, the Access version and the tool version. The error description, call trail, context, the dialog's text, user names and the computer name are never sent, because an Access message can quote field values. Events group by front end, kind, error number and procedure (without the line), so the same error after an edit that moved it stays one issue; identical findings in one call go as one event with an `occurrences` tag. Self-hosted DSNs with a secret, port or path work too.

Each send has a 5 s timeout and never fails a tool call; after the first failed send the session stops trying. `error_log` results include `ReportedToSentry`, the number of findings the accepted events covered. Code: [`ReviewSentry.ps1`](ReviewSentry.ps1); checks without Access or a network: `powershell -NoProfile -File ReviewSentry.Tests.ps1`.

## Typical review run

1. `review_lock` with your name, if the PC is shared.
2. `open_app` on the front end, then `error_log` with `mark=true`.
3. For each page to check: `open_form` at its designed size, `capture`, `controls` to measure spacing, `dialogs`.
4. If the app has palettes: `set_palette`, `capture` again.
5. `resize_form` small (about 1100x700) and maximized, `capture`.
6. `error_log` (new entries are findings unless explained), `close_app`, `review_unlock`.

## Comparing two builds

A capture shows that something moved; a snapshot comparison says which control, which property and by how much, including changes a capture can't show, such as a hidden control, a removed event handler or a new record source.

1. Build the base branch into the review front end, `open_app`, `snapshot` with `label=main`, `close_app`.
2. Rebuild the same front end from the PR, `open_app`, `snapshot` with `label=pr`.
3. `compare_snapshots` with `before=main`, `after=pr`.

The report lists, per form, `+` for controls added and `-` for controls removed (with their position), and `~ control.Property: before -> after` for each change. Positions and sizes are in 96 dpi px, and colours are `#RRGGBB`. A form whose snapshot failed on either side is listed with `?` as not compared. The report is also saved as `snapshots\compare main to pr.txt`.

- Properties Access rewrites on every save are ignored: `Checksum`, `GUID`, `NameMap`, `RecSrcDt`, printer settings (`PrtMip`, `PrtDevMode`, `PrtDevNames`), `LayoutCached*`, `OverlapFlags`, and the form's own `Left`, `Top`, `Right` and `Bottom` (where its design window last sat; controls keep theirs).
- Pictures and other binary values compare by a hash, so a redrawn pill shows as changed without the bytes.
- The code behind each form is not compared; it is in the PR's own diff.
- A snapshot is the saved design. Colours and positions a form's code sets when it opens (for example a theme applied on load) are not in it; check those with `capture` and `controls`.
- Two snapshots of the same build should compare equal. If they don't, the property that differs is one more Access rewrites on save: add it to `$SnapshotIgnored` in `FormSnapshot.ps1`.

`FormSnapshot.Tests.ps1` checks the reading and comparing against a sample form: `powershell -NoProfile -File tools\review-mcp\FormSnapshot.Tests.ps1`.

## Exporting form designs

To hand a front end's forms to another app, or rebuild them as HTML pages, export a snapshot as JSON:

1. Build the front end, `open_app`, `snapshot` with a label such as `main`. `close_app` if nothing else is needed.
2. `export_design` with `label=main` and, for an access-modern-ui app, `theme_file` set to the full path of its build palette (`src\themes\light\modTheme.bas`).

It writes `designs\<label>` in the output folder's base, next to `snapshots`:

```
<form>.json    one per form
index.json     the forms written (with section and control counts) and the forms whose snapshot failed
tokens.json    with theme_file: each CLR_* token of the build palette as #RRGGBB, for CSS variables
```

Each `<form>.json` (schema `access-form-design/1`):

```json
{
  "schema": "access-form-design/1", "name": "dlgExample", "kind": "Form", "units": "px", "width": 400,
  "properties": { "Caption": "Input", "PopUp": "NotDefault", "RecordSource": "" },
  "events": { "OnLoad": "[Event Procedure]" },
  "sections": [
    { "name": "Detail", "type": "Section", "height": 180,
      "properties": { "BackColor": { "hex": "#FEFFFF", "token": "CLR_DIALOG" } }, "events": {},
      "controls": [
        { "name": "btnOK", "type": "CommandButton", "left": 288, "top": 132, "width": 96, "height": 32,
          "properties": { "Caption": "OK", "ForeColor": { "hex": "#FEFFFE", "token": "CLR_BTN_FG" } },
          "events": { "OnClick": "[Event Procedure]" }, "children": [] } ] } ]
}
```

- `left`, `top`, `width`, `height` and other sizes Access keeps in twips are 96 dpi px (twips / 15), the kit's `px()` units. A control's `left` and `top` are measured from its section's top-left corner, including controls on a tab page or in an option group, which are listed in that control's `children`. Sections are stacked top to bottom in the order listed.
- Every other saved property is in `properties` under its Access name, so nothing the snapshot holds is lost. Text is unquoted, whole numbers are numbers, colours are `{ hex, token }` (`token` only with `theme_file`; the generated palettes make each build colour unique, so the token is exact) and system colours are `{ system: <number> }`.
- Event properties (`On...`, `Before...`, `After...`) are in `events`: `[Event Procedure]` means the code behind the form handles it.
- A control saves only what differs from the form's defaults for its type; those defaults are merged in, so each control's `properties` is complete as far as the form records it. Properties left at Access's own built-in defaults are not saved by Access and so are not listed.
- `NotDefault` is kept as saved: the property is set to the value that is not Access's default (for a yes/no property, the opposite of the default).
- Pictures and conditional formats are a `<binary, n lines, hash>` placeholder. The code behind the form is not included; it is in source control (`src\forms\<form>.vba` in kit apps).
- The design is the form as the build saved it. What the form's code changes when it opens (kit forms are themed to the user's palette, resized to their content and have controls shown or hidden on load) is not in it. Map `token` to a CSS variable to get palette switching back.

`FormDesign.Tests.ps1` checks the export against a sample form: `powershell -NoProfile -File FormDesign.Tests.ps1`.

## OneNote pages

For apps that publish pages into OneNote (SPM's Publish to OneNote writes a leadership binder with a hidden `spm.key` record ID on each page and `spm.block` on the outline it owns). The tools use the OneNote desktop app's own COM API (`OneNote.Application`): no app registration, no Microsoft Graph, no sign-in, and nothing leaves the PC. They never edit a page.

**Boundary.** Only notebooks stored on this PC are used. A notebook whose path is a web address (OneDrive, SharePoint, a work tenant) is never listed, read or captured; `onenote_notebooks` gives only how many such notebooks are open. When **Front-end folder** / `ACCESS_REVIEW_ALLOWED_FOLDERS` is set, a local notebook must be inside it too. Review with a throwaway local notebook.

**Needs** the OneNote desktop app (Microsoft 365 or OneNote 2016), run at least once by the same Windows user. OneNote for Windows 10 has no COM API. OneNote is one process per user, so the tools attach to the one already running (or start it) and never quit it.

| Tool | Details |
|---|---|
| `onenote_open_notebook` | `path`: the notebook's folder or `.onetoc2`, full path. Opens it (never creates one). |
| `onenote_pages` | `notebook`, `section` (name or the listed path) and `meta_name` are optional. With `meta_name` each page's record ID comes back as `Key`. |
| `onenote_page` | `page`: page ID, record ID (with `meta_name`) or title. Returns the title, page Meta tags, page colour, quick styles (font, size, colour), and each outline with its position and size in points, its own Meta tags and its lines. A line is a paragraph with its indent, quick style, colours (`#rrggbb`) and tags, a table row with its cells and shading, or an image with its size. Saves the page XML in an `onenote-<time>` folder. |
| `onenote_capture` | Opens the page in a new OneNote window, parks that window past the desktop edge (`mode` `offscreen`, default) or at the top left (`onscreen`), sizes it (`width_px` x `height_px`, default 1280x900), captures until two frames match, then closes that window. The user's own OneNote windows keep their page. |
| `onenote_snapshot` | Saves each page's XML under `onenote-snapshots\<label>` with an `index.json`. Pages are keyed by record ID when they have one, else by section and title. |
| `onenote_compare` | Pages added and removed, and for each page: title, Meta tag and page colour changes, quick style changes, outlines added, removed, moved or resized (matched by their Meta tag, else by order), and lines added and removed. Saves the report as `onenote-snapshots\compare <before> to <after>.txt`. |
| `onenote_check_ids` | `meta_name` defaults to `spm.key`. Lists pages without it and IDs on more than one page; with `expected` (e.g. `command`, `people`, `program:12`) also IDs missing and IDs not expected. `Ok` is true when nothing is wrong. |

A typical check of a publish: open the throwaway notebook, `onenote_snapshot label=before`, publish from the app, `onenote_snapshot label=after`, `onenote_compare`, `onenote_check_ids` with the expected IDs, then `onenote_capture` the pages that changed. Publish twice and compare again: an update in place shows changed lines only, never a second page or a second `spm.block` outline.

Capture limits: OneNote may come to the front for a moment when it opens the page window. If a capture comes back one solid colour, OneNote did not draw while off screen on that machine; use `mode onscreen` on a dedicated review machine.

## Office apps

The `office_*` tools drive the other Office apps through the same COM object models VBA uses. They cover the whole object model without a tool per method:

- **`office_invoke`** walks a VBA-like member path from a target and reads the end, sets it (`value`) or calls it. A target is an app you started (`excel`, `word`, `powerpoint`, `outlook`, `visio`, `project`, `publisher`), `access` for the instance `open_app` started, or a handle. Arguments are literals: `'text'` or `"text"` (double a quote to include one), numbers, `true`, `false`, `null`, `missing` or an empty slot, and `@h3` for a handle; `Name:=value` passes a named argument. Extra JSON arguments (`args`) go after the last member's own, for blocks of cells (arrays of rows) or `{"handle": "h3"}`. Any object that comes back is kept as a handle (`h1`, `h2`...) to use as a later target or argument; `fields` reads members of it and `enumerate` lists a collection's items.
- **`office_members`** reads an object's type library: every method and property with its parameters (optional ones end in `?`), its description, and whether `office_invoke` refuses it. Use it to find the member you need instead of guessing.
- **Typed helpers** for the common jobs: `office_open`, `office_new`, `office_save_as`, `office_export_pdf`, `office_capture`, `office_run_macro`, `excel_range`, `word_text`, `powerpoint_slides`, `outlook_items`.

Examples of `office_invoke` paths:

| Target | Path | Effect |
|---|---|---|
| a workbook handle | `Worksheets('Summary').Range("A1:C4").Value2` | Reads the cells as rows. |
| a workbook handle | `Worksheets(1).Range('C1').Formula` with `value` `=SUM(A1:B1)` | Writes a formula. |
| a document handle | `SaveAs2(FileName:='C:\Review\out.docx', FileFormat:=16)` | Named arguments. |
| a presentation handle | `Slides(1).Shapes('Title 1').TextFrame.TextRange.Text` | Reads a title. |
| a slide handle | `Shapes.AddTextbox(1, 40, 40, 400, 60)` | Returns the new shape as a handle. |
| `excel` | `Workbooks` with `enumerate=true`, `fields` `["Name", "FullName"]` | Lists open workbooks. |

**Guards.** Each call keeps these rules (the lists are in `OfficeCom.ps1`):

- Only the apps in **Office apps** / `ACCESS_REVIEW_OFFICE_APPS`. Outlook is off unless listed, and its items are read only: properties can't be set, and `Delete`, `Move`, `Save`, `Reply`, `Forward` and the like are refused, as is `Send` everywhere. OneNote stays with the `onenote_*` tools, which keep to local notebooks.
- `office_start` starts the app's own instance, with alerts off and macros disabled for every file it opens (`AutomationSecurity` force-disable). It never takes over a single-instance app the user has open (see **Attach to running Office apps**). In off-screen mode every window of that instance is kept past the desktop edge.
- Files open only from the allowed folder (and the capture folder), never from the internet (Mark of the Web), through no junction or symbolic link, and read-only unless `read_only=false`. The same checks apply to a path given to any method through `office_invoke` (`Workbooks.Open`, `SaveAs`...). Macros run only in a file opened with `enable_macros=true`, through `office_run_macro`.
- Members that run code the caller supplies or reach outside the documents are refused: formula and macro engines (`ExecuteExcel4Macro`, `Evaluate`, `Eval`, `Run`, `WordBasic`), DDE, `Shell`, `SendKeys`, `OnTime`/`OnKey`, the VBA editor (`VBE`, `VBProject`), add-ins, `ExecuteMso`, hyperlinks (`Follow`, `FollowHyperlink`), mail (`Send`, `SendMail`, routing), printing (`PrintOut`; use `office_export_pdf`), `Quit` (use `office_quit`) and `AutomationSecurity`. Through `access`, `RunSQL`, `RunCommand` and opening or closing databases are refused too (the Access tools do those).
- Every text argument or value that looks like a path or address is checked: a full path must be inside the allowed or capture folder and contain no `..`, and web, `file:` and other addresses are refused. A method that reads or writes files (`SaveAs`, `Open`, `ExportAsFixedFormat`, `InsertFile`, `AddPicture`...) must be given a full path, not a bare name that would land in the app's default folder.
- Formulas that reach code, other programs or the web are refused: DDE links (`=cmd|...`), `CALL`, `REGISTER`, `EXEC`, `RTD`, `WEBSERVICE`, `FILTERXML`, `HYPERLINK`.
- Every call goes under the same guard as Access calls: message boxes are captured and closed (No when there is one), and a call that runs past its timeout ends the app, never one the tools attached to.
- Every `office_*`, `excel_*`, `word_*`, `powerpoint_*` and `outlook_*` call is appended to `office-audit.log` in the capture folder's base: time, tool, app, target, path, a 300-character preview of the arguments, and whether it succeeded or why not.

The generic tool is still a way to change documents: open copies in the allowed folder, not originals.

**Needs** desktop Office (Microsoft 365 or 2016 and later) on the same PC. `office_apps` says which apps are installed. Visio, Project and Publisher are supported through `office_invoke`; `office_open` handles Excel, Word, PowerPoint and Visio files. `outlook_items` needs an Outlook profile.

**Check it on a review machine** (it starts Excel, Word and, when installed, PowerPoint; never run it on a PC someone is using):

```
powershell -NoProfile -ExecutionPolicy Bypass -File OfficeSmoke.ps1
```

## Limits

- Hover effects, cursor changes and flicker during a real click can't be seen without a real pointer. Check those by hand.
- `PrintWindow` can't draw a minimized window, and a locked or disconnected Windows session draws nothing. The capture warns when an image comes back one solid colour. On a server, keep the review session unlocked: use auto-logon, or end Remote Desktop with `tscon <session id> /dest:console` instead of disconnecting.
- A window can't be bigger than the desktop, so the Access window is at most the desktop size (default 1920x1200). On a server with a small default resolution, raise it, or forms larger than the Access window will be clipped (`ClippedByAccessWindow` says so).
- Typing into fields and clicking buttons on a form isn't included yet. Use `run` for public procedures in the meantime.

PowerShell can't read Access form, control and section properties directly (their type information fails, so `$form.Hwnd` comes back empty). The tools read them with `Application.Eval("Forms![name]...")` and set the form's inside size late-bound.

## Using it from PowerShell

```powershell
. .\AccessReview.ps1
Open-ReviewApp -Path 'C:\Review\MyApp.accdb'
Get-ReviewErrorLog -Mark
Open-ReviewForm frmMain
Save-ReviewCapture frmMain -Label before
Get-ReviewErrorLog
Save-ReviewSnapshot -Label pr
Close-ReviewApp
Compare-ReviewSnapshot -Before main -After pr

. .\OneNoteReview.ps1
Open-ReviewNotebook -Path 'C:\Review\Notebooks\Binder'
Save-ReviewOneNoteSnapshot -Label before -MetaName spm.key
Get-ReviewOneNotePage -Page command -MetaName spm.key
Save-ReviewOneNoteCapture -Page command -MetaName spm.key -Label navy
Test-ReviewOneNoteRecordIds -Expected command, people

. .\OfficeReview.ps1
Start-OfficeApp excel
$wb = Open-OfficeDocument excel 'C:\Review\Budget.xlsx'
Invoke-ExcelRange -Handle $wb.Handle -Range 'A1:D10'
Invoke-OfficeMember -Target $wb.Handle -Path "Worksheets(1).Name"
Get-OfficeMembers -Target $wb.Handle -Filter 'Export*'
Stop-OfficeApp excel
```

## Build the package

From the repository root, with Node.js installed:

```
npx @anthropic-ai/mcpb pack . access-review.mcpb
```

That validates `manifest.json` and zips the folder (minus what `.mcpbignore` lists). When you change the tools, bump `version` in `manifest.json` and `$ServerVersion` in `AccessReviewMcp.ps1` together.

## Licence

MIT. See [LICENSE](LICENSE).
