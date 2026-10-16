# VBA Email Composer: required files only

| File | What it is | Edit it? |
|---|---|---|
| `modMailer.bas` | Shared engine: opens the .oft, applies sections and tokens, displays the mail | No |
| `IMailModel.cls` | Contract every model class follows | No |
| `modTheme.bas` | Common header, footer, colours and fonts for all forms | Settings only |
| `modTemplateEditor.bas` | Edit Template / Save Template for users | No |
| `modSetup.bas` | One-time setup and the T07 sample template | No |
| `frmMain_code.txt` | Main menu (paste into a blank UserForm `frmMain`) | No |
| `clsT07.cls` | **Example model** for T07 | Per template |
| `frmT07_code.txt` | **Example form code** for T07 | Per template |

## Install (new .xlsm)

1. Save a new workbook as `.xlsm` and press `Alt+F11`.
2. **Insert > UserForm**, name it `frmMain`, and paste `frmMain_code.txt` into its code window.
3. **File > Import**: `modMailer.bas`, `IMailModel.cls`, `modTheme.bas`, `modTemplateEditor.bas`, `modSetup.bas`, `clsT07.cls`.
4. **Insert > UserForm**, name it `frmT07`, draw the controls listed at the top of `frmT07_code.txt`, and paste that code in.
5. **Debug > Compile**, then run `SetupWorkbook` and `CreateSampleTemplate`.
6. Click **Main Menu** on the Templates sheet and pick T07.

## Adding a template (e.g. T08)

1. Add a row to `tblTemplates`: `T08 | Name | Templates\T08_x.oft | To | CC | Y`.
2. Click **Edit Template**, enter `T08`, write the mail with `{{Tokens}}` and `{{#SECTION}}...{{/SECTION}}`, then click **Save Template**.
3. **Insert > Class Module** named `clsT08`. Copy `clsT07` and change:
   - the properties
   - `TemplateID`
   - `Validate`
   - `Tokens`
   - `Sections`
4. **Insert > UserForm** named `frmT08`. Draw its controls. In `btnCreate_Click`, copy the controls into a `clsT08` and call `ComposeFromModel m`.

## Rules

- A section listed in `Sections()` is kept when its value is True and removed when False.
- A section **not** listed is kept only when all of its tokens are filled. Use this for optional lines.
- Put each section marker on its own line (press Enter, not Shift+Enter).
- Placeholders use letters, digits and `_` only. Hidden formatting Outlook adds inside them is cleaned automatically.
- The mail is always displayed for review and never sent automatically. Classic Outlook is required.
- Put `nostyle` in a control's Tag to keep its own look instead of the theme.
