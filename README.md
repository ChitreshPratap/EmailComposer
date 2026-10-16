# VBA Email Composer - Setup Guide

## Files

| File | What it is | How to add it |
|---|---|---|
| `modMailEngine.bas` | Compose engine: blocks, tokens, Outlook, send log | File > Import |
| `modTemplateRules.bas` | **Fields, visibility and block rules. The only module you edit per template** | File > Import |
| `modValidate.bas` | Template checker (split tokens, block markers, missing fields) | File > Import |
| `modSetup.bas` | One-time setup: sheets, table, buttons, sample templates | File > Import |
| `clsFieldWatch.cls` | Class that reacts to typing in runtime-created fields | File > Import |
| `frmCompose_code.txt` | Code for the UserForm (all controls are created at runtime) | Paste into a blank UserForm |

No library references are needed. Outlook, RegExp and Dictionary are late-bound.

## Install (about 5 minutes)

1. Create a new workbook and **save it as `.xlsm`** (for example `EmailComposer.xlsm`).
2. Press `Alt+F11` to open the VBA editor.
3. **Insert > UserForm.** In the Properties window (`F4`), set `(Name)` to **`frmCompose`**. Leave the form blank.
4. Double-click the form, delete any code in the window, and paste the full contents of `frmCompose_code.txt`.
5. **File > Import File...** and import, one at a time: `modMailEngine.bas`, `modTemplateRules.bas`, `modValidate.bas`, `modSetup.bas`, `clsFieldWatch.cls`.
6. **Debug > Compile VBAProject.** It should finish with no errors.
7. Back in Excel, press `Alt+F8` and run **`SetupWorkbook`**. This creates the `Templates`, `SendLog` and `Validation` sheets and adds three buttons.
8. Click **Create Sample Templates** (Outlook must be installed). This writes `Templates\T05_CatchUpUpdate.oft` and `Templates\T06_StatusUpdate.oft` next to the workbook.
9. Click **Validate All Templates**. Both should show `OK`.
10. Click **Compose Email**, pick T05, and try `CatchUpDays` = 45 and then 12.

Do step 3 before step 5. Creating the UserForm adds the *Microsoft Forms 2.0* reference that `clsFieldWatch` needs.

## Template markup rules (for whoever edits the .oft)

| Element | Syntax | Rule |
|---|---|---|
| Value | `{{ClaimNo}}` | No spaces. Must match a field key in `GetFields`. |
| Optional / conditional section | `{{#NAME}}` ... `{{/NAME}}` | Each marker **on its own line** (press Enter, not Shift+Enter). Each name used once. |
| Tokens in subject | `Claim {{ClaimNo}}` | Supported. |
| Tokens in To / CC | in `tblTemplates`, e.g. `{{CustomerEmail}}` | Supported. |

If you edit a placeholder in Outlook, **delete it and retype it in one go**. Formatting part of a token (making half of it bold, autocorrect, and so on) splits it in the HTML. **Always run Validate after editing.**

### How a section decides to stay or go

| Situation | Behaviour |
|---|---|
| Block has a rule in `BlockRule` | Kept only when the rule returns True |
| Block has **no** rule (AUTO) | Kept only when **every** `{{token}}` inside it is filled |

## Adding template #41

1. Create the email in Outlook, put in the tokens and blocks, then **File > Save As > Outlook Template (.oft)** into the `Templates` folder.
2. Add a row to `tblTemplates`: `T41 | Name | Templates\T41_xxx.oft | To | CC | Y`.
3. In `modTemplateRules.GetFields`, add a `Case "T41"` with its `AddF` lines.
4. Only if needed: add visibility rules in `IsFieldVisible` and block rules in `BlockRule`.
5. Run **Validate All Templates**.

`FilePath` can be relative to the workbook (`Templates\x.oft`), a full path (`D:\...`) or a UNC path (`\\server\share\...`).

## Behaviour of the T05 sample

| Input | Form shows | Email contains |
|---|---|---|
| CatchUpDays = 45 | Approval Ref, Approver Name | LONG_CATCHUP paragraph |
| CatchUpDays = 12 | Target Date | SHORT_CATCHUP paragraph |
| CatchUpDays = 30 | Target Date | SHORT_CATCHUP (change `<= 30` in `IsShortCatchUp` if needed) |
| Escalation Ref left blank | field still shown (optional) | ESCALATION paragraph removed |

## Known limits

- The mail is always **displayed**, never sent automatically.
- `CreateItemFromTemplate` does not add your Outlook signature. Put the signature text in the `.oft` itself.
- Section removal works on normal paragraphs. A block inside a table cell clears the text but leaves the empty cell or row.
- Protect the VBA project (VBA editor > Tools > VBAProject Properties > Protection) before distributing, so users can edit the `.oft` wording but not the rules.
