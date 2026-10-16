# VBA Email Composer - Setup Guide

## Files

| File | What it is | How to add it |
|---|---|---|
| `modMailEngine.bas` | Compose engine: blocks, tokens, Outlook, send log | File > Import |
| `modTemplateRules.bas` | **Fields, visibility and block rules. The only module you edit per template** | File > Import |
| `modValidate.bas` | Template checker (split tokens, block markers, missing fields) | File > Import |
| `modSetup.bas` | One-time setup: sheets, table, buttons, sample templates | File > Import |
| `clsFieldWatch.cls` | Class that reacts to typing in fields | File > Import |
| `clsFormEngine.cls` | Shared logic for every **designed** template form | File > Import |
| `modTheme.bas` | **One place for the look of all forms**: header, footer, colours, fonts, logo | File > Import |
| `modStyles.bas` | **Format painter**: every control copies the look of its sample in `frmStyleMaster` | File > Import |
| `modFormBuilder.bas` | Generates a starter designed form (`frmT05`...) for a template | File > Import |
| `frmCompose_code.txt` | Code for the picker / automatic form (all controls created at runtime) | Paste into a blank UserForm |

No library references are needed. Outlook, RegExp and Dictionary are late-bound.

## Install (about 5 minutes)

1. Create a new workbook and **save it as `.xlsm`** (for example `EmailComposer.xlsm`).
2. Press `Alt+F11` to open the VBA editor.
3. **Insert > UserForm.** In the Properties window (`F4`), set `(Name)` to **`frmCompose`**. Leave the form blank.
4. Double-click the form, delete any code in the window, and paste the full contents of `frmCompose_code.txt`.
5. **File > Import File...** and import, one at a time: `modMailEngine.bas`, `modTemplateRules.bas`, `modValidate.bas`, `modSetup.bas`, `modTheme.bas`, `modStyles.bas`, `modFormBuilder.bas`, `clsFieldWatch.cls`, `clsFormEngine.cls`.
6. **Debug > Compile VBAProject.** It should finish with no errors.
7. Back in Excel, press `Alt+F8` and run **`SetupWorkbook`**. This creates the `Templates`, `SendLog` and `Validation` sheets and adds three buttons.
8. Click **Create Sample Templates** (Outlook must be installed). This writes `Templates\T05_CatchUpUpdate.oft` and `Templates\T06_StatusUpdate.oft` next to the workbook.
9. Click **Validate All Templates**. Both should show `OK`.
10. Click **Compose Email**, pick T05, and try `CatchUpDays` = 45 and then 12.

Do step 3 before step 5. Creating the UserForm adds the *Microsoft Forms 2.0* reference that `clsFieldWatch` needs.

## Designing a separate form per template

Each template can have its own UserForm named **`frm<TemplateID>`** (`frmT05`, `frmT06`...). If one exists, it is used automatically. If not, the automatic stacked form is used, so you can design templates gradually.

### Create the starter form (once per template)

1. One-time: File > Options > Trust Center > Trust Center Settings > Macro Settings > tick **Trust access to the VBA project object model**. Only the person building forms needs this.
2. Run `BuildDesignForm` (`Alt+F8`) and enter `T05`, or run `BuildAllDesignForms` for every template without a form.
3. Open `frmT05` in the VBA editor. It already contains every `lbl_`/`fld_` control, the two buttons and the code.
4. Arrange it however you want: two columns, Frames, MultiPage tabs, different fonts and colours, an image, help text.

### Naming rules (the only rules)

| Control name | Meaning |
|---|---|
| `fld_<Key>` | Input for field `<Key>` (TextBox or ComboBox). The key must match `GetFields`. |
| `lbl_<Key>` | Its caption. Hidden or shown together with the field. Optional. |
| `sec_<Anything>` | A Frame that hides automatically when **all** `fld_` controls inside it are hidden |
| `btnCreate`, `btnCancel` | Buttons (already wired up) |

Everything else is free, including extra labels, pictures and frames not named `sec_`.

### Tip for the catch-up case

Draw a Frame `sec_Long` containing `fld_ApprovalRef` and `fld_ApproverName`, and a Frame `sec_Short` containing `fld_TargetDate`. Place the two frames **on top of each other** at the same position. Only one is ever visible, so the form has no gaps.

### Behaviour

- Positions are never moved at runtime. What you design is what users see. Hidden controls simply disappear.
- Required, visibility and block rules still come from `modTemplateRules`. A designed form changes the **look**, not the rules.
- On opening, the form checks itself and warns about any `fld_` control that is missing or not defined in `GetFields`.
- A designed form can have its own extra code, for example `eng.SetValue "ClaimNo", ActiveCell.Value` in `UserForm_Initialize` to pre-fill a value.
- To open a template directly from a button, create a small macro that calls `ShowTemplate "T05"` and assign it to the button.
- When you add a field later, add it to `GetFields` and draw a `fld_` control for it. The design check reminds you if you forget.

## Common look and feel (modTheme)

Every form calls `ApplyTheme` when it opens, so all forms share the same frame:

```
+-----------------------------------------------+
| Claims Email Composer                  [logo] |  <- header (APP_NAME)
| T05  |  Catch-up Update                       |  <- template ID | DisplayName
|===============================================|  <- accent strip
|                                               |
|   your designed content (fields, frames)      |
|                                               |
|                   [Create Email]  [Cancel]    |
|-----------------------------------------------|
| Fields marked * are required...   v1.0 | user |  <- footer
+-----------------------------------------------+
```

| To change | Edit in `modTheme` |
|---|---|
| Application name, version, footer text | `APP_NAME`, `APP_VERSION`, `FOOTER_TEXT` |
| Colours | the `Clr...` functions (RGB values) |
| Font | `FONT_NAME`, `FONT_SIZE` |
| Header and footer height | `HEADER_H`, `FOOTER_H` |
| Logo | `LOGO_FILE`. Use **.jpg, .bmp or .gif**, because VBA's `LoadPicture` cannot read .png. Leave it `""` for no logo. |

Rules for designers:
- Draw **only the content**. The header is added above it and the form grows to fit, so don't leave empty space for it.
- Don't name controls starting with `thm_`. That prefix is reserved.
- To keep a special look on one control, put `nostyle` in its **Tag**.
- The theme is applied at runtime. In the VBA designer you still see your plain layout, and the themed look appears when the form runs.

## Consistent control styles (modStyles + frmStyleMaster)

Design one sample of each control type **once**, and every label, combo box, button and frame on every form copies it.

### Set up

1. Run `BuildStyleMaster` (needs *Trust access to the VBA project object model*). It creates `frmStyleMaster` with one sample per style, pre-set to the current theme colours.
2. Open `frmStyleMaster` in the VBA designer and format the samples with the Properties window: font, size, bold, ForeColor, BackColor, BorderColor, SpecialEffect, TextAlign.
3. Done. Every form copies the samples when it opens. If Excel was already open while you edited the master, run `ResetStyleCache` or reopen the workbook.
4. Optional: run `PaintAllTemplateForms` to write the styles **into the designs** too, so the VBA designer shows the same look. Then save the workbook.

### Samples and who uses them

| Sample in frmStyleMaster | Applied to |
|---|---|
| `sty_Label` | every Label |
| `sty_RequiredLabel` | Labels whose caption ends with `*` |
| `sty_TextBox`, `sty_ComboBox`, `sty_ListBox` | every control of that type |
| `sty_Frame`, `sty_CheckBox`, `sty_OptionButton` | every control of that type |
| `sty_PrimaryButton` | `btnCreate` |
| `sty_CommandButton` | all other buttons |
| `sty_Header`, `sty_AppName`, `sty_TemplateName`, `sty_Accent` | the header. The header **height** follows `sty_Header`. |
| `sty_Footer`, `sty_FooterText` | the footer. The footer **height** follows `sty_Footer`. |
| the form's own BackColor | the background of every form |
| **custom** `sty_<Name>` (for example `sty_Heading`) | any control whose Tag contains `style:<Name>` |

### What is copied

| Copied | Not copied |
|---|---|
| Font name, size, bold, italic, underline; ForeColor, BackColor, BackStyle, BorderStyle, BorderColor, SpecialEffect, TextAlign | Position, width, caption, list values and behaviour, so your layouts stay intact |

Height is copied only if you set `COPY_HEIGHT = True` in `modStyles`. It applies to single-line text boxes, combo boxes and buttons.

### Priority

The **`nostyle` Tag** keeps a control's own look. Otherwise **`style:<Name>` Tag** wins, then the type sample, then the coded colours in `modTheme`, which are used when no master exists.

### Adding your own style

To add a style, for example a red warning label:
1. Add a Label named `sty_Warning` to `frmStyleMaster` and format it.
2. On any form, set a label's Tag to `style:Warning`.

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
