# Office.js Custom Functions #NAME? Repro

Minimal reproduction for a bug where Excel does not re-resolve `#NAME?` custom function formulas when the add-in that provides them is loaded after the formulas are already entered.

## Setup

```bash
npm install

# Install mkcert and generate trusted localhost certs (required for Office add-ins)
brew install mkcert
mkcert -install
mkdir -p certs
mkcert -cert-file certs/localhost.pem -key-file certs/localhost-key.pem localhost

npm start
```

This starts an HTTPS dev server at `https://localhost:3000`. Verify it's working by opening `https://localhost:3000/functions.json` in your browser — it should load without a certificate warning.

Then sideload `manifest.xml` into Excel. The add-in registers two custom functions:

- `CONTOSO.ADD(a, b)` — adds two numbers
- `CONTOSO.GET(value)` — returns the input string

## Test Workbook

The included `Custom Functions Protected Sheet Repro.xlsx` has three sheets, all with `=CONTOSO.ADD(1,2)` in A1 and `=CONTOSO.GET("test")` in B1:

- **Protected** — protected (password: `password`), formulas stuck as `#NAME?`
- **Unprotected** — unprotected, formulas stuck as `#NAME?`
- **Correct** — unprotected, formulas properly resolved (entered with add-in loaded)

## Underlying XML

Unzipping the `.xlsx` reveals the difference. When formulas are stuck (`Protected` and `Unprotected` sheets), the XML stores them as-entered with a cached `#NAME?` error:

```xml
<c r="A1" t="e" cm="1">
  <f t="array" aca="1" ref="A1" ca="1">CONTOSO.ADD(1,2)</f>
  <v>#NAME?</v>
</c>
<c r="B1" t="e" cm="1">
  <f t="array" aca="1" ref="B1" ca="1">CONTOSO.GET("test")</f>
  <v>#NAME?</v>
</c>
```

When formulas are properly resolved (`Correct` sheet), Excel rewrites them with a `_xldudf_` prefix and stores the computed value:

```xml
<c r="A1" cm="1">
  <f t="array" ref="A1">_xldudf_CONTOSO_ADD(1,2)</f>
  <v>3</v>
</c>
<c r="B1" t="str" cm="1">
  <f t="array" ref="B1">_xldudf_CONTOSO_GET("test")</f>
  <v>test</v>
</c>
```

The `_xldudf_` (Excel User Defined Function) prefix indicates Excel has recognized the function as belonging to a custom functions add-in. The stuck formulas never get rewritten to this form — even after the add-in loads and a recalculation is forced.

## Reproduction Steps

### 1. Confirm the add-in works

1. Open a new workbook with the add-in already sideloaded
2. Enter `=CONTOSO.ADD(1,2)` in any cell
3. It should return `3`

### 2. Reproduce the bug (unprotected sheet)

1. Remove/unload the add-in
2. Open the included `Custom Functions Protected Sheet Repro.xlsx` (or create a new workbook and enter the formulas manually)
3. The formulas show `#NAME?` (expected, add-in not loaded)
4. Sideload the add-in
5. **Expected:** The formulas resolve
6. **Actual:** The formulas stay as `#NAME?`
7. "Calculate Sheet" and force-recalculate (Cmd+Shift+Alt+F9) do not help
8. Clicking into the cell and pressing Enter *does* resolve it — but this must be done for every affected cell individually

### 3. Reproduce the bug (protected sheet — no workaround)

1. Follow the same steps on Sheet1 (protected, password: `password`)
2. Sideload the add-in
3. The formulas remain `#NAME?` with no way to resolve them — you cannot click into cells to re-enter them on a protected sheet

## Impact

This makes it impossible to migrate users from a legacy add-in (e.g., a COM/XLL add-in) to an Office.js custom functions add-in when workbooks contain pre-existing formulas, especially on protected sheets where the manual re-entry workaround is unavailable.
