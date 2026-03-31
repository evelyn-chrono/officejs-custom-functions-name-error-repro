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

## Reproduction Steps

### 1. Confirm the add-in works

1. Open a new workbook with the add-in already sideloaded
2. Enter `=CONTOSO.ADD(1,2)` in any cell
3. It should return `3`

### 2. Reproduce the bug (unprotected sheet)

1. Remove/unload the add-in
2. Open a new workbook (or use the included `Custom Functions Protected Sheet Repro.xlsx`)
3. Enter `=CONTOSO.GET("hello")` in a cell — it shows `#NAME?` (expected, add-in not loaded)
4. Sideload the add-in
5. **Expected:** The formula resolves to `hello`
6. **Actual:** The formula stays as `#NAME?`
7. "Calculate Sheet" and force-recalculate (Cmd+Shift+Alt+F9) do not help
8. Clicking into the cell and pressing Enter *does* resolve it — but this must be done for every affected cell individually

### 3. Reproduce the bug (protected sheet — no workaround)

1. Follow steps 1–3 above
2. Protect the sheet (Review > Protect Sheet)
3. Sideload the add-in
4. The formulas remain `#NAME?` with no way to resolve them — you cannot click into cells to re-enter them on a protected sheet

## Impact

This makes it impossible to migrate users from a legacy add-in (e.g., a COM/XLL add-in) to an Office.js custom functions add-in when workbooks contain pre-existing formulas, especially on protected sheets where the manual re-entry workaround is unavailable.
