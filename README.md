# Excel Formula Formatter

Give complex Excel formulas the same polish you expect from code. This Visual Studio Code extension keeps formulas readable, helps you explore function names, and now lets you collapse everything back down when you need compact copies.

## Why You'll Like It

- **Clean formatting** – Convert dense formulas into a well-indented layout that is easy to scan.
- **One-click minify** – Run `Excel Formula: Minify Document or Selection` from the Command Palette when you need a single-line version.
- **Real-time IntelliSense** – Start typing and get completions for common Excel functions with signature help.
- **Syntax-aware highlighting** – Colours functions, cell references, strings, booleans, and error literals so the important bits stand out.
- **File explorer integration** – `.xlf` files show a dedicated Excel Formula icon for quick visual recognition.

## Install

1. Open the Extensions view in VS Code (`Ctrl+Shift+X` / `Cmd+Shift+X`).
2. Search for **"Excel Formula Formatter"** and click **Install**.
3. Reload the window if prompted and you're ready to format formulas.

> Tip: If you're testing a local build, use **Extensions: Install from VSIX...** and select the generated package.
}
## Everyday Use

1. Save or open a file with the `.xlf` extension (or pick **Excel Formula** in the language mode picker).
2. Paste or type your formula. The extension formats on paste automatically; you can also run **Format Document** at any time.
3. Need the compact version? Press `Ctrl+Shift+P` / `Cmd+Shift+P`, run **Excel Formula: Minify Document or Selection**, and copy the result.
4. Trigger completions with `Ctrl+Space` to browse function snippets.

## Commands

| Command | Purpose |
| ------- | ------- |
| `Excel Formula: Minify Document or Selection` | Collapse the active selection or entire file into a single-line formula while preserving strings and operators. |
| `Excel Formula: Beautify Document or Selection` | Manually beautify the selection ot the entire file. |

## Configuration

The extension ships with sensible defaults for Excel formula editing. Override them in your `settings.json` if you prefer a different setup:

```json
"[excel-formula]": {
  "editor.tabSize": 2,
  "editor.insertSpaces": true,
  "editor.formatOnPaste": true
}
```

## Known Limitations

- Array constants are expanded one element per line. Alternate layouts are planned.
- Only built-in Excel functions are suggested today; custom workbook functions are not yet included.

## Share Feedback

Spotted a bug or want a new capability? Open an issue in the repository and include a sample formula so we can reproduce it quickly.

## For Developers

Want to tinker with the formatter logic or help the project grow? Here's how to get started:

1. Clone the repository and install dependencies:
   ```bash
   git clone https://github.com/<your-org>/excel-formula-formatter.git
   cd excel-formula-formatter
   npm install
   ```
2. Open the folder in VS Code and press `F5` to launch the extension in a new Extension Development Host window.
3. Run sanity checks with `node run-format-tests.js`. Add new samples under `test/` before submitting a PR.
4. Package a distributable build via `npx vsce package`.

We welcome issues, feature ideas, and pull requests—please describe the scenario and include example formulas so reviews go quickly.

Enjoy cleaner Excel formulas without leaving VS Code!
