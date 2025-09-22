# Excel Formula Formatter

Excel Formula Formatter adds a purpose-built editing experience for Excel formulas inside VS Code. It keeps pasted formulas readable, surfaces the built-in functions you expect, and highlights the pieces that matter in complex expressions.

## Features

- Syntax-aware highlighting for functions, cell references, structured references, booleans, and error literals.
- IntelliSense completions for common Excel functions with quick signature reminders and callable snippets.
- Document and range formatting that expands compact formulas into an indented, easy-to-scan layout.
- Smart indentation rules so pressing <kbd>Enter</kbd> inside parentheses keeps nested expressions aligned.
- Ships with sensible editor defaults (`insertSpaces`, two-space indents, and `formatOnPaste`) scoped to the `excel-formula` language ID.

## Getting Started

1. Save formulas in a file with the `.xlf` extension or select **Excel Formula** as the language mode.
2. Paste or type your formula. The formatter will run automatically on paste (or invoke **Format Document** manually).
3. Trigger completions with <kbd>Ctrl</kbd>+<kbd>Space</kbd> to browse common Excel functions and insert ready-to-edit call templates.

## Formatting Behaviour

The formatter normalises whitespace, places each argument on its own line, and indents based on nesting depth for parentheses and array literals. It preserves string contents and range references, and is idempotent—reformatting an already formatted block leaves it unchanged.

## Extension Settings

The extension contributes configuration defaults under the `[excel-formula]` scope. Override them in your settings if you prefer different indentation or paste behaviour.

```
"[excel-formula]": {
  "editor.tabSize": 2,
  "editor.insertSpaces": true,
  "editor.formatOnPaste": true
}
```

## Known Limitations

- Array constants are expanded one element per line. Future updates may add configurable layout styles.
- The completion list focuses on the most common built-in functions. Custom names and add-in functions are not yet suggested.

## Release Notes

### 0.0.1

- Added syntax highlighting, completions, indentation rules, and a paste-friendly formatter for Excel formulas.
