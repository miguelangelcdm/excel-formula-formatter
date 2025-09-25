const vscode = require('vscode');

// Function metadata lives in a separate module for easier maintenance.
const FUNCTION_ITEMS = require('./function-items');
const FUNCTION_LOOKUP = new Map(
  FUNCTION_ITEMS.map(item => [item.name.toUpperCase(), item])
);

// Entry point for the VS Code extension lifecycle.
function activate(context) {
  const selector = { language: 'excel-formula', scheme: '*' };

  // Autocomplete provider built from FUNCTION_ITEMS above.
  const completionProvider = vscode.languages.registerCompletionItemProvider(
    selector,
    {
      provideCompletionItems(document, position) {
        const items = FUNCTION_ITEMS.map(fn => {
          const item = new vscode.CompletionItem(fn.name, vscode.CompletionItemKind.Function);
          item.detail = fn.signature;
          item.documentation = fn.detail;
          item.insertText = new vscode.SnippetString(`${fn.name}($0)`);
          item.commitCharacters = ['('];
          return item;
        });
        return items;
      }
    },
    '('
  );

  // Full-document formatter invoked by the built-in Format Document command.
  const formattingProvider = vscode.languages.registerDocumentFormattingEditProvider(selector, {
    provideDocumentFormattingEdits(document) {
      const text = document.getText();
      const formatted = formatFormula(text);
      const fullRange = new vscode.Range(
        document.positionAt(0),
        document.positionAt(text.length)
      );
      if (formatted === text) {
        return [];
      }
      const ensureTrailingNewline = /\r?\n$/.test(text);
      const finalText = ensureTrailingNewline && !/\r?\n$/.test(formatted) ? `${formatted}\n` : formatted;
      return [vscode.TextEdit.replace(fullRange, finalText)];
    }
  });

  // Range formatter for Format Selection scenarios.
  const rangeFormattingProvider = vscode.languages.registerDocumentRangeFormattingEditProvider(selector, {
    provideDocumentRangeFormattingEdits(document, range) {
      const text = document.getText(range);
      const formatted = formatFormula(text);
      if (formatted === text) {
        return [];
      }
      return [vscode.TextEdit.replace(range, formatted)];
    }
  });

  const hoverProvider = vscode.languages.registerHoverProvider(selector, {
    provideHover(document, position) {
      const wordRange = document.getWordRangeAtPosition(position, /[A-Za-z][A-Za-z0-9_\.]*/);
      if (!wordRange) {
        return null;
      }

      const functionName = document.getText(wordRange).toUpperCase();
      const metadata = FUNCTION_LOOKUP.get(functionName);
      if (!metadata) {
        return null;
      }

      const markdown = new vscode.MarkdownString();
      markdown.appendCodeblock(metadata.signature, 'excel-formula');
      if (metadata.detail) {
        markdown.appendMarkdown(`\n\n${metadata.detail}`);
      }

      return new vscode.Hover(markdown, wordRange);
    }
  });

  // Command palette entry to collapse formulas to a compact representation.
  const minifyCommand = vscode.commands.registerCommand('excel-formula.minifyFormula', async () => {
    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document.languageId !== 'excel-formula') {
      vscode.window.showInformationMessage('Open an Excel Formula document before running minify.');
      return;
    }

    await transformDocumentOrSelection(editor, minifyFormula);
  });

  const beautifyCommand = vscode.commands.registerCommand('excel-formula.beautifyFormula', async () => {
    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document.languageId !== 'excel-formula') {
      vscode.window.showInformationMessage('Open an Excel Formula document before running beautify.');
      return;
    }

    await transformDocumentOrSelection(editor, formatFormula);
  });

  context.subscriptions.push(
    completionProvider,
    formattingProvider,
    rangeFormattingProvider,
    hoverProvider,
    minifyCommand,
    beautifyCommand
  );
}

function deactivate() {}

// Friendly formatter that breaks multi-line formulas into readable chunks.
function formatFormula(input) {
  if (!input.trim()) {
    return input;
  }
  const lines = input.split(/\r?\n/);
  const groups = [];
  let current = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      if (current.length) {
        groups.push(current);
        current = [];
      }
      groups.push(null);
      continue;
    }
    if (trimmed.startsWith('=')) {
      if (current.length) {
        groups.push(current);
      }
      current = [trimmed];
    } else {
      current.push(trimmed);
    }
  }

  if (current.length) {
    groups.push(current);
  }

  while (groups.length && groups[groups.length - 1] === null) {
    groups.pop();
  }

  const formattedLines = [];

  groups.forEach(section => {
    if (section === null) {
      if (formattedLines.length && formattedLines[formattedLines.length - 1] !== '') {
        formattedLines.push('');
      }
      return;
    }
    const formula = section.join(' ');
    const formatted = formatSingleFormula(formula);
    if (!formatted) {
      return;
    }
    formatted.split('\n').forEach(line => formattedLines.push(line));
  });

  return formattedLines.join('\n');
}

function minifyFormula(input) {
  if (!input.trim()) {
    return input;
  }

  const lines = input.split(/\r?\n/);
  const groups = [];
  let current = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      if (current.length) {
        groups.push(current);
        current = [];
      }
      groups.push(null);
      continue;
    }
    if (trimmed.startsWith('=')) {
      if (current.length) {
        groups.push(current);
      }
      current = [trimmed];
    } else {
      current.push(trimmed);
    }
  }

  if (current.length) {
    groups.push(current);
  }

  while (groups.length && groups[groups.length - 1] === null) {
    groups.pop();
  }

  const minifiedLines = [];

  groups.forEach(section => {
    if (section === null) {
      if (minifiedLines.length && minifiedLines[minifiedLines.length - 1] !== '') {
        minifiedLines.push('');
      }
      return;
    }
    const formula = section.join(' ');
    const minified = minifySingleFormula(formula);
    if (!minified) {
      return;
    }
    minifiedLines.push(minified);
  });

  return minifiedLines.join('\n');
}

// Removes optional whitespace while retaining separation between words/strings.
function minifySingleFormula(formula) {
  const tokens = tokenize(formula);
  if (!tokens.length) {
    return formula.trim();
  }

  let result = '';
  let lastToken = null;

  tokens.forEach(token => {
    if (token.type === 'operator' || token.type === 'punct') {
      if (token.value === ')' || token.value === '}') {
        result = result.trimEnd();
      }
      result += token.value;
      lastToken = token;
      return;
    }

    if (token.type === 'word') {
      if (lastToken && (lastToken.type === 'word' || lastToken.type === 'string')) {
        result += ' ';
      }
      result += token.value;
      lastToken = token;
      return;
    }

    if (token.type === 'string') {
      if (lastToken && (lastToken.type === 'word' || lastToken.type === 'string')) {
        result += ' ';
      }
      result += token.value;
      lastToken = token;
    }
  });

  return result.trim();
}

// Pretty-printer responsible for indentation and punctuation placement.
function formatSingleFormula(formula) {
  const tokens = tokenize(formula);
  if (!tokens.length) {
    return formula.trim();
  }
  const indentUnit = '  ';
  let indentLevel = 0;
  const lines = [];
  let currentLine = '';
  let lastToken = null;

  const ensureIndent = () => {
    if (!currentLine) {
      currentLine = indentUnit.repeat(indentLevel);
    }
  };

  const ensureSpaceBefore = () => {
    if (!currentLine) {
      ensureIndent();
      return;
    }
    if (/\s$/.test(currentLine) || currentLine.endsWith('(') || currentLine.endsWith('{')) {
      return;
    }
    currentLine += ' ';
  };

  const closeLine = () => {
    lines.push(currentLine.trimEnd());
    currentLine = '';
  };

  tokens.forEach((token, index) => {
    if (token.type === 'punct') {
      if (token.value === '(' || token.value === '{') {
        currentLine = currentLine.trimEnd();
        ensureIndent();
        currentLine += token.value;
        closeLine();
        indentLevel += 1;
      } else if (token.value === ')' || token.value === '}') {
        if (currentLine.trim().length) {
          closeLine();
        }
        indentLevel = Math.max(indentLevel - 1, 0);
        currentLine = indentUnit.repeat(indentLevel) + token.value;
      } else if (token.value === ',' || token.value === ';') {
        currentLine = currentLine.trimEnd();
        ensureIndent();
        currentLine += token.value;
        closeLine();
      } else {
        ensureIndent();
        currentLine += token.value;
      }
      lastToken = token;
      return;
    }

    if (token.type === 'operator') {
      const atLineStart = lines.length === 0 && !currentLine.trim();
      if (token.value === '=' && atLineStart) {
        ensureIndent();
        currentLine += '=';
        lastToken = token;
        return;
      }
      const isUnary = (token.value === '-' || token.value === '+') && (!lastToken || lastToken.type === 'operator' || (lastToken.type === 'punct' && ['(', '{', ',', ';'].includes(lastToken.value)));
      ensureIndent();
      if (isUnary) {
        currentLine += token.value;
      } else {
        ensureSpaceBefore();
        currentLine += token.value;
        currentLine += ' ';
      }
      lastToken = token;
      return;
    }

    if (token.type === 'string' || token.type === 'word') {
      ensureIndent();
      if (currentLine && !/\s$/.test(currentLine) && !currentLine.endsWith('(') && !currentLine.endsWith('{') && !currentLine.endsWith('=')) {
        currentLine += ' ';
      }
      currentLine += token.value;
      lastToken = token;
      return;
    }
  });

  if (currentLine.trim().length) {
    closeLine();
  }

  return lines.join('\n');
}

// Shared tokenizer so formatting and minify stay in sync.
function tokenize(source) {
  const tokens = [];
  let buffer = '';
  let inString = false;

  const flushBuffer = () => {
    if (!buffer) {
      return;
    }
    tokens.push({ type: 'word', value: buffer });
    buffer = '';
  };

  for (let i = 0; i < source.length; i += 1) {
    const ch = source[i];

    if (inString) {
      buffer += ch;
      if (ch === '"') {
        if (i + 1 < source.length && source[i + 1] === '"') {
          buffer += source[i + 1];
          i += 1;
          continue;
        }
        tokens.push({ type: 'string', value: buffer });
        buffer = '';
        inString = false;
      }
      continue;
    }

    if (ch === '"') {
      flushBuffer();
      buffer = '"';
      inString = true;
      continue;
    }

    if (/\s/.test(ch)) {
      flushBuffer();
      continue;
    }

    const twoChar = source.slice(i, i + 2);
    if (twoChar === '>=' || twoChar === '<=' || twoChar === '<>') {
      flushBuffer();
      tokens.push({ type: 'operator', value: twoChar });
      i += 1;
      continue;
    }

    if (',;(){}'.includes(ch)) {
      flushBuffer();
      tokens.push({ type: 'punct', value: ch });
      continue;
    }

    if ('+-*/^&=%'.includes(ch) || ['=', '>', '<'].includes(ch)) {
      flushBuffer();
      tokens.push({ type: 'operator', value: ch });
      continue;
    }

    buffer += ch;
  }

  if (buffer) {
    tokens.push({ type: 'word', value: buffer });
  }

  return tokens;
}

async function transformDocumentOrSelection(editor, transformFn) {
  const document = editor.document;
  const selections = editor.selections;
  const newlinePattern = /\r?\n$/;

  const hasSelection = selections.some(selection => !selection.isEmpty);
  if (hasSelection) {
    await editor.edit(editBuilder => {
      selections.forEach(selection => {
        if (selection.isEmpty) {
          return;
        }
        const original = document.getText(selection);
        const transformed = transformFn(original);
        if (transformed !== original) {
          editBuilder.replace(selection, transformed);
        }
      });
    });
    return;
  }

  const text = document.getText();
  const transformed = transformFn(text);
  if (transformed === text) {
    return;
  }

  const ensureTrailingNewline = newlinePattern.test(text);
  const finalText = ensureTrailingNewline && !newlinePattern.test(transformed)
    ? `${transformed}\n`
    : transformed;

  const fullRange = new vscode.Range(
    document.positionAt(0),
    document.positionAt(text.length)
  );

  await editor.edit(editBuilder => {
    editBuilder.replace(fullRange, finalText);
  });
}

module.exports = {
  activate,
  deactivate,
  _formatFormula: formatFormula,
  _minifyFormula: minifyFormula,
};
