const vscode = require('vscode');

const FUNCTION_ITEMS = [
  { name: 'ABS', signature: 'ABS(number)', detail: 'Return the absolute value of a number.' },
  { name: 'ACOS', signature: 'ACOS(number)', detail: 'Return the arccosine of a number.' },
  { name: 'ACOSH', signature: 'ACOSH(number)', detail: 'Return the inverse hyperbolic cosine of a number.' },
  { name: 'ADDRESS', signature: 'ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])', detail: 'Create a cell reference as text.' },
  { name: 'AND', signature: 'AND(logical1, [logical2], ...)', detail: 'Return TRUE if all arguments are TRUE.' },
  { name: 'ARRAYTOTEXT', signature: 'ARRAYTOTEXT(array, [format])', detail: 'Return an array of text values from the provided array.' },
  { name: 'AVERAGE', signature: 'AVERAGE(number1, [number2], ...)', detail: 'Average the numbers provided as arguments.' },
  { name: 'AVERAGEIF', signature: 'AVERAGEIF(range, criteria, [average_range])', detail: 'Average the cells that meet a single criterion.' },
  { name: 'AVERAGEIFS', signature: 'AVERAGEIFS(average_range, criteria_range1, criteria1, ...)', detail: 'Average the cells that meet multiple criteria.' },
  { name: 'BYROW', signature: 'BYROW(array, lambda)', detail: 'Apply a LAMBDA to each row and return an array of results.' },
  { name: 'BYCOL', signature: 'BYCOL(array, lambda)', detail: 'Apply a LAMBDA to each column and return an array of results.' },
  { name: 'CEILING', signature: 'CEILING(number, significance)', detail: 'Round a number up to the nearest multiple of significance.' },
  { name: 'CHOOSE', signature: 'CHOOSE(index_num, value1, [value2], ...)', detail: 'Pick a value from the list of arguments.' },
  { name: 'CHOOSECOLS', signature: 'CHOOSECOLS(array, col_num1, [col_num2], ...)', detail: 'Return specific columns from an array.' },
  { name: 'CHOOSEROWS', signature: 'CHOOSEROWS(array, row_num1, [row_num2], ...)', detail: 'Return specific rows from an array.' },
  { name: 'COLUMN', signature: 'COLUMN([reference])', detail: 'Return the column number of a reference.' },
  { name: 'COLUMNS', signature: 'COLUMNS(array)', detail: 'Return the number of columns in an array.' },
  { name: 'COUNT', signature: 'COUNT(value1, [value2], ...)', detail: 'Count the number of cells that contain numbers.' },
  { name: 'COUNTA', signature: 'COUNTA(value1, [value2], ...)', detail: 'Count the number of non-empty cells.' },
  { name: 'COUNTBLANK', signature: 'COUNTBLANK(range)', detail: 'Count the number of empty cells in a range.' },
  { name: 'COUNTIF', signature: 'COUNTIF(range, criteria)', detail: 'Count cells that meet a single criterion.' },
  { name: 'COUNTIFS', signature: 'COUNTIFS(criteria_range1, criteria1, ...)', detail: 'Count cells that meet multiple criteria.' },
  { name: 'DATE', signature: 'DATE(year, month, day)', detail: 'Return the serial number of a particular date.' },
  { name: 'DATEDIF', signature: 'DATEDIF(start_date, end_date, unit)', detail: 'Calculate the difference between two dates.' },
  { name: 'DATEVALUE', signature: 'DATEVALUE(date_text)', detail: 'Convert a date in text to a serial number.' },
  { name: 'DROP', signature: 'DROP(array, rows, [columns])', detail: 'Drop rows or columns from an array.' },
  { name: 'EDATE', signature: 'EDATE(start_date, months)', detail: 'Return a date moved forward or backward by a number of months.' },
  { name: 'EOMONTH', signature: 'EOMONTH(start_date, months)', detail: 'Return the last day of the month before or after the start date.' },
  { name: 'EXPAND', signature: 'EXPAND(array, rows, columns, [pad_with])', detail: 'Expand an array to specified dimensions.' },
  { name: 'FILTER', signature: 'FILTER(array, include, [if_empty])', detail: 'Filter a range with the provided criteria.' },
  { name: 'IF', signature: 'IF(logical_test, value_if_true, [value_if_false])', detail: 'Check a condition and return a value if TRUE or FALSE.' },
  { name: 'IFERROR', signature: 'IFERROR(value, value_if_error)', detail: 'Return a value if there is an error, otherwise return the original value.' },
  { name: 'IFS', signature: 'IFS(logical_test1, value_if_true1, ...)', detail: 'Evaluate multiple conditions and return the first TRUE result.' },
  { name: 'INDEX', signature: 'INDEX(array, row_num, [column_num], [area_num])', detail: 'Return a value or reference of the cell at the intersection of a row and column.' },
  { name: 'INDIRECT', signature: 'INDIRECT(ref_text, [a1])', detail: 'Return the reference specified by a text string.' },
  { name: 'INT', signature: 'INT(number)', detail: 'Round a number down to the nearest integer.' },
  { name: 'ISBLANK', signature: 'ISBLANK(value)', detail: 'Return TRUE if value refers to an empty cell.' },
  { name: 'ISNUMBER', signature: 'ISNUMBER(value)', detail: 'Return TRUE if value is a number.' },
  { name: 'LAMBDA', signature: 'LAMBDA(parameter, calculation)', detail: 'Create reusable custom functions inline.' },
  { name: 'LET', signature: 'LET(name1, value1, calculation_or_name2, ...)', detail: 'Assign names to calculation results and reuse them.' },
  { name: 'MATCH', signature: 'MATCH(lookup_value, lookup_array, [match_type])', detail: 'Return the relative position of an item in an array.' },
  { name: 'MAX', signature: 'MAX(number1, [number2], ...)', detail: 'Return the largest number in the set.' },
  { name: 'MAXIFS', signature: 'MAXIFS(max_range, criteria_range1, criteria1, ...)', detail: 'Return the maximum value that satisfies multiple conditions.' },
  { name: 'MIN', signature: 'MIN(number1, [number2], ...)', detail: 'Return the smallest number in the set.' },
  { name: 'MINIFS', signature: 'MINIFS(min_range, criteria_range1, criteria1, ...)', detail: 'Return the minimum value that satisfies multiple conditions.' },
  { name: 'MMULT', signature: 'MMULT(array1, array2)', detail: 'Return the matrix product of two arrays.' },
  { name: 'OFFSET', signature: 'OFFSET(reference, rows, cols, [height], [width])', detail: 'Return a reference offset from a given reference.' },
  { name: 'OR', signature: 'OR(logical1, [logical2], ...)', detail: 'Return TRUE if any argument is TRUE.' },
  { name: 'RANDARRAY', signature: 'RANDARRAY([rows], [columns], [min], [max], [whole_number])', detail: 'Return an array of random numbers.' },
  { name: 'REDUCE', signature: 'REDUCE(initial_value, array, lambda)', detail: 'Reduce an array to an accumulated value via a LAMBDA.' },
  { name: 'ROUND', signature: 'ROUND(number, num_digits)', detail: 'Round a number to the specified number of digits.' },
  { name: 'SCAN', signature: 'SCAN(initial_value, array, lambda)', detail: 'Return an array of intermediate values via a LAMBDA.' },
  { name: 'SEQUENCE', signature: 'SEQUENCE(rows, [columns], [start], [step])', detail: 'Generate an array of sequential numbers.' },
  { name: 'SORT', signature: 'SORT(array, [sort_index], [sort_order], [by_col])', detail: 'Sort the contents of a range or array.' },
  { name: 'SORTBY', signature: 'SORTBY(array, by_array1, [sort_order1], ...)', detail: 'Sort a range or array by corresponding values in another range.' },
  { name: 'SUM', signature: 'SUM(number1, [number2], ...)', detail: 'Add all the numbers in a range of cells.' },
  { name: 'SUMIF', signature: 'SUMIF(range, criteria, [sum_range])', detail: 'Add the cells that meet a single criterion.' },
  { name: 'SUMIFS', signature: 'SUMIFS(sum_range, criteria_range1, criteria1, ...)', detail: 'Add the cells that meet multiple criteria.' },
  { name: 'SWITCH', signature: 'SWITCH(expression, value1, result1, [default_or_value2, result2], ...)', detail: 'Evaluate an expression against a list of values.' },
  { name: 'TAKE', signature: 'TAKE(array, rows, [columns])', detail: 'Return a specified number of contiguous rows or columns.' },
  { name: 'TEXT', signature: 'TEXT(value, format_text)', detail: 'Format a number and convert it to text.' },
  { name: 'TEXTAFTER', signature: 'TEXTAFTER(text, delimiter, [instance], ...)', detail: 'Return text that occurs after a given delimiter.' },
  { name: 'TEXTBEFORE', signature: 'TEXTBEFORE(text, delimiter, [instance], ...)', detail: 'Return text that occurs before a given delimiter.' },
  { name: 'TEXTJOIN', signature: 'TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)', detail: 'Join text items with a delimiter.' },
  { name: 'TRANSPOSE', signature: 'TRANSPOSE(array)', detail: 'Return the transpose of an array.' },
  { name: 'TRIM', signature: 'TRIM(text)', detail: 'Remove spaces from text except for single spaces between words.' },
  { name: 'UNIQUE', signature: 'UNIQUE(array, [by_col], [exactly_once])', detail: 'Return unique values from a range or array.' },
  { name: 'UPPER', signature: 'UPPER(text)', detail: 'Convert text to uppercase.' },
  { name: 'VALUE', signature: 'VALUE(text)', detail: 'Convert text that appears in a recognized format to a number.' },
  { name: 'VLOOKUP', signature: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])', detail: 'Lookup a value in a table by column.' },
  { name: 'XLOOKUP', signature: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])', detail: 'Return a value from a range or array.' },
  { name: 'XMATCH', signature: 'XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])', detail: 'Return the relative position of an item in a range.' }
];

function activate(context) {
  const selector = { language: 'excel-formula', scheme: '*' };

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

  context.subscriptions.push(completionProvider, formattingProvider, rangeFormattingProvider);
}

function deactivate() {}

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

module.exports = {
  activate,
  deactivate,
  _formatFormula: formatFormula,
};
