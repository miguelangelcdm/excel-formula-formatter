const vm = require('vm');
const assert = require('assert');

const code = require('fs').readFileSync(require.resolve('./extension.js'), 'utf8');

class Disposable { dispose() {} }

// Superset of the run-format-tests.js mock so activate() can also run here.
const state = {
  completion: null,
  openHandlers: [],
  changeHandlers: [],
  closeHandlers: [],
  collections: []
};

const vscodeMock = {
  languages: {
    registerCompletionItemProvider(selector, provider, ...triggers) {
      state.completion = { provider, triggers };
      return new Disposable();
    },
    registerDocumentFormattingEditProvider() { return new Disposable(); },
    registerDocumentRangeFormattingEditProvider() { return new Disposable(); },
    registerHoverProvider() { return new Disposable(); },
    createDiagnosticCollection(name) {
      const collection = {
        name,
        disposed: false,
        entries: new Map(),
        set(document, diagnostics) { this.entries.set(document, diagnostics); },
        delete(document) { this.entries.delete(document); },
        dispose() { this.disposed = true; }
      };
      state.collections.push(collection);
      return collection;
    }
  },
  workspace: {
    textDocuments: [],
    onDidOpenTextDocument(handler) { state.openHandlers.push(handler); return new Disposable(); },
    onDidChangeTextDocument(handler) { state.changeHandlers.push(handler); return new Disposable(); },
    onDidCloseTextDocument(handler) { state.closeHandlers.push(handler); return new Disposable(); }
  },
  window: { activeTextEditor: undefined, showInformationMessage() {} },
  commands: { registerCommand() { return new Disposable(); } },
  CompletionItem: class { constructor(label, kind) { this.label = label; this.kind = kind; } },
  CompletionItemKind: { Function: 0 },
  SnippetString: class { constructor(value) { this.value = value; } },
  Position: class { constructor(line, character) { this.line = line; this.character = character; } },
  Range: class { constructor(start, end) { this.start = start; this.end = end; } },
  TextEdit: { replace: (range, text) => ({ range, newText: text }) },
  MarkdownString: class { appendCodeblock() {} appendMarkdown() {} },
  Hover: class { constructor(contents, range) { this.contents = contents; this.range = range; } },
  Diagnostic: class {
    constructor(range, message, severity) {
      this.range = range;
      this.message = message;
      this.severity = severity;
    }
  },
  DiagnosticSeverity: { Error: 0, Warning: 1, Information: 2, Hint: 3 }
};

// Minimal text document for the diagnostics wiring tests.
function makeDocument(text, languageId = 'excel-formula') {
  const lineStarts = [0];
  for (let i = 0; i < text.length; i += 1) {
    if (text[i] === '\n') {
      lineStarts.push(i + 1);
    }
  }
  return {
    languageId,
    getText: () => text,
    offsetAt(position) {
      const start = lineStarts[Math.min(position.line, lineStarts.length - 1)];
      return Math.min(start + position.character, text.length);
    },
    positionAt(offset) {
      let line = lineStarts.length - 1;
      while (line > 0 && lineStarts[line] > offset) {
        line -= 1;
      }
      return new vscodeMock.Position(line, offset - lineStarts[line]);
    }
  };
}

const sandbox = {
  module: { exports: {} },
  exports: {},
  require: (name) => (name === 'vscode' ? vscodeMock : require(name)),
  console,
  setTimeout: (fn, ms) => setTimeout(fn, ms),
  clearTimeout: (handle) => clearTimeout(handle)
};

vm.createContext(sandbox);
vm.runInContext(code, sandbox, { filename: 'extension.js' });

const {
  activate,
  deactivate,
  _formatFormula: formatFormula,
  _minifyFormula: minifyFormula
} = sandbox.module.exports;

// Clone checker results into host objects so assert.deepStrictEqual does not
// trip over vm-context vs host Array/Object prototypes.
const checkFormula = input => JSON.parse(JSON.stringify(sandbox.module.exports._checkFormula(input)));

// Activate once against the mock so wiring can be exercised.
activate({ subscriptions: [] });

// The broken sample from GitHub issue #2, verbatim.
const issueExample = `=BITOR(
  BITOR(
    BITOR(
      BITOR(
        BITOR(
          AM3,
          BITLSHIFT(
            AM4,
            SUM(
              $AK$3
            )
          )
        ),
        BITLSHIFT(
          AM5,
          SUM(
            $A$K3:$AK$4
          )
        )
      ),
      BITLSHIFT(
        AM6,
        SUM(
          $AK$3:$AK$5
        )
      ),
    ),
    BITLSHIFT(
      AM7,
      SUM(
        $AK$3:$AK$6,
      )
    )
  ),
  BITLSHIFT(
    AM8
    SUM(
      $AK$3:$AK$7
    )
  )
)`;

const samples = [
  '=IF(AND(A1>5,B1<10),SUM(C1:C10),0)',
  '=LET(x, SUM(A1:A5), y, FILTER(B1:B5, B1:B5>0), x + SUM(y))',
  '=SUMIFS(Table1[Amount], Table1[Category], "Travel", Table1[Date], ">=" & DATE(2024,1,1))',
  '=MMULT({1,2;3,4},TRANSPOSE({1;2}))'
];

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

// --- checkFormula: clean inputs ---------------------------------------------

test('empty input yields no diagnostics', () => {
  assert.deepStrictEqual(checkFormula(''), []);
});

test('balanced formulas produce no diagnostics', () => {
  samples.forEach(input => {
    assert.deepStrictEqual(checkFormula(input), [], `unexpected diagnostics for ${input}`);
  });
  assert.deepStrictEqual(checkFormula('=SUM(\n  A1,\n  B1\n)'), []);
});

// --- checkFormula: the four checks ------------------------------------------

test('unmatched closer is reported', () => {
  const diagnostics = checkFormula('=SUM(A1))');
  assert.strictEqual(diagnostics.length, 1);
  assert.strictEqual(diagnostics[0].message, "Unmatched ')'");
  assert.deepStrictEqual(
    { line: diagnostics[0].line, character: diagnostics[0].character },
    { line: 0, character: 8 }
  );
});

test('unclosed openers are reported at their positions', () => {
  const diagnostics = checkFormula('=IF(A1,\n  SUM(B1\n');
  const unclosed = diagnostics.filter(d => d.message === "Unclosed '('");
  assert.strictEqual(unclosed.length, 2);
  assert.deepStrictEqual(
    unclosed.map(d => [d.line, d.character]),
    [[0, 3], [1, 5]]
  );
});

test('mismatched bracket pair is reported', () => {
  const messages = checkFormula('=SUM({1,2)').map(d => d.message);
  assert.ok(messages.includes("Unmatched ')'"), messages.join('|'));
  assert.ok(messages.includes("Unclosed '{'"), messages.join('|'));
});

test('unterminated string literal is reported', () => {
  const diagnostics = checkFormula('=CONCAT(A1,"abc');
  const unterminated = diagnostics.find(d => d.message === 'Unterminated string literal');
  assert.ok(unterminated, diagnostics.map(d => d.message).join('|'));
  assert.strictEqual(unterminated.character, 11);
});

test('trailing comma before ) is reported at the comma', () => {
  const diagnostics = checkFormula('=SUM(A1,)');
  assert.strictEqual(diagnostics.length, 1);
  assert.strictEqual(diagnostics[0].message, "Trailing comma before ')'");
  assert.strictEqual(diagnostics[0].character, 7);
});

test('trailing comma before } is reported', () => {
  const messages = checkFormula('=SUM({1,2;3,4,})').map(d => d.message);
  assert.ok(messages.includes("Trailing comma before '}'"), messages.join('|'));
});

test('trailing comma reports correct line in multi-line input', () => {
  const diagnostics = checkFormula('=SUM(\n  A1,\n)');
  assert.strictEqual(diagnostics.length, 1);
  assert.strictEqual(diagnostics[0].message, "Trailing comma before ')'");
  assert.strictEqual(diagnostics[0].line, 1);
  assert.strictEqual(diagnostics[0].character, 4);
});

test('missing comma before a call is reported', () => {
  const diagnostics = checkFormula('=BITLSHIFT(AM8 SUM(A1))');
  assert.strictEqual(diagnostics.length, 1);
  assert.strictEqual(diagnostics[0].message, "Missing ',' before 'SUM'");
  assert.strictEqual(diagnostics[0].character, 15);
});

test('word right after a closed call is reported', () => {
  const messages = checkFormula('=BITLSHIFT(SUM(A1) AM8)').map(d => d.message);
  assert.ok(messages.includes("Missing ',' before 'AM8'"), messages.join('|'));
});

// --- checkFormula: deliberate non-reports (conservative rules) --------------

test('bare adjacent words inside a call are NOT flagged', () => {
  // Known limitation: the word-word rule only fires when the second word
  // opens a call, so "A1 B2" stays silent to avoid false positives.
  assert.deepStrictEqual(checkFormula('=SUM(A1 B2)'), []);
});

test('juxtaposition outside a call/array is NOT flagged', () => {
  assert.deepStrictEqual(checkFormula('=SUM(A1)B1'), []);
});

// --- issue #2 sample ---------------------------------------------------------

test('issue #2 sample yields trailing-comma and missing-comma diagnostics', () => {
  const messages = checkFormula(issueExample).map(d => d.message);
  assert.ok(messages.some(m => m.startsWith('Trailing comma before')), messages.join('|'));
  assert.ok(messages.some(m => m.startsWith("Missing ',' before")), messages.join('|'));
  // Parentheses in the sample are balanced: nothing else should fire.
  assert.ok(!messages.some(m => m.startsWith('Unmatched')), messages.join('|'));
  assert.ok(!messages.some(m => m.startsWith('Unclosed')), messages.join('|'));
  assert.strictEqual(messages.filter(m => m.startsWith('Trailing comma')).length, 2);
  assert.strictEqual(messages.filter(m => m.startsWith("Missing ','")).length, 1);
});

// --- formatter/minify guards --------------------------------------------------

test('formatFormula stays idempotent on the existing samples', () => {
  samples.forEach(input => {
    const formatted = formatFormula(input);
    assert.strictEqual(formatFormula(formatted), formatted, `not idempotent: ${input}`);
  });
});

test('minify after format round-trips to the minified original', () => {
  samples.forEach(input => {
    assert.strictEqual(
      minifyFormula(formatFormula(input)),
      minifyFormula(input),
      `round-trip failed: ${input}`
    );
  });
});

test('minifyFormula still collapses a multi-line formula', () => {
  const input = '=IF(\n  A1 > 5,\n  "big",\n  SUM(B1:B9)\n)';
  assert.strictEqual(minifyFormula(input), '=IF(A1>5,"big",SUM(B1:B9))');
});

// --- postfix percent literals -------------------------------------------------

test('percent literals keep their number when formatted', () => {
  const input = '=IF(A1*2.50%<=2.40,2.40,A1*3%)';
  const formatted = formatFormula(input);
  assert.ok(formatted.includes('2.50%'), `literal split up in:\n${formatted}`);
  assert.ok(formatted.includes('3%'), `literal split up in:\n${formatted}`);
  // A space before '%' means the tokenizer split the literal (issue: "2.50 %").
  assert.ok(!formatted.includes(' %'), `stray space detached a percent in:\n${formatted}`);
  assert.strictEqual(formatFormula(formatted), formatted, 'formatter not idempotent');
});

test('percent formulas round-trip through format and minify', () => {
  const input = '=ROUND(IF(PaymentCountry="PER",Amount*2.50%,Amount*3.00%),2)';
  const roundTripped = minifyFormula(formatFormula(input));
  assert.strictEqual(roundTripped, minifyFormula(input));
  assert.ok(roundTripped.includes('Amount*2.50%'), roundTripped);
  assert.ok(roundTripped.includes('Amount*3.00%'), roundTripped);
});

test('percent literals do not confuse the checker', () => {
  assert.deepStrictEqual(checkFormula('=SUM(A1*2.5%,B1*3%)'), []);
});

// --- issue #3: completion items must not commit on '(' -----------------------

test('completion items have no commitCharacters (issue #3)', () => {
  assert.ok(state.completion, 'completion provider was not registered');
  assert.ok(state.completion.triggers.includes('('), "'(' trigger should be kept");
  const items = state.completion.provider.provideCompletionItems(makeDocument('='), null);
  assert.ok(items.length > 0, 'no completion items returned');
  items.forEach(item => {
    assert.strictEqual(item.commitCharacters, undefined, `commitCharacters set on ${item.label}`);
    assert.strictEqual(item.insertText.value, `${item.label}($0)`);
  });
});

// --- diagnostics wiring --------------------------------------------------------

test('opened documents are validated immediately', () => {
  const collection = state.collections[0];
  const document = makeDocument('=SUM(A1,)');
  state.openHandlers.forEach(handler => handler(document));
  const diagnostics = collection.entries.get(document);
  assert.ok(diagnostics, 'no diagnostics published for open document');
  assert.strictEqual(diagnostics.length, 1);
  assert.ok(diagnostics[0] instanceof vscodeMock.Diagnostic);
  assert.strictEqual(diagnostics[0].message, "Trailing comma before ')'");
  assert.strictEqual(diagnostics[0].range.start.line, 0);
  assert.strictEqual(diagnostics[0].range.start.character, 7);
});

test('non excel-formula documents are ignored', () => {
  const collection = state.collections[0];
  const document = makeDocument('=SUM(A1,', 'plaintext');
  state.openHandlers.forEach(handler => handler(document));
  assert.strictEqual(collection.entries.has(document), false);
});

test('change events are debounced then validated', async () => {
  const collection = state.collections[0];
  const document = makeDocument('=SUM(A1))');
  state.changeHandlers.forEach(handler => handler({ document }));
  state.changeHandlers.forEach(handler => handler({ document }));
  assert.strictEqual(collection.entries.has(document), false, 'validation ran before debounce expired');
  await new Promise(resolve => setTimeout(resolve, 450));
  const diagnostics = collection.entries.get(document);
  assert.ok(diagnostics, 'no diagnostics after debounce');
  assert.strictEqual(diagnostics[0].message, "Unmatched ')'");
});

test('closed documents drop their diagnostics', () => {
  const collection = state.collections[0];
  const document = makeDocument('=SUM(A1)');
  state.openHandlers.forEach(handler => handler(document));
  state.closeHandlers.forEach(handler => handler(document));
  assert.strictEqual(collection.entries.has(document), false);
});

test('deactivate disposes the diagnostic collection', () => {
  deactivate();
  assert.strictEqual(state.collections[0].disposed, true);
});

// --- runner --------------------------------------------------------------------

(async () => {
  let failures = 0;
  for (const { name, fn } of tests) {
    try {
      await fn();
      console.log(`ok   ${name}`);
    } catch (err) {
      failures += 1;
      console.error(`FAIL ${name}`);
      console.error(`     ${err && err.message}`);
    }
  }
  console.log(`\n${tests.length - failures}/${tests.length} passed`);
  if (failures) {
    process.exitCode = 1;
  }
})();
