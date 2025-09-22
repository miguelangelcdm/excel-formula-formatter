const fs = require('fs');
const vm = require('vm');

const code = fs.readFileSync(require.resolve('./extension.js'), 'utf8');

class Disposable { dispose() {} }

const vscodeMock = {
  languages: {
    registerCompletionItemProvider() { return new Disposable(); },
    registerDocumentFormattingEditProvider() { return new Disposable(); },
    registerDocumentRangeFormattingEditProvider() { return new Disposable(); }
  },
  CompletionItem: class {},
  CompletionItemKind: { Function: 0 },
  SnippetString: class { constructor(value) { this.value = value; } },
  Range: class {},
  TextEdit: { replace: (range, text) => ({ range, newText: text }) }
};

const sandbox = {
  module: { exports: {} },
  exports: {},
  require: (name) => {
    if (name === 'vscode') {
      return vscodeMock;
    }
    return require(name);
  },
  console
};

vm.createContext(sandbox);
vm.runInContext(code, sandbox, { filename: 'extension.js' });

const { _formatFormula: formatFormula } = sandbox.module.exports;

const samples = [
  {
    name: 'Basic IF',
    input: '=IF(AND(A1>5,B1<10),SUM(C1:C10),0)'
  },
  {
    name: 'LET with FILTER',
    input: '=LET(x, SUM(A1:A5), y, FILTER(B1:B5, B1:B5>0), x + SUM(y))'
  },
  {
    name: 'Structured reference',
    input: '=SUMIFS(Table1[Amount], Table1[Category], "Travel", Table1[Date], ">=" & DATE(2024,1,1))'
  },
  {
    name: 'Array literal',
    input: '=MMULT({1,2;3,4},TRANSPOSE({1;2}))'
  }
];

samples.forEach(({ name, input }) => {
  const formatted = formatFormula(input);
  const idempotent = formatFormula(formatted) === formatted;
  console.log(`--- ${name} ---`);
  console.log(formatted);
  console.log(idempotent ? 'Idempotent: yes' : 'Idempotent: no');
});
