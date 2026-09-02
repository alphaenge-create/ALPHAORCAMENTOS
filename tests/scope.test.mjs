import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const { parse } = require('@babel/parser');
const traverse = require('@babel/traverse').default;
const globais = new Set(['window', 'document', 'navigator', 'localStorage', 'sessionStorage', 'indexedDB', 'console', 'alert', 'confirm', 'setTimeout', 'clearTimeout', 'setInterval', 'clearInterval', 'URL', 'Blob', 'FileReader', 'ResizeObserver', 'requestAnimationFrame', 'cancelAnimationFrame', 'fetch', 'crypto', 'TextEncoder', 'TextDecoder', 'structuredClone', 'Response', 'CompressionStream', 'DecompressionStream', 'AbortController', 'performance']);
for (const arquivo of ['../src/App.jsx', '../src/propostas/vendaXlsx.js', '../src/propostas/alphaProposal.js', '../src/utils/venda.js']) test(`variáveis usadas no escopo correto: ${arquivo}`, () => {
  const ast = parse(fs.readFileSync(new URL(arquivo, import.meta.url), 'utf8'), { sourceType: 'module', plugins: ['jsx'] });
  const erros = [];
  traverse(ast, { ReferencedIdentifier(p) { if (!p.scope.hasBinding(p.node.name) && !globais.has(p.node.name)) erros.push(`${p.node.name}:${p.node.loc.start.line}`); } });
  assert.deepEqual([...new Set(erros)], []);
});
