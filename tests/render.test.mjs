import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { buildSync } from 'esbuild';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';

const require = createRequire(import.meta.url);
const arquivoApp = new URL('../src/App.jsx', import.meta.url);
const projeto = {
  id: 'teste', nome: 'Orçamento de teste', precos: [], bdi: {},
  clienteCadastro: { nome: 'Cliente de teste', local: 'Local de teste', regimeMateriais: 'cliente', modeloProposta: 'alpha' },
  etapas: [{ id: 'e', nome: 'Mobilização', itens: [{ id: 'i', quantidade: 2, servico: 'Serviço', unidade: 'un', insumos: [{ id: 's', tipo: 'OUTROS', descricao: 'Serviço avulso', coeficiente: 1, valorUnitario: 500 }] }] }],
};

const carregar = (entrada, conteudo) => {
  const options = { bundle: true, platform: 'node', format: 'cjs', packages: 'external', write: false, loader: { '.png': 'dataurl' }, define: { 'import.meta.env': '{}' }, logLevel: 'silent' };
  const output = buildSync(conteudo ? { ...options, stdin: { contents: conteudo, loader: 'jsx', resolveDir: path.dirname(entrada), sourcefile: entrada } } : { ...options, entryPoints: [entrada] });
  const modulo = { exports: {} };
  new Function('require', 'module', 'exports', output.outputFiles[0].text)(require, modulo, modulo.exports);
  return modulo.exports;
};

for (const aba of ['precovenda', 'planilha', 'custo', 'materiais', 'maoobra', 'cronograma', 'bdi']) test(`renderização real da aba ${aba}, com fixture isolada e sem efeitos de rede`, () => {
  const source = fs.readFileSync(arquivoApp, 'utf8')
    .replace('useState("projetos")', `useState(${JSON.stringify(aba)})`)
    .replace('const [projetos, setProjetos] = useState([])', `const [projetos, setProjetos] = useState(${JSON.stringify([projeto])})`)
    .replace('const [projetoAtivoId, setProjetoAtivoId] = useState("")', 'const [projetoAtivoId, setProjetoAtivoId] = useState("teste")');
  const { default: App } = carregar(fs.realpathSync(arquivoApp), source);
  const html = renderToStaticMarkup(React.createElement(App));
  assert.ok(html.includes('Orçamentador por CPU'));
  if (aba === 'precovenda') {
    assert.ok(html.includes('VALOR FINAL DE VENDA'));
    assert.ok(html.includes('1.000,00'));
  }
});

test('DOCX real usa os mesmos valores e recusa totais incoerentes antes de exportar', async () => {
  const { criarPropostaAlphaDocxBlob } = carregar(fs.realpathSync(new URL('../src/propostas/alphaProposal.js', import.meta.url)));
  const dados = { grupos: [{ numero: '1', nome: 'Mobilização', total: 1000, itens: [{ numero: '1.1', descricao: 'Serviço', unidade: 'un', quantidade: 2, unitario: 500, total: 1000 }] }], totalGeral: 1000 };
  const logoData = fs.readFileSync(new URL('../src/assets/alpha-engenharia-logo.png', import.meta.url));
  const blob = await criarPropostaAlphaDocxBlob(dados, { logoData });
  assert.ok(blob.size > 1000);
  // Inspeciona o XML real empacotado, sem Word, downloads ou alteração de arquivo.
  const zip = require('fflate').unzipSync(new Uint8Array(await blob.arrayBuffer()));
  const xml = new TextDecoder().decode(zip['word/document.xml']);
  assert.ok(xml.includes('1.000,00'));
  assert.ok(xml.includes('500,00'));
  await assert.rejects(criarPropostaAlphaDocxBlob({ ...dados, totalGeral: 999 }, { logoData }), /Total da proposta/);
});
