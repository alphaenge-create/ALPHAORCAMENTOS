// Leitura somente. Arquivos de clientes nunca são copiados para o repositório.
import fs from 'node:fs';
import zlib from 'node:zlib';
import XLSX from 'xlsx-js-style';
import { buildCatalog } from '../src/utils/calculos.js';
import { calcularPrecoVendaProjeto, validarProposta, arredondarMoeda } from '../src/utils/venda.js';
const [projetoPath, basePath, xlsxPath] = process.argv.slice(2);
if (!projetoPath || !basePath) throw new Error('Uso: node scripts/auditar-orcamento.mjs projeto.json.gz base.json.gz [original.xlsx]');
const ler = (path) => JSON.parse(path.endsWith('.gz') ? zlib.gunzipSync(fs.readFileSync(path)) : fs.readFileSync(path, 'utf8'));
const origem = ler(projetoPath);
const projeto = origem.projeto || origem;
const { cpus = [] } = ler(basePath);
const catalog = new Map(buildCatalog(cpus, [projeto], projeto.id, projeto.precos || []).map(p => [p.key, p]));
const r = calcularPrecoVendaProjeto(projeto.etapas, projeto.bdi, cpus, catalog, projeto.clienteCadastro);
const desconto = Math.min(r.valorVenda, Math.max(0, arredondarMoeda(Number(projeto.descontoNegociacao) || 0)));
let erroValidacao = null;
try { validarProposta(r.grupos, arredondarMoeda(r.valorVenda - desconto), desconto, r.problemas); }
catch (erro) { erroValidacao = erro.message; process.exitCode = 1; }
console.log(JSON.stringify({ salvoEm: origem.savedAt, erroValidacao, totalBruto: r.valorVenda, desconto, totalCalculado: arredondarMoeda(r.valorVenda - desconto), problemas: r.problemas, avisos: r.avisos, etapas: r.grupos.map(g => ({ nome: g.nome, subtotal: g.total, itens: g.itens.map(i => ({ descricao: i.descricao, quantidade: i.quantidade, unitario: i.unitario, total: i.total })) })) }, null, 2));
if (xlsxPath) {
  const wb = XLSX.readFile(xlsxPath);
  const ws = wb.Sheets.VENDA;
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const totalAntigo = rows.find(row => row.includes('TOTAL GERAL'))?.[7];
  console.log(JSON.stringify({ totalXlsxOriginal: totalAntigo, diferenca: Math.round((r.valorVenda - desconto - totalAntigo) * 100) / 100 }));
}
