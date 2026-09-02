import test from 'node:test';
import assert from 'node:assert/strict';
import XLSX from 'xlsx-js-style';
import { calcularPrecoVendaProjeto, itemVendaResumo, montarItensProposta, validarProposta, insumoEhMaterial, somarMoeda, arredondarMoeda, valorVendaUnitarioInsumo, formatarPrecisao, ratearMoeda } from '../src/utils/venda.js';
import { insumoValorUnitario, decomporInsumo } from '../src/utils/calculos.js';
import { etapasComOpcaoAtiva } from '../src/utils/alternativas.js';
import { criarAbaVendaModelo } from '../src/propostas/vendaXlsx.js';
import { consolidarMaoDeObra } from '../src/utils/maoObra.js';

const folha = (tipo, valorUnitario, coeficiente = 1, descricao = tipo) => ({ tipo, valorUnitario, coeficiente, descricao });
const etapasCom = (insumos, quantidade = 1) => [{ id: 'e1', nome: 'Etapa de teste', itens: [{ id: 'i1', servico: 'Serviço de teste', unidade: 'm²', quantidade, insumos }] }];
const bdi = { lucro: 0.2, dasAnexoIV: 0.1 };
const conferir = (etapas, calculo, desconto = 0) => {
  const grupos = montarItensProposta(etapas, calculo, [], null);
  validarProposta(grupos, arredondarMoeda(calculo.valorVenda - desconto), desconto, calculo.problemas);
  assert.equal(somarMoeda(grupos.map(g => g.total)), calculo.valorVenda);
  return grupos;
};

test('OUTROS, mobilização e serviços não são materiais', () => {
  for (const tipo of ['OUTROS', 'SERVIÇO', '', undefined, 'MO', 'EQUIP']) assert.equal(insumoEhMaterial(tipo), false);
  for (const tipo of ['MAT', 'material', 'MATERIAIS']) assert.equal(insumoEhMaterial(tipo), true);
  const etapas = etapasCom([folha('OUTROS', 500)]);
  const r = calcularPrecoVendaProjeto(etapas, bdi, [], null, { regimeMateriais: 'cliente' });
  assert.equal(r.valorVenda, 666.67);
  conferir(etapas, r);
});

test('subcomposição MAT conserva MO e equipamentos, exclui somente as folhas materiais do cliente', () => {
  const sub = { ...folha('MAT', 999), subCpuInsumos: [folha('MO', 20, 6), folha('MAT', 50, 2), folha('EQUIP', 10)] };
  const etapas = etapasCom([sub], 3);
  const r = calcularPrecoVendaProjeto(etapas, {}, [], null, { regimeMateriais: 'cliente' });
  assert.equal(r.valorVenda, 390);
  assert.equal(r.custoMateriaisCliente, 300);
  assert.equal(r.custoMaoObra, 360);
  assert.equal(r.custoDireto, 690);
  assert.equal(r.custoDiretoContratado, 390);
  conferir(etapas, r);
});

test('snapshots locais continuam válidos sem a CPU original na base', () => {
  const insumo = { ...folha('MAT', 9000, 0.25, 'Mistura'), subCpuInsumos: [folha('MO', 20, 6)] };
  assert.equal(insumoValorUnitario(insumo, []), 120);
  assert.equal(calcularPrecoVendaProjeto(etapasCom([insumo], 4), {}, []).valorVenda, 120);
  const profissionais = consolidarMaoDeObra(etapasCom([insumo], 4), []);
  assert.equal(profissionais[0].qtd, 6);
  assert.equal(profissionais[0].total, 120);
});

test('rateio de descontos e semanas fecha os centavos do cronograma', () => {
  assert.deepEqual(ratearMoeda(1, [1, 1, 1]), [0.34, 0.33, 0.33]);
  assert.deepEqual(ratearMoeda(0, [1, 2]), [0, 0]);
  assert.deepEqual(ratearMoeda(10, [0, 0]), [0, 0]);
  const liquidos = ratearMoeda(2159.98 - 159.98, [674.16, 1485.82]);
  assert.equal(somarMoeda(liquidos), 2000);
  for (const total of liquidos) assert.equal(somarMoeda(ratearMoeda(total, [33.33, 33.33, 33.34])), total);
  assert.equal(somarMoeda(ratearMoeda(100 * 0.4, [20, 20, 0])), 40);
});

test('unidades conflitantes são avisadas sem presumir conversão; material do cliente não gera falso preço ausente', () => {
  const etapas = etapasCom([{ ...folha('MO', 30, 1, 'Profissional'), unidade: 'H' }, { ...folha('MO', 30, 1, 'Profissional'), unidade: 'UN' }, folha('MAT', '')]);
  const r = calcularPrecoVendaProjeto(etapas, {}, [], null, { regimeMateriais: 'cliente' });
  assert.match(r.avisos.join(' '), /unidade/);
  assert.deepEqual(r.problemas, []);
  assert.equal(r.valorVenda, 60);
});

test('coeficientes de vários níveis, banco de preços do orçamento e preços zero', () => {
  const sub = { ...folha('MAT', 0, 0.5, 'Nível 1'), subCpuInsumos: [{ ...folha('MAT', 0, 0.25, 'Nível 2'), subCpuInsumos: [folha('MO', 5, 6, 'Profissional')] }] };
  const catalog = new Map([['profissional', { valorUnitario: 40 }]]);
  assert.equal(calcularPrecoVendaProjeto(etapasCom([sub], 2), {}, [], catalog).valorVenda, 60);
  const zero = new Map([['profissional', { valorUnitario: 0 }]]);
  const r = calcularPrecoVendaProjeto(etapasCom([sub], 2), {}, [], zero);
  assert.equal(r.valorVenda, 0);
  assert.deepEqual(r.problemas, []);
});

test('BDI e INSS incidem sobre a MO interna; faturamento direto só sobre material selecionado', () => {
  const ins = { ...folha('MAT', 0, 1, 'Mistura'), subCpuInsumos: [folha('MO', 100), folha('MAT', 100), folha('EQUIP', 100)] };
  const etapas = etapasCom([ins]);
  const taxas = { lucro: 0.2, retencaoInss: 0.1, faturamentoDireto: true, materiais: {}, materiaisFaturamentoDireto: ['mistura'] };
  const r = calcularPrecoVendaProjeto(etapas, taxas);
  assert.equal(r.valorVenda, 353.33);
  assert.equal(r.custoMateriaisFaturamentoDireto, 100);
  assert.equal(arredondarMoeda(r.retencaoInssValor), 13.33);
  assert.equal(arredondarMoeda(valorVendaUnitarioInsumo(ins, r, [], null)), 353.33);
  conferir(etapas, r);
  const collem = calcularPrecoVendaProjeto(etapas, { ...taxas, collemAtivo: true, collemX: 2, collemY: 1.5 });
  assert.equal(collem.valorVenda, 117.78);
});

test('alternativas inativas não entram nos totais e troca de opção é consistente', () => {
  const etapas = etapasCom([folha('MO', 50)]);
  etapas[0].gruposAlternativas = [{ id: 'g', opcaoAtivaId: 'a', opcoes: [{ id: 'a' }, { id: 'b' }] }];
  etapas[0].itens.push(...['a','b'].map((opcao, i) => ({ id: opcao, quantidade: 1, insumos: [folha('MO', (i + 1) * 100)], alternativaGrupoId: 'g', alternativaOpcaoId: opcao })));
  assert.equal(calcularPrecoVendaProjeto(etapas, {}).valorVenda, 150);
  const outra = etapasComOpcaoAtiva(etapas, 'e1', 'g', 'b');
  const r = calcularPrecoVendaProjeto(outra, {});
  assert.equal(r.valorVenda, 250);
  conferir(outra, r, 20);
});

test('quantidades fracionárias e grandes não arredondam o unitário antes da multiplicação', () => {
  for (const quantidade of [0, 0.000123456, 12.3456789, 1, 1000000]) {
    const etapas = etapasCom([folha('MO', 19.87654321, 0.1234567)], quantidade);
    const r = calcularPrecoVendaProjeto(etapas, bdi);
    const unitario = 19.87654321 * 0.1234567 * (1.2 / 0.9);
    assert.equal(r.valorVenda, arredondarMoeda(unitario * quantidade));
    assert.equal(itemVendaResumo(etapas[0].itens[0], r, [], null).quantidade, quantidade);
    conferir(etapas, r);
  }
  assert.equal(formatarPrecisao(0.000123456), '0,00');
});

test('apresentação usa duas casas sem truncar valores nem alterar os totais calculados', () => {
  for (const [valor, exibido] of [[674.1573033707865, '674,16'], [1.239, '1,24'], [1.2, '1,20'], [0, '0,00'], [1234567.89123, '1.234.567,89']]) {
    assert.equal(formatarPrecisao(valor), exibido);
  }
  const etapas = etapasCom([folha('MO', 674.1573033707865)], 12.3456789);
  const r = calcularPrecoVendaProjeto(etapas, {});
  const item = r.grupos[0].itens[0];
  assert.equal(item.quantidade, 12.3456789);
  assert.equal(item.unitario, 674.1573033707865);
  assert.equal(item.total, arredondarMoeda(12.3456789 * 674.1573033707865));
  conferir(etapas, r);
});

test('subtotais e total são somas dos itens em centavos, não dependem da ordenação', () => {
  const etapas = etapasCom([folha('MO', 0.005)]);
  etapas[0].itens.push({ ...etapas[0].itens[0], id: 'i2' });
  etapas.push({ ...etapas[0], id: 'e2', itens: [{ ...etapas[0].itens[0], id: 'i3' }] });
  const r = calcularPrecoVendaProjeto(etapas, {});
  assert.equal(r.valorVenda, 0.03);
  assert.equal(calcularPrecoVendaProjeto([...etapas].reverse(), {}).valorVenda, 0.03);
  conferir(etapas, r);
});

for (const modelo of ['alpha', 'collem']) test(`XLSX ${modelo}: fórmulas completas, caches e round-trip`, () => {
  const etapas = etapasCom([folha('OUTROS', 500)]);
  etapas.push({ id: 'e2', nome: 'Etapa vazia', itens: [] });
  const r = calcularPrecoVendaProjeto(etapas, bdi);
  const grupos = conferir(etapas, r);
  const ws = criarAbaVendaModelo(grupos, r.FatorBdi, modelo, 10);
  assert.equal(ws.G5.f, 'ROUND(E5*F5,2)');
  assert.equal(ws.H4.f, 'ROUND(SUM(G5:G5),2)');
  assert.equal(ws.H6.f, '0');
  assert.equal(ws.H4.v, 666.67);
  const total = Object.values(ws).find(c => c?.f?.startsWith('MAX(0,ROUND(SUM('));
  assert.equal(total.v, 656.67);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'VENDA');
  const copy = XLSX.read(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer' }).Sheets.VENDA;
  assert.equal(copy.G5.v, 666.67);
  assert.equal(copy.G5.f, ws.G5.f);
  assert.equal(copy.H4.f, ws.H4.f);
  assert.equal(copy.F5.v, grupos[0].itens[0].unitario);
  assert.equal(ws.E5.z, '#,##0.00');
  assert.equal(ws.F5.z, '#,##0.00');
  assert.equal(XLSX.utils.format_cell(copy.F5), '666.67');
});

test('campos numéricos em branco são zero, sem bloqueio nem alteração dos dados', () => {
  for (const etapas of [etapasCom([folha('MO', '')]), etapasCom([folha('MO', 5)], ''), etapasCom([folha('MO', 5, '')])]) {
    const original = JSON.stringify(etapas);
    const r = calcularPrecoVendaProjeto(etapas, {});
    assert.equal(r.valorVenda, 0);
    assert.deepEqual(r.problemas, []);
    conferir(etapas, r);
    assert.equal(JSON.stringify(etapas), original);
  }
  assert.match(calcularPrecoVendaProjeto(etapasCom([folha('MO', '')]), {}).avisos.join(' '), /considerado zero/);
});

test('bloqueia ciclo, valores inválidos, BDI impossível e subtotais divergentes', () => {
  for (const etapas of [etapasCom([folha('MO', -1)]), etapasCom([folha('MO', 5)], -1), etapasCom([folha('MO', 5, 'abc')])]) {
    const r = calcularPrecoVendaProjeto(etapas, {});
    assert.throws(() => conferir(etapas, r), /bloqueada/);
  }
  const loop = { id: 'loop', descricao: 'Ciclo', codigo: 'loop', insumos: [folha('MAT', 1, 1, 'Ciclo')] };
  const etapas = etapasCom([folha('MAT', 1, 1, 'Ciclo')]);
  const r = calcularPrecoVendaProjeto(etapas, {}, [loop]);
  assert.match(r.problemas.join(' '), /circular/);
  assert.throws(() => validarProposta(r.grupos, r.valorVenda, 0, r.problemas), /bloqueada/);
  const taxaInvalida = calcularPrecoVendaProjeto(etapasCom([folha('MO', 1)]), { dasAnexoIV: 1 });
  assert.match(taxaInvalida.problemas.join(' '), /BDI inválido/);
  const valido = calcularPrecoVendaProjeto(etapasCom([folha('MO', 10)]), {});
  assert.throws(() => validarProposta([{ ...valido.grupos[0], total: 0 }], 0), /Subtotal/);
  assert.throws(() => validarProposta(valido.grupos, 999), /Total da proposta/);
  assert.throws(() => validarProposta(valido.grupos, 0, 20), /Total da proposta/);
});
