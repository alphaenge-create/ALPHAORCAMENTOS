import { decomporInsumo, precoKey } from "./calculos.js";
import { itensAtivosDaEtapa } from "./alternativas.js";
import { norm, num, normalizarBusca } from "./format.js";

export const BDI_PADRAO = {
  custoInicial: 0, admCentral: 0.04, contabilidade: 0.01,
  contingenciamento: 0.02, custoFinanceiro: 0.03, dasAnexoIV: 0.13,
  art: 0, retencaoInss: 0, lucro: 0.42, collemAtivo: false, collemX: 1, collemY: 1,
};

export const materialPorContaCliente = (cliente) => cliente?.regimeMateriais === "cliente";
export const materialFaturamentoDireto = (cliente) => cliente?.regimeMateriais === "faturamentoDireto";
export const insumoEhMaterial = (tipo) => ["mat", "material", "materiais"].includes(norm(tipo).trim());
export const insumoEhMaoDeObra = (tipo) => norm(tipo).trim() === "mo" || norm(tipo).includes("mao de obra");

// OUTROS, serviços e categorias desconhecidas nunca são descartados como material.
export const insumoComFaturamentoDireto = (insumo, configuracao = {}, cliente = {}) => {
  if (!insumoEhMaterial(insumo?.tipo) || materialPorContaCliente(cliente)) return false;
  if (!configuracao.faturamentoDireto && !materialFaturamentoDireto(cliente)) return false;
  const selecionados = configuracao.materiaisFaturamentoDireto;
  return !Array.isArray(selecionados) || selecionados.includes(precoKey(insumo.descricao)) ||
    (insumo.chavesComposicoes || []).some((chave) => selecionados.includes(chave));
};

export const arredondarMoeda = (valor) => Math.sign(valor) * Math.round((Math.abs(valor) + Number.EPSILON * Math.max(1, Math.abs(valor))) * 100) / 100;
export const somarMoeda = (valores) => valores.reduce((soma, valor) => soma + Math.round(arredondarMoeda(valor) * 100), 0) / 100;
// Distribui centavos pelo maior resto: as parcelas sempre fecham com o total.
export const ratearMoeda = (total, pesos) => {
  const soma = pesos.reduce((s, p) => s + Math.max(0, num(p)), 0);
  const centavos = Math.round(arredondarMoeda(total) * 100);
  if (!soma || !Number.isSafeInteger(centavos) || centavos < 0) return pesos.map(() => 0);
  const parcelas = pesos.map((peso, indice) => {
    const exato = centavos * (Math.max(0, num(peso)) / soma);
    return { indice, centavos: Math.floor(exato), resto: exato - Math.floor(exato) };
  });
  const faltantes = centavos - parcelas.reduce((s, p) => s + p.centavos, 0);
  const prioridade = [...parcelas].sort((a, b) => b.resto - a.resto || a.indice - b.indice);
  for (let i = 0; i < faltantes; i++) prioridade[i % prioridade.length].centavos++;
  return parcelas.map((p) => p.centavos / 100);
};
// Duas casas apenas na apresentação; o cálculo mantém os números originais.
export const formatarPrecisao = (valor) => Number(valor).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

export const fatorVendaInsumo = (insumo, bdiCalc = {}, cliente = bdiCalc.cliente || {}) => {
  if (materialPorContaCliente(cliente) && insumoEhMaterial(insumo?.tipo)) return 0;
  if (insumoComFaturamentoDireto(insumo, bdiCalc, cliente)) return bdiCalc.FatorBdiMateriais ?? 1;
  if (insumoEhMaoDeObra(insumo?.tipo)) return bdiCalc.FatorBdiMaoObra ?? bdiCalc.FatorBdi ?? 1;
  return bdiCalc.FatorBdi ?? 1;
};

export const valorVendaUnitarioInsumo = (insumo, bdiCalc, cpus, catalogMap, cliente = bdiCalc.cliente || {}) => {
  const base = { ...insumo, coeficiente: 1 };
  return decomporInsumo(base, cpus, catalogMap).reduce((total, folha) =>
    total + folha.coeficiente * folha.valorUnitario * fatorVendaInsumo(folha, bdiCalc, cliente), 0);
};

export const itemVendaResumo = (item, bdiCalc, cpus, catalogMap, cliente = bdiCalc.cliente || {}) => {
  const quantidade = num(item.quantidade);
  const unitario = (item.insumos || []).reduce((total, insumo) =>
    total + num(insumo.coeficiente) * valorVendaUnitarioInsumo(insumo, bdiCalc, cpus, catalogMap, cliente), 0);
  return { quantidade, unitario, total: arredondarMoeda(quantidade * unitario) };
};

export const montarItensProposta = (etapas, bdiCalc, cpus, catalogMap, cliente = bdiCalc.cliente || {}) =>
  (etapas || []).map((etapa, indice) => {
    const itens = itensAtivosDaEtapa(etapa).map((item, indiceItem) => ({
      id: item.id, numero: `${indice + 1}.${indiceItem + 1}`,
      descricao: item.servico || item.descricao || "", unidade: item.unidade || "",
      ...itemVendaResumo(item, bdiCalc, cpus, catalogMap, cliente),
    }));
    return { id: etapa.id, numero: String(indice + 1), nome: etapa.nome || `Etapa ${indice + 1}`, itens, total: somarMoeda(itens.map((item) => item.total)) };
  });

export const calcularPrecoVendaProjeto = (etapas, bdi = BDI_PADRAO, cpus = [], catalogMap = null, cliente = {}) => {
  bdi = bdi || BDI_PADRAO;
  const problemas = [];
  const avisos = [];
  const fator = (taxas = {}) => {
    const indices = [taxas.admCentral ?? taxas.adminCentral, taxas.contabilidade, taxas.contingenciamento, taxas.custoFinanceiro, taxas.lucro].map(num);
    const taxasTributos = [num(taxas.dasAnexoIV), num(taxas.art)];
    const tributos = taxasTributos[0] + taxasTributos[1];
    if ([...indices, ...taxasTributos].some((n) => !Number.isFinite(n) || n < 0) || tributos >= 1) {
      problemas.push("BDI inválido: verifique índices e tributos (a soma de DAS e ART deve ser inferior a 100%).");
      return 1;
    }
    return indices.reduce((produto, indice) => produto * (1 + indice), 1) / (1 - tributos);
  };
  const FatorBdiBase = fator(bdi);
  const faturamentoDireto = !materialPorContaCliente(cliente) && (!!bdi.faturamentoDireto || materialFaturamentoDireto(cliente));
  const FatorBdiMateriaisBase = faturamentoDireto && bdi.materiais ? fator(bdi.materiais) : FatorBdiBase;
  const retencaoInss = num(bdi.retencaoInss);
  if (!Number.isFinite(retencaoInss) || retencaoInss < 0 || retencaoInss >= 1) problemas.push("Retenção de INSS inválida.");
  const FatorBdiMaoObraBase = FatorBdiBase / (1 - Math.min(0.99, Math.max(0, retencaoInss || 0)));
  const collemAtivo = !!bdi.collemAtivo;
  const collemX = num(bdi.collemX) || 1;
  const collemY = num(bdi.collemY) || 1;
  if (collemAtivo && (!Number.isFinite(collemX * collemY) || collemX <= 0 || collemY <= 0 || num(bdi.collemX) === 0 || num(bdi.collemY) === 0)) problemas.push("Divisores COLLEM inválidos.");
  const divisorCollem = collemAtivo && collemX * collemY > 0 ? collemX * collemY : 1;
  const config = {
    cliente, faturamentoDireto, materiaisFaturamentoDireto: bdi.materiaisFaturamentoDireto,
    FatorBdi: FatorBdiBase / divisorCollem,
    FatorBdiMateriais: FatorBdiMateriaisBase / divisorCollem,
    FatorBdiMaoObra: FatorBdiMaoObraBase / divisorCollem,
  };
  let custoDireto = 0;
  let custoMaoObra = 0;
  let custoMateriaisFaturamentoDireto = 0;
  let custoMateriaisCliente = 0;
  let valorVendaSemArredondamento = 0;
  const unidadesPorDescricao = new Map();
  for (const etapa of etapas || []) for (const item of itensAtivosDaEtapa(etapa)) {
    const quantidade = num(item.quantidade);
    if (!Number.isFinite(quantidade) || quantidade < 0) problemas.push(`Quantidade inválida: ${item.servico || item.descricao}`);
    if (quantidade === 0) continue;
    if (!(item.insumos || []).length) problemas.push(`CPU sem insumos: ${item.servico || item.descricao}`);
    for (const insumo of item.insumos || []) {
      const folhas = decomporInsumo(insumo, cpus, catalogMap, 1, new Set(), problemas);
      for (const folha of folhas) {
        const excluido = materialPorContaCliente(cliente) && insumoEhMaterial(folha.tipo);
        const descricao = precoKey(folha.descricao);
        const unidade = normalizarBusca(folha.unidade);
        if (!excluido) {
          const registro = unidadesPorDescricao.get(descricao) || { descricao: folha.descricao, unidades: new Set() };
          const unidadeCatalogo = normalizarBusca(catalogMap?.get(descricao)?.unidade);
          if (unidade) registro.unidades.add(unidade);
          if (unidadeCatalogo) registro.unidades.add(unidadeCatalogo);
          unidadesPorDescricao.set(descricao, registro);
        }
        if (!excluido && folha.precoAusente) avisos.push(`Preço em branco considerado zero: ${folha.descricao}. Confira se esse consumo deve mesmo ficar sem custo.`);
        if (!excluido && (!Number.isFinite(folha.valorUnitario) || folha.valorUnitario < 0)) problemas.push(`Preço inválido: ${folha.descricao}`);
        const custo = quantidade * folha.coeficiente * (Number.isFinite(folha.valorUnitario) ? folha.valorUnitario : 0);
        custoDireto += custo;
        if (excluido) { custoMateriaisCliente += custo; continue; }
        if (insumoEhMaoDeObra(folha.tipo)) custoMaoObra += custo;
        if (insumoComFaturamentoDireto(folha, config, cliente)) custoMateriaisFaturamentoDireto += custo;
        valorVendaSemArredondamento += custo * fatorVendaInsumo(folha, config, cliente);
      }
    }
  }
  for (const registro of unidadesPorDescricao.values()) if (registro.unidades.size > 1) {
    avisos.push(`Confira a unidade do insumo ${registro.descricao}: ${[...registro.unidades].sort().join("/")}. O banco usa a descrição para vincular seu preço.`);
  }
  const grupos = montarItensProposta(etapas, config, cpus, catalogMap, cliente);
  const valorVenda = somarMoeda(grupos.map((grupo) => grupo.total));
  const custoDiretoContratado = custoDireto - custoMateriaisCliente;
  const totalDiValor = valorVenda - custoDiretoContratado;
  return {
    ...config, grupos, problemas: [...new Set(problemas)], avisos: [...new Set(avisos)],
    bdiRate: FatorBdiBase - 1, bdiRateMateriais: FatorBdiMateriaisBase - 1, bdiRateMaoObra: FatorBdiMaoObraBase - 1,
    FatorBdiBase, FatorBdiMateriaisBase, FatorBdiMaoObraBase,
    retencaoInss, custoMaoObra, custoMateriaisFaturamentoDireto, custoMateriaisCliente,
    custoBaseBdiGeral: custoDiretoContratado - custoMateriaisFaturamentoDireto,
    retencaoInssValorBase: custoMaoObra * (FatorBdiMaoObraBase - FatorBdiBase),
    retencaoInssValor: custoMaoObra * (FatorBdiMaoObraBase - FatorBdiBase) / divisorCollem,
    collemAtivo, collemX, collemY, divisorCollem,
    custoDireto, custoDiretoContratado, totalDiValor,
    totalDiRate: custoDiretoContratado > 0 ? totalDiValor / custoDiretoContratado : 0,
    valorVendaSemArredondamento, valorVendaBase: valorVendaSemArredondamento * divisorCollem, valorVenda,
  };
};

// Interrompe todas as exportações antes de produzir um documento inconsistente.
export const validarProposta = (grupos, totalGeral, desconto = 0, problemas = []) => {
  const erros = [...problemas];
  for (const grupo of grupos) {
    for (const item of grupo.itens) {
      if ([item.quantidade, item.unitario, item.total].some((valor) => !Number.isFinite(valor) || valor < 0) || !Number.isSafeInteger(Math.round(item.total * 100))) erros.push(`Valores inválidos ou acima da precisão segura no item ${item.numero}.`);
      if (Math.abs(arredondarMoeda(item.quantidade * item.unitario) - item.total) > 0.001) erros.push(`Quantidade × unitário não confere no item ${item.numero}.`);
    }
    if (Math.abs(somarMoeda(grupo.itens.map((item) => item.total)) - grupo.total) > 0.001) erros.push(`Subtotal não confere na etapa ${grupo.numero}.`);
  }
  const bruto = somarMoeda(grupos.map((grupo) => grupo.total));
  if (![bruto, desconto, totalGeral].every(Number.isFinite) || desconto < 0 || desconto > bruto || Math.abs(arredondarMoeda(bruto - desconto) - totalGeral) > 0.001) erros.push("Total da proposta não confere com os itens e o desconto.");
  if (erros.length) throw new Error(`Proposta bloqueada para evitar valores incorretos:\n${[...new Set(erros)].slice(0, 12).join("\n")}`);
  return { totalBruto: bruto, desconto, totalGeral };
};
