import XLSX from "xlsx-js-style";
import { num } from "../utils/format.js";
import { arredondarMoeda, somarMoeda, validarProposta } from "../utils/venda.js";
const aplicarDescontoNegociacao = (bruto, desconto) => arredondarMoeda(Math.max(0, bruto - desconto));

const XLSX_MOEDA = '_-"R$"\\ * #,##0.00_-;\\-"R$"\\ * #,##0.00_-;_-"R$"\\ * "-"??_-;_-@';
const XLSX_NUMERO = "#,##0.00";

const estiloVendaBase = {
  font: { name: "Aptos Narrow", sz: 12 },
  alignment: { vertical: "center" },
};

const estiloVendaTitulo = {
  font: { name: "Aptos Narrow", sz: 14, bold: true },
  fill: { fgColor: { rgb: "7B9A56" } },
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
};

const estiloVendaCabecalho = {
  font: { name: "Aptos Narrow", sz: 12, bold: true },
  fill: { fgColor: { rgb: "D8D8D8" } },
  alignment: { vertical: "center" },
};

const estiloVendaGrupo = {
  font: { name: "Aptos Narrow", sz: 12, bold: true },
  fill: { fgColor: { rgb: "E2EFD9" } },
  alignment: { vertical: "center" },
};

const estiloVendaTotal = {
  ...estiloVendaGrupo,
  font: { name: "Aptos Narrow", sz: 11, bold: true },
  alignment: { horizontal: "center", vertical: "center" },
};

const estiloVendaCollemBase = {
  font: { name: "Calibri", sz: 11 },
  alignment: { vertical: "center" },
};

const estiloVendaCollemTitulo = {
  font: { name: "Calibri", sz: 14, bold: true, color: { rgb: "FFFFFF" } },
  fill: { fgColor: { rgb: "538DD5" } },
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
};

const estiloVendaCollemCabecalho = {
  font: { name: "Calibri", sz: 12, bold: true },
  fill: { fgColor: { rgb: "DCE6F1" } },
  alignment: { vertical: "center" },
};

const estiloVendaCollemGrupo = {
  font: { name: "Calibri", sz: 12, bold: true },
  fill: { fgColor: { rgb: "C5D9F1" } },
  alignment: { vertical: "center" },
};

const estiloVendaCollemTotal = {
  font: { name: "Calibri", sz: 12, bold: true, color: { rgb: "FFFFFF" } },
  fill: { fgColor: { rgb: "538DD5" } },
  alignment: { horizontal: "center", vertical: "center" },
};

const aplicarEstiloLinha = (ws, row, startCol, endCol, style) => {
  for (let col = startCol; col <= endCol; col += 1) {
    const addr = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
    if (!ws[addr]) ws[addr] = { t: "s", v: "" };
    ws[addr].s = style;
  }
};

const aplicarFormatoNumerico = (ws, row, cols, formato) => {
  cols.forEach((col) => {
    const addr = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
    if (ws[addr]) ws[addr].z = formato;
  });
};

const aplicarAlinhamento = (ws, row, cols, alignment) => {
  cols.forEach((col) => {
    const addr = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
    if (!ws[addr]) return;
    ws[addr].s = {
      ...(ws[addr].s || {}),
      alignment: { ...(ws[addr].s?.alignment || {}), ...alignment },
    };
  });
};

export const criarAbaVendaModelo = (
  grupos,
  fatorVenda = 1,
  modelo = "alpha",
  descontoNegociacao = 0
) => {
  const collem = modelo === "collem";
  const estilos = collem
    ? {
        base: estiloVendaCollemBase,
        titulo: estiloVendaCollemTitulo,
        cabecalho: estiloVendaCollemCabecalho,
        grupo: estiloVendaCollemGrupo,
        total: estiloVendaCollemTotal,
      }
    : {
        base: estiloVendaBase,
        titulo: estiloVendaTitulo,
        cabecalho: estiloVendaCabecalho,
        grupo: estiloVendaGrupo,
        total: estiloVendaTotal,
      };
  const rows = [[], [null, "PLANILHA DE MATERIAL"], [null, "ITEM", "DESCRIÇÃO DOS SERVIÇOS", "UNID.", "QUANT.", "VALOR UNIT.", "VALOR TOTAL", "TOTAL DO ITEM"]];
  const groupRows = [];
  const itemRows = [];

  grupos.forEach((grupo) => {
    const groupRowNumber = rows.length + 1;
    const itemStartRow = groupRowNumber + 1;
    rows.push([null, grupo.numero, grupo.nome, null, null, null, null, grupo.total]);
    groupRows.push({ row: groupRowNumber, itemStartRow });

    grupo.itens.forEach((item) => {
      const itemRowNumber = rows.length + 1;
      rows.push([null, item.numero, item.descricao, item.unidade, item.quantidade, item.unitario, item.total, null]);
      itemRows.push(itemRowNumber);
    });

    const itemEndRow = rows.length;
    const group = groupRows[groupRows.length - 1];
    group.itemEndRow = itemEndRow;
    rows[groupRowNumber - 1][7] = grupo.total;
  });

  rows.push([]);
  const totalBruto = somarMoeda(grupos.map((grupo) => grupo.total));
  const descontoAplicado = Math.min(
    totalBruto,
    Math.max(0, arredondarMoeda(num(descontoNegociacao)))
  );
  let descontoRow = 0;
  if (descontoAplicado > 0) {
    descontoRow = rows.length + 1;
    rows.push([null, "DESCONTO DA NEGOCIAÇÃO", null, null, null, null, null, -descontoAplicado]);
  }
  const totalRow = rows.length + 1;
  const totalGeral = aplicarDescontoNegociacao(totalBruto, descontoAplicado);
  validarProposta(grupos, totalGeral, descontoAplicado);
  rows.push([null, "TOTAL GERAL", null, null, null, null, null, totalGeral]);

  const ws = XLSX.utils.aoa_to_sheet(rows);
  // As células exibem duas casas, mas preservam os valores originais nas
  // fórmulas. Somente o total de cada CPU é arredondado antes da soma.
  itemRows.forEach((row) => { ws[`G${row}`].f = `ROUND(E${row}*F${row},2)`; });
  groupRows.forEach(({ row, itemStartRow, itemEndRow }) => {
    ws[`H${row}`].f = itemEndRow >= itemStartRow ? `ROUND(SUM(G${itemStartRow}:G${itemEndRow}),2)` : "0";
  });
  const totalAddr = XLSX.utils.encode_cell({ r: totalRow - 1, c: 7 });
  const ultimaLinhaValores = Math.max(4, totalRow - 1);
  ws[totalAddr] = {
    t: "n",
    v: totalGeral,
    f: `MAX(0,ROUND(SUM(H4:H${ultimaLinhaValores}),2))`,
    z: XLSX_MOEDA,
  };
  ws["!merges"] = [
    { s: { r: 1, c: 1 }, e: { r: 1, c: 7 } },
    ...(descontoRow > 0
      ? [{ s: { r: descontoRow - 1, c: 1 }, e: { r: descontoRow - 1, c: 6 } }]
      : []),
    { s: { r: totalRow - 1, c: 1 }, e: { r: totalRow - 1, c: 6 } },
  ];
  ws["!cols"] = [
    { wch: 8.88671875 },
    { wch: 5.33203125 },
    { wch: 56.33203125 },
    { wch: 6.33203125 },
    { wch: 8 },
    { wch: 13.21875 },
    { wch: 14 },
    { wch: 15.5546875 },
    { wch: 4.5546875 },
    { wch: 4 },
  ];

  ws["!rows"] = rows.map(() => ({ hpt: 14.25 }));
  aplicarEstiloLinha(ws, 2, 2, 8, estilos.titulo);
  aplicarEstiloLinha(ws, 3, 2, 8, estilos.cabecalho);
  aplicarAlinhamento(ws, 3, [2, 4, 5, 6, 7, 8], { horizontal: "center" });
  aplicarAlinhamento(ws, 3, [3], { horizontal: "left", wrapText: true });
  aplicarFormatoNumerico(ws, 3, [8], XLSX_MOEDA);
  aplicarFormatoNumerico(ws, 3, [9], "0.00");

  groupRows.forEach(({ row }) => {
    aplicarEstiloLinha(ws, row, 2, 8, estilos.grupo);
    aplicarAlinhamento(ws, row, [2, 4, 5], { horizontal: "center" });
    aplicarAlinhamento(ws, row, [3], { horizontal: "left", wrapText: true });
    aplicarAlinhamento(ws, row, [8], { horizontal: "right" });
    aplicarFormatoNumerico(ws, row, [8], XLSX_MOEDA);
  });

  itemRows.forEach((row) => {
    aplicarEstiloLinha(ws, row, 2, 8, estilos.base);
    aplicarAlinhamento(ws, row, [2, 4, 5], { horizontal: "center" });
    aplicarAlinhamento(ws, row, [3], { horizontal: "left", wrapText: true });
    aplicarAlinhamento(ws, row, [6, 7, 8], { horizontal: "right" });
    aplicarFormatoNumerico(ws, row, [5], XLSX_NUMERO);
    aplicarFormatoNumerico(ws, row, [6], XLSX_NUMERO);
    aplicarFormatoNumerico(ws, row, [7, 8], XLSX_MOEDA);
  });

  if (descontoRow > 0) {
    aplicarEstiloLinha(ws, descontoRow, 2, 8, estilos.grupo);
    aplicarAlinhamento(ws, descontoRow, [2], { horizontal: "left" });
    aplicarAlinhamento(ws, descontoRow, [8], { horizontal: "right" });
    aplicarFormatoNumerico(ws, descontoRow, [8], XLSX_MOEDA);
  }

  aplicarEstiloLinha(ws, totalRow, 2, 8, estilos.total);
  aplicarFormatoNumerico(ws, totalRow, [8], XLSX_MOEDA);

  return ws;
};

export { XLSX_MOEDA };
