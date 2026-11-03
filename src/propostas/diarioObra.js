import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  ImageRun,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from "docx";
import alphaLogo from "../assets/alpha-engenharia-logo.png";

const VERDE = "7B9A56";
const VERDE_CLARO = "E2EFD9";
const CINZA = "666666";
const PRETO = "111111";
const BRANCO = "FFFFFF";
const LARGURA = 10490;
const SEM_BORDA = { style: BorderStyle.NONE, size: 0, color: BRANCO };
const BORDAS_LIMPAS = { top: SEM_BORDA, bottom: SEM_BORDA, left: SEM_BORDA, right: SEM_BORDA, insideHorizontal: SEM_BORDA, insideVertical: SEM_BORDA };
const BORDAS = {
  top: { style: BorderStyle.SINGLE, size: 4, color: "CFCFCF" },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: "CFCFCF" },
  left: { style: BorderStyle.SINGLE, size: 4, color: "CFCFCF" },
  right: { style: BorderStyle.SINGLE, size: 4, color: "CFCFCF" },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "CFCFCF" },
  insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "CFCFCF" },
};

const run = (text, options = {}) => new TextRun({ text: String(text ?? ""), font: "Arial", size: options.size || 20, bold: options.bold, color: options.color || PRETO });
const paragraph = (children, options = {}) => new Paragraph({ children: Array.isArray(children) ? children : [run(children, options)], alignment: options.alignment || AlignmentType.LEFT, spacing: { before: options.before || 0, after: options.after ?? 100, line: options.line || 260 } });
const safeName = (value) => String(value || "RDO").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9_-]+/g, "_").replace(/^_+|_+$/g, "") || "RDO";
const moneyless = (value) => String(value || "");

const cell = (text, options = {}) => new TableCell({ children: [paragraph(text, { size: options.size || 18, bold: options.bold, color: options.color, alignment: options.alignment })], width: options.width ? { size: options.width, type: WidthType.DXA } : undefined, columnSpan: options.columnSpan, shading: options.fill ? { fill: options.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, borders: options.borders || BORDAS, verticalAlign: VerticalAlign.CENTER, margins: { top: 70, bottom: 70, left: 80, right: 80 } });
const sectionTitle = (text) => new Paragraph({ children: [run(text, { bold: true, size: 22, color: PRETO })], shading: { fill: VERDE_CLARO, type: ShadingType.CLEAR, color: "auto" }, spacing: { before: 220, after: 100 }, indent: { left: 80 } });
const download = (blob, filename) => { const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.href = url; link.download = filename; document.body.appendChild(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000); };

const header = (logo, titulo) => new Header({ children: [
  new Table({ width: { size: LARGURA, type: WidthType.DXA }, columnWidths: [1500, 5990, 3000], borders: BORDAS_LIMPAS, rows: [new TableRow({ children: [
    new TableCell({ children: [new Paragraph({ children: [new ImageRun({ data: logo, transformation: { width: 76, height: 58 }, type: "png" })] })], borders: BORDAS_LIMPAS }),
    new TableCell({ children: [paragraph("ALPHA ENGENHARIA", { bold: true, size: 30, color: VERDE }), paragraph("Ordem de serviço e acompanhamento de obra", { size: 16, color: CINZA, after: 0 })], borders: BORDAS_LIMPAS, verticalAlign: VerticalAlign.CENTER }),
    new TableCell({ children: [paragraph("RELATÓRIO DIÁRIO DE OBRA", { bold: true, size: 22, alignment: AlignmentType.RIGHT, color: VERDE }), paragraph(titulo, { size: 16, alignment: AlignmentType.RIGHT, color: CINZA, after: 0 })], borders: BORDAS_LIMPAS, verticalAlign: VerticalAlign.CENTER }),
  ] })] }),
  new Paragraph({ children: [run("", { size: 8 })], border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: VERDE } }, spacing: { after: 100 } }),
] });

export const exportarDiarioObraDocx = async ({ projeto, cliente, registro, titulo, ferramentas = [], alphaDados }) => {
  const resposta = await fetch(alphaLogo);
  if (!resposta.ok) throw new Error("Não foi possível carregar a logo da ALPHA.");
  const logo = new Uint8Array(await resposta.arrayBuffer());
  const nomeRdo = titulo || registro.titulo || `RDO - ${projeto?.nome || "Obra"}`;
  const funcionarios = (registro.funcionarios || []).filter((item) => item.nomes || item.quantidade);
  const linhas = funcionarios.length ? funcionarios.map((item) => new TableRow({ cantSplit: true, children: [
    cell(item.funcao, { width: 2200, bold: true }),
    new TableCell({ children: String(item.nomes || "Nome não informado").split(",").map((nome) => paragraph(`${nome.trim()};`, { size: 18, after: 20 })), width: { size: 6790, type: WidthType.DXA }, borders: BORDAS, verticalAlign: VerticalAlign.TOP, margins: { top: 70, bottom: 70, left: 80, right: 80 } }),
    cell(item.quantidade || "0", { width: 1500, alignment: AlignmentType.CENTER }),
  ] })) : [new TableRow({ children: [cell("Nenhum profissional informado.", { columnSpan: 3 })] })];
  const tabelaFuncionarios = new Table({ width: { size: LARGURA, type: WidthType.DXA }, layout: TableLayoutType.FIXED, columnWidths: [2200, 6790, 1500], borders: BORDAS, rows: [new TableRow({ children: [cell("FUNÇÃO", { fill: VERDE, color: BRANCO, bold: true, width: 2200 }), cell("PROFISSIONAIS", { fill: VERDE, color: BRANCO, bold: true, width: 6790 }), cell("QTD.", { fill: VERDE, color: BRANCO, bold: true, width: 1500, alignment: AlignmentType.CENTER })] }), ...linhas] });
  const linhasDados = [
    ["Contratada / empreiteira", alphaDados?.nome || "Alpha Engenharia e Serviços Ltda"],
    ["CNPJ da contratada", alphaDados?.documento || "Não informado"],
    ["Cliente / contratante", cliente?.nome || "Não informado"],
    ["CNPJ do cliente", cliente?.documento || "Não informado"],
    ["Contrato nº", cliente?.numeroProposta || "A definir"],
    ["Obra", projeto?.nome || "Orçamento sem nome"],
    ["Endereço", cliente?.endereco || cliente?.local || "A informar"],
    ["Data", registro.data || "Não informada"],
  ];
  const dadosTable = new Table({ width: { size: LARGURA, type: WidthType.DXA }, layout: TableLayoutType.FIXED, columnWidths: [2200, 3045, 2200, 3045], borders: BORDAS, rows: [0, 2, 4, 6].map((index) => new TableRow({ children: [cell(linhasDados[index][0], { fill: "F1F4EE", bold: true, width: 2200, size: 16 }), cell(linhasDados[index][1], { width: 3045, size: 17 }), cell(linhasDados[index + 1][0], { fill: "F1F4EE", bold: true, width: 2200, size: 16 }), cell(linhasDados[index + 1][1], { width: 3045, size: 17 })] })) });
  const ferramentasTexto = ferramentas.length ? ferramentas.map((item) => `${item.nome} (Qtd.: ${item.quantidade || "1"})`).join("; ") : "Nenhuma ferramenta informada.";
  const documento = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1500, right: 707, bottom: 850, left: 707, header: 0, footer: 0 } } }, headers: { default: header(logo, nomeRdo) }, footers: { default: new Footer({ children: [paragraph(" ", { after: 0 })] }) }, children: [
    sectionTitle("Identificação do contrato"), dadosTable,
    sectionTitle("Mão de obra presente"), tabelaFuncionarios,
    sectionTitle("Ferramentas utilizadas"), paragraph(moneyless(ferramentasTexto), { after: 160 }),
    sectionTitle("Atividades executadas"), paragraph(registro.atividades || registro.descricao || "Não informado", { after: 180 }),
    new Table({ width: { size: LARGURA, type: WidthType.DXA }, borders: BORDAS_LIMPAS, columnWidths: [5245, 5245], rows: [new TableRow({ children: [cell("Responsável pela execução", { borders: { top: { style: BorderStyle.SINGLE, size: 4, color: CINZA }, bottom: SEM_BORDA, left: SEM_BORDA, right: SEM_BORDA }, alignment: AlignmentType.CENTER, size: 16 }), cell("Responsável pelo acompanhamento", { borders: { top: { style: BorderStyle.SINGLE, size: 4, color: CINZA }, bottom: SEM_BORDA, left: SEM_BORDA, right: SEM_BORDA }, alignment: AlignmentType.CENTER, size: 16 })] })] }),
  ] }] });
  const blob = await Packer.toBlob(documento);
  download(blob, `${safeName(nomeRdo)}.docx`);
};
