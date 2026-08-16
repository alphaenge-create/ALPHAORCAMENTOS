import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  ImageRun,
  Packer,
  PageNumber,
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
const CINZA = "A6A6A6";
const BRANCO = "FFFFFF";
const PRETO = "111111";
const FONTE = "Arial Narrow";
const SEM_BORDA = { style: BorderStyle.NONE, size: 0, color: BRANCO };
const BORDAS_LIMPAS = {
  top: SEM_BORDA,
  bottom: SEM_BORDA,
  left: SEM_BORDA,
  right: SEM_BORDA,
  insideHorizontal: SEM_BORDA,
  insideVertical: SEM_BORDA,
};
const BORDAS_TABELA = {
  top: { style: BorderStyle.SINGLE, size: 2, color: "E7E7E7" },
  bottom: { style: BorderStyle.SINGLE, size: 2, color: "E7E7E7" },
  left: SEM_BORDA,
  right: SEM_BORDA,
  insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: "E7E7E7" },
  insideVertical: SEM_BORDA,
};

const moeda = (valor) =>
  Number(valor || 0).toLocaleString("pt-BR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

const nomeArquivoSeguro = (valor) =>
  String(valor || "Orcamento")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9_-]+/g, "_")
    .replace(/^_+|_+$/g, "") || "Orcamento";

const baixarBlob = (blob, nome) => {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = nome;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
};

const run = (texto, opcoes = {}) =>
  new TextRun({
    text: String(texto ?? ""),
    font: FONTE,
    size: opcoes.size || 24,
    bold: opcoes.bold,
    color: opcoes.color || PRETO,
  });

const paragrafo = (conteudo, opcoes = {}) =>
  new Paragraph({
    children: Array.isArray(conteudo) ? conteudo : [run(conteudo, opcoes)],
    alignment: opcoes.alignment || AlignmentType.JUSTIFIED,
    spacing: {
      before: opcoes.before || 0,
      after: opcoes.after ?? 100,
      line: opcoes.line || 300,
    },
    keepNext: opcoes.keepNext,
    pageBreakBefore: opcoes.pageBreakBefore,
  });

const tituloSecao = (texto, opcoes = {}) =>
  paragrafo([run(texto, { bold: true, size: opcoes.size || 24 })], {
    alignment: AlignmentType.LEFT,
    before: opcoes.before ?? 160,
    after: opcoes.after ?? 80,
    line: 276,
    keepNext: true,
    pageBreakBefore: opcoes.pageBreakBefore,
  });

const itemLista = (texto, opcoes = {}) =>
  new Paragraph({
    children: [run(texto, { bold: opcoes.bold })],
    bullet: { level: 0 },
    spacing: { after: opcoes.after ?? 45, line: opcoes.line || 276 },
  });

const celula = (conteudo, opcoes = {}) =>
  new TableCell({
    children: [
      new Paragraph({
        children: [run(conteudo, {
          bold: opcoes.bold,
          size: opcoes.size || 18,
          color: opcoes.color || PRETO,
        })],
        alignment: opcoes.alignment || AlignmentType.LEFT,
        spacing: { before: 0, after: 0, line: 220 },
      }),
    ],
    width: opcoes.width ? { size: opcoes.width, type: WidthType.DXA } : undefined,
    columnSpan: opcoes.columnSpan,
    shading: opcoes.fill
      ? { fill: opcoes.fill, type: ShadingType.CLEAR, color: "auto" }
      : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 55, bottom: 55, left: 65, right: 65 },
    borders: opcoes.borders || BORDAS_TABELA,
  });

const faixa = () =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: BORDAS_LIMPAS,
    rows: [
      new TableRow({
        children: [celula("", { fill: VERDE, borders: BORDAS_LIMPAS, size: 2 })],
        height: { value: 180, rule: "atLeast" },
      }),
    ],
  });

const criarCabecalho = (logo, numeroProposta) =>
  new Header({
    children: [
      faixa(),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [6800, 2800],
        borders: BORDAS_LIMPAS,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    children: [
                      new ImageRun({
                        data: logo,
                        transformation: { width: 112, height: 86 },
                        type: "png",
                      }),
                    ],
                    spacing: { before: 120, after: 60 },
                  }),
                ],
                borders: BORDAS_LIMPAS,
                verticalAlign: VerticalAlign.CENTER,
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    children: [run(numeroProposta, { bold: true, color: CINZA })],
                    alignment: AlignmentType.RIGHT,
                    spacing: { before: 520, after: 0 },
                  }),
                ],
                borders: BORDAS_LIMPAS,
                verticalAlign: VerticalAlign.CENTER,
              }),
            ],
          }),
        ],
      }),
    ],
  });

const criarRodape = () =>
  new Footer({
    children: [
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [7200, 2400],
        borders: BORDAS_LIMPAS,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: BORDAS_LIMPAS,
                children: [
                  new Paragraph({
                    children: [run("ALPHA ENGENHARIA E SERVIÇOS", { bold: true, size: 18, color: CINZA })],
                    spacing: { after: 0 },
                  }),
                  new Paragraph({
                    children: [run("Rua José Da Costa, 116 - São João Batista | Belo Horizonte | Telefone: 31 9 9203-1783", { size: 18, color: CINZA })],
                    spacing: { after: 0 },
                  }),
                ],
              }),
              new TableCell({
                borders: BORDAS_LIMPAS,
                verticalAlign: VerticalAlign.BOTTOM,
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        children: ["Página ", PageNumber.CURRENT, " de ", PageNumber.TOTAL_PAGES],
                        font: FONTE,
                        size: 18,
                        color: CINZA,
                      }),
                    ],
                    alignment: AlignmentType.RIGHT,
                    spacing: { after: 0 },
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      faixa(),
    ],
  });

const criarTabelaAlternativas = (comparativos) => {
  const linhas = [
    new TableRow({
      tableHeader: true,
      cantSplit: true,
      children: [
        celula("GRUPO", { fill: VERDE, color: BRANCO, bold: true, width: 2900 }),
        celula("ALTERNATIVA", { fill: VERDE, color: BRANCO, bold: true, width: 2600 }),
        celula("SITUAÇÃO", { fill: VERDE, color: BRANCO, bold: true, width: 1900 }),
        celula("TOTAL DA PROPOSTA", { fill: VERDE, color: BRANCO, bold: true, width: 2200, alignment: AlignmentType.RIGHT }),
      ],
    }),
  ];
  comparativos.forEach((grupo) =>
    (grupo.opcoes || []).forEach((opcao) =>
      linhas.push(
        new TableRow({
          cantSplit: true,
          children: [
            celula(grupo.grupoNome, { fill: opcao.selecionada ? VERDE_CLARO : undefined }),
            celula(opcao.nome, { fill: opcao.selecionada ? VERDE_CLARO : undefined }),
            celula(opcao.selecionada ? "Considerada no total" : "Alternativa", { fill: opcao.selecionada ? VERDE_CLARO : undefined }),
            celula(`R$ ${moeda(opcao.valorVenda)}`, { fill: opcao.selecionada ? VERDE_CLARO : undefined, alignment: AlignmentType.RIGHT, bold: opcao.selecionada }),
          ],
        })
      )
    )
  );
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    columnWidths: [2900, 2600, 1900, 2200],
    borders: BORDAS_TABELA,
    rows: linhas,
  });
};

const criarTabelaValores = (grupos, totalGeral) => {
  const larguras = [650, 4000, 700, 850, 1050, 1050, 1300];
  const linhas = [
    new TableRow({
      tableHeader: true,
      cantSplit: true,
      children: [
        celula("ITEM", { fill: VERDE, color: BRANCO, bold: true, width: larguras[0] }),
        celula("DESCRIÇÃO DOS SERVIÇOS", { fill: VERDE, color: BRANCO, bold: true, width: larguras[1] }),
        celula("UNID.", { fill: VERDE, color: BRANCO, bold: true, width: larguras[2], alignment: AlignmentType.CENTER }),
        celula("QUANT.", { fill: VERDE, color: BRANCO, bold: true, width: larguras[3], alignment: AlignmentType.RIGHT }),
        celula("VALOR UNIT.", { fill: VERDE, color: BRANCO, bold: true, width: larguras[4], alignment: AlignmentType.RIGHT }),
        celula("VALOR TOTAL", { fill: VERDE, color: BRANCO, bold: true, width: larguras[5], alignment: AlignmentType.RIGHT }),
        celula("TOTAL DO ITEM", { fill: VERDE, color: BRANCO, bold: true, width: larguras[6], alignment: AlignmentType.RIGHT }),
      ],
    }),
  ];

  grupos.forEach((grupo) => {
    linhas.push(
      new TableRow({
        cantSplit: true,
        children: [
          celula(`${grupo.numero}.`, { fill: VERDE_CLARO, bold: true }),
          celula(grupo.nome, { fill: VERDE_CLARO, bold: true }),
          celula("", { fill: VERDE_CLARO }),
          celula("", { fill: VERDE_CLARO }),
          celula("", { fill: VERDE_CLARO }),
          celula("", { fill: VERDE_CLARO }),
          celula(`R$ ${moeda(grupo.total)}`, { fill: VERDE_CLARO, bold: true, alignment: AlignmentType.RIGHT }),
        ],
      })
    );
    (grupo.itens || []).forEach((item) =>
      linhas.push(
        new TableRow({
          cantSplit: true,
          children: [
            celula(item.numero),
            celula(item.descricao),
            celula(item.unidade, { alignment: AlignmentType.CENTER }),
            celula(moeda(item.quantidade), { alignment: AlignmentType.RIGHT }),
            celula(`R$ ${moeda(item.unitario)}`, { alignment: AlignmentType.RIGHT }),
            celula(`R$ ${moeda(item.total)}`, { alignment: AlignmentType.RIGHT }),
            celula("", { alignment: AlignmentType.RIGHT }),
          ],
        })
      )
    );
  });

  linhas.push(
    new TableRow({
      cantSplit: true,
      children: [
        celula("TOTAL GERAL", { columnSpan: 6, fill: VERDE, color: BRANCO, bold: true, alignment: AlignmentType.CENTER, size: 20 }),
        celula(`R$ ${moeda(totalGeral)}`, { fill: VERDE, color: BRANCO, bold: true, alignment: AlignmentType.RIGHT, size: 20 }),
      ],
    })
  );

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    columnWidths: larguras,
    borders: BORDAS_TABELA,
    rows: linhas,
  });
};

export const criarPropostaAlphaDocxBlob = async (dados, opcoes = {}) => {
  const respostaLogo = opcoes.logoData ? null : await fetch(alphaLogo);
  if (respostaLogo && !respostaLogo.ok) throw new Error("Não foi possível carregar a logo da ALPHA.");
  const logo = opcoes.logoData || new Uint8Array(await respostaLogo.arrayBuffer());
  const hoje = new Date();
  const dataHoje = hoje.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit", year: "numeric" });
  const grupos = dados.grupos || [];
  const totalGeral = Number(dados.totalGeral || 0);
  const conteudo = [
    paragrafo([run("PROPOSTA DE PRESTAÇÃO DE SERVIÇOS", { bold: true, size: 32 })], {
      alignment: AlignmentType.CENTER,
      before: 80,
      after: 600,
      line: 300,
    }),
    paragrafo(`Belo Horizonte ${dataHoje}`, { alignment: AlignmentType.RIGHT, after: 240 }),
    paragrafo(`Aos cuidados de ${dados.nomeCliente || "Cliente"}${dados.contato ? ` - ${dados.contato}` : ""}.`, { alignment: AlignmentType.LEFT, after: 180 }),
    paragrafo([run("Ref. ", { bold: true }), run(dados.nomeProjeto || "Orçamento", { bold: true })], { alignment: AlignmentType.LEFT, after: 180 }),
    paragrafo([run("Endereço da Obra: ", { bold: true }), run(dados.localObra || "")], { alignment: AlignmentType.LEFT, after: 220 }),
    tituloSecao("Escopo do Serviço:", { before: 0, after: 80 }),
    ...grupos.map((grupo) => itemLista(grupo.nome, { bold: true, after: 75, line: 290 })),
    tituloSecao("Responsabilidade da ALPHA ENGENHARIA:", { before: 0, pageBreakBefore: true }),
    ...(dados.responsabilidadesAlpha || []).map((item) => itemLista(item)),
    tituloSecao("Responsabilidade do Cliente:"),
    ...(dados.responsabilidadesCliente || []).map((item) => itemLista(item)),
    tituloSecao("Valores:"),
    paragrafo([
      run("Segue relação da mão de obra especializada para execução e acompanhamento dos serviços apresentados em visita técnica, totalizando o valor de "),
      run(`R$ ${moeda(totalGeral)}`, { bold: true }),
      run("."),
    ]),
    paragrafo([run("Condição dos materiais: ", { bold: true }), run(dados.descricaoRegimeMateriais || "")]),
  ];

  if ((dados.comparativos || []).length) {
    conteudo.push(tituloSecao("ALTERNATIVAS TÉCNICAS"));
    conteudo.push(criarTabelaAlternativas(dados.comparativos));
  }

  conteudo.push(tituloSecao("PLANILHA DE MATERIAL"));
  conteudo.push(criarTabelaValores(grupos, totalGeral));
  conteudo.push(tituloSecao("Condições de pagamento:", { before: 1050, pageBreakBefore: true }));
  conteudo.push(paragrafo(dados.condicoesPagamento || "A definir entre as partes."));
  conteudo.push(paragrafo("Pagamento via PIX (52.903.822/0001-86) 5 dias após a emissão da NF."));
  conteudo.push(tituloSecao("Prazo para Execução:"));
  conteudo.push(itemLista(dados.prazoExecucao || "A definir conforme cronograma aprovado entre as partes."));
  if (dados.observacoes) {
    conteudo.push(tituloSecao("Observações:"));
    String(dados.observacoes).split(/\r?\n/).filter(Boolean).forEach((linha) => conteudo.push(paragrafo(linha)));
  }
  conteudo.push(
    new Paragraph({
      children: [run("ALPHA ENGENHARIA E SERVIÇOS", { bold: true })],
      alignment: AlignmentType.CENTER,
      border: { top: { style: BorderStyle.SINGLE, size: 6, color: PRETO, space: 8 } },
      indent: { left: 3000, right: 3000 },
      spacing: { before: 900, after: 0 },
    })
  );

  const documento = new Document({
    features: { updateFields: true },
    styles: {
      default: {
        document: {
          run: { font: FONTE, size: 24, color: PRETO },
          paragraph: { spacing: { line: 300 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: {
              top: 1701,
              right: 709,
              bottom: 1417,
              left: 709,
              header: 0,
              footer: 0,
            },
          },
        },
        headers: { default: criarCabecalho(logo, dados.numeroProposta || "") },
        footers: { default: criarRodape() },
        children: conteudo,
      },
    ],
  });

  return Packer.toBlob(documento);
};

export const gerarPropostaAlphaDocx = async (dados) => {
  const blob = await criarPropostaAlphaDocxBlob(dados);
  baixarBlob(blob, `${nomeArquivoSeguro(dados.nomeProjeto)}_Proposta_ALPHA.docx`);
};
