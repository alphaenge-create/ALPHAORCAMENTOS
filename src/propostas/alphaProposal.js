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
const FONTE = "Calibri";
const LARGURA_PAGINA = 11906;
const LARGURA_CONTEUDO = 10490;
const MARGEM_ESQUERDA = 709;
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
    size: opcoes.size || 22,
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
    width: { size: LARGURA_PAGINA, type: WidthType.DXA },
    indent: { size: -MARGEM_ESQUERDA, type: WidthType.DXA },
    columnWidths: [LARGURA_PAGINA],
    borders: BORDAS_LIMPAS,
    rows: [
      new TableRow({
        children: [celula("", { fill: VERDE, borders: BORDAS_LIMPAS, size: 2, width: LARGURA_PAGINA })],
        height: { value: 450, rule: "atLeast" },
      }),
    ],
  });

const separadorTabelas = () =>
  new Paragraph({
    children: [run("", { size: 2 })],
    spacing: { before: 0, after: 0, line: 2 },
  });

const criarCabecalho = (logo, numeroProposta) =>
  new Header({
    children: [
      faixa(),
      separadorTabelas(),
      new Table({
        width: { size: LARGURA_CONTEUDO, type: WidthType.DXA },
        columnWidths: [7500, 2990],
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
                        transformation: { width: 96, height: 74 },
                        type: "png",
                      }),
                    ],
                    spacing: { before: 240, after: 60 },
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
        width: { size: LARGURA_CONTEUDO, type: WidthType.DXA },
        columnWidths: [7700, 2790],
        borders: BORDAS_LIMPAS,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: BORDAS_LIMPAS,
                children: [
                  new Paragraph({
                    children: [run("ALPHA ENGENHARIA E SERVIÇOS", { bold: true, size: 20, color: CINZA })],
                    spacing: { after: 0 },
                  }),
                  new Paragraph({
                    children: [run("Rua José Da Costa, 116 – São João Batista", { size: 20, color: CINZA })],
                    spacing: { after: 0 },
                  }),
                  new Paragraph({
                    children: [run("Belo Horizonte", { size: 20, color: CINZA })],
                    spacing: { after: 0 },
                  }),
                  new Paragraph({
                    children: [run("Telefone: 31 9 9203-1783", { size: 20, color: CINZA })],
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
                        size: 20,
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
      separadorTabelas(),
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
    width: { size: LARGURA_CONTEUDO, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    columnWidths: [3100, 2800, 2050, 2540],
    borders: BORDAS_TABELA,
    rows: linhas,
  });
};

const criarTabelaValores = (grupos, totalGeral, descontoNegociacao = 0) => {
  const larguras = [680, 4380, 730, 870, 1130, 1130, 1570];
  const linhas = [
    new TableRow({
      tableHeader: true,
      cantSplit: true,
      children: [
        celula("PLANILHA DE MATERIAL", {
          columnSpan: 7,
          fill: VERDE,
          color: BRANCO,
          bold: true,
          size: 22,
          alignment: AlignmentType.CENTER,
        }),
      ],
    }),
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

  if (descontoNegociacao > 0) {
    linhas.push(
      new TableRow({
        cantSplit: true,
        children: [
          celula("DESCONTO DA NEGOCIAÇÃO", {
            columnSpan: 6,
            fill: VERDE_CLARO,
            bold: true,
            alignment: AlignmentType.LEFT,
            size: 20,
          }),
          celula(`- R$ ${moeda(descontoNegociacao)}`, {
            fill: VERDE_CLARO,
            bold: true,
            alignment: AlignmentType.RIGHT,
            size: 20,
          }),
        ],
      })
    );
  }

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
    width: { size: LARGURA_CONTEUDO, type: WidthType.DXA },
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
  const descontoNegociacao = Math.max(0, Number(dados.descontoNegociacao || 0));
  const abertura = [
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
  ];
  const corpo = [
    tituloSecao("Responsabilidade da ALPHA ENGENHARIA:", { before: 200 }),
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
    corpo.push(tituloSecao("ALTERNATIVAS TÉCNICAS"));
    corpo.push(criarTabelaAlternativas(dados.comparativos));
  }

  corpo.push(criarTabelaValores(grupos, totalGeral, descontoNegociacao));
  const fechamento = [
    tituloSecao("Condições de pagamento:", { before: 200 }),
    paragrafo(dados.condicoesPagamento || "A definir entre as partes."),
    paragrafo("Pagamento via PIX (52.903.822/0001-86) 5 dias após a emissão da NF."),
    tituloSecao("Prazo para Execução:"),
    itemLista(dados.prazoExecucao || "A definir conforme cronograma aprovado entre as partes."),
  ];
  if (dados.observacoes) {
    fechamento.push(tituloSecao("Observações:"));
    String(dados.observacoes).split(/\r?\n/).filter(Boolean).forEach((linha) => fechamento.push(paragrafo(linha)));
  }
  fechamento.push(
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
    evenAndOddHeaderAndFooters: true,
    styles: {
      default: {
        document: {
          run: { font: FONTE, size: 22, color: PRETO },
          paragraph: { spacing: { line: 276 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: {
              top: 1843,
              right: 707,
              bottom: 851,
              left: MARGEM_ESQUERDA,
              header: 0,
              footer: 0,
            },
          },
        },
        headers: {
          default: criarCabecalho(logo, dados.numeroProposta || ""),
          even: criarCabecalho(logo, dados.numeroProposta || ""),
        },
        footers: {
          default: criarRodape(),
          even: criarRodape(),
        },
        children: [...abertura, ...corpo, ...fechamento],
      },
    ],
  });

  return Packer.toBlob(documento);
};

export const gerarPropostaAlphaDocx = async (dados) => {
  const blob = await criarPropostaAlphaDocxBlob(dados);
  baixarBlob(blob, `${nomeArquivoSeguro(dados.nomeProjeto)}_Proposta_ALPHA.docx`);
};
