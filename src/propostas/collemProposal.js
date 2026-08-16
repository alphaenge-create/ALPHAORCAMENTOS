import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  ImageRun,
  Packer,
  PageBreak,
  Paragraph,
  TextRun,
} from "docx";
import collemLogo from "../assets/collem-logo.png";

const COLLEM = {
  nome: "COLLEM CONSTRUTORA MOHALLEM LTDA",
  responsavel: "Geraldo Belloni Perez",
  cnpj: "21.442.256/0001-29",
  endereco: "Rua Canoas, 719, Bairro Betânia - Belo Horizonte/MG.",
  telefone: "(31) 9 9203-1783",
  rodapeEndereco:
    "Rua Canoas, 719, Bairro Betânia - Belo Horizonte - MG cep: 30580 - 040 Tel.: (031) 3303-1999",
  email: "collemconstrutora@collem.com.br",
  site: "www.collem.com.br",
};

const unidades = [
  "zero",
  "um",
  "dois",
  "três",
  "quatro",
  "cinco",
  "seis",
  "sete",
  "oito",
  "nove",
];
const especiais = [
  "dez",
  "onze",
  "doze",
  "treze",
  "quatorze",
  "quinze",
  "dezesseis",
  "dezessete",
  "dezoito",
  "dezenove",
];
const dezenas = [
  "",
  "",
  "vinte",
  "trinta",
  "quarenta",
  "cinquenta",
  "sessenta",
  "setenta",
  "oitenta",
  "noventa",
];
const centenas = [
  "",
  "cento",
  "duzentos",
  "trezentos",
  "quatrocentos",
  "quinhentos",
  "seiscentos",
  "setecentos",
  "oitocentos",
  "novecentos",
];
const escalas = [
  ["", ""],
  ["mil", "mil"],
  ["milhão", "milhões"],
  ["bilhão", "bilhões"],
  ["trilhão", "trilhões"],
];

const numeroAte999PorExtenso = (valor) => {
  const numero = Math.trunc(valor);
  if (numero === 0) return "zero";
  if (numero === 100) return "cem";

  const partes = [];
  const centena = Math.trunc(numero / 100);
  const resto = numero % 100;
  if (centena) partes.push(centenas[centena]);
  if (resto >= 10 && resto < 20) {
    partes.push(especiais[resto - 10]);
  } else {
    const dezena = Math.trunc(resto / 10);
    const unidade = resto % 10;
    if (dezena) partes.push(dezenas[dezena]);
    if (unidade) partes.push(unidades[unidade]);
  }
  return partes.join(" e ");
};

const inteiroPorExtenso = (valor) => {
  const numero = Math.max(0, Math.trunc(valor));
  if (numero === 0) return "zero";

  const grupos = [];
  let restante = numero;
  let escala = 0;
  while (restante > 0 && escala < escalas.length) {
    const grupo = restante % 1000;
    if (grupo) grupos.unshift({ grupo, escala });
    restante = Math.trunc(restante / 1000);
    escala += 1;
  }

  return grupos
    .map(({ grupo, escala: indice }, posicao) => {
      let texto;
      if (indice === 1 && grupo === 1) texto = "mil";
      else {
        texto = numeroAte999PorExtenso(grupo);
        if (indice > 0) {
          texto += ` ${escalas[indice][grupo === 1 ? 0 : 1]}`;
        }
      }

      if (posicao === 0) return texto;
      const usarE = grupo < 100 || grupo % 100 === 0;
      return `${usarE ? "e " : ""}${texto}`;
    })
    .join(" ");
};

const moedaPorExtenso = (valor) => {
  const totalCentavos = Math.round((Number(valor) || 0) * 100);
  const reais = Math.trunc(totalCentavos / 100);
  const centavos = totalCentavos % 100;
  const partes = [];
  if (reais) {
    partes.push(`${inteiroPorExtenso(reais)} ${reais === 1 ? "real" : "reais"}`);
  }
  if (centavos) {
    partes.push(
      `${inteiroPorExtenso(centavos)} ${centavos === 1 ? "centavo" : "centavos"}`
    );
  }
  const texto = partes.length ? partes.join(" e ") : "zero reais";
  return texto.charAt(0).toUpperCase() + texto.slice(1);
};

const formatarMoeda = (valor) =>
  Number(valor || 0).toLocaleString("pt-BR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

const formatarPercentual = (valor) =>
  Number(valor || 0).toLocaleString("pt-BR", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2,
  });

const formatarDataLonga = (data = new Date()) =>
  data.toLocaleDateString("pt-BR", {
    day: "2-digit",
    month: "long",
    year: "numeric",
  });

const escaparHtml = (valor) =>
  String(valor ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");

const nomeArquivoSeguro = (valor) =>
  String(valor || "Orcamento")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9-_ ]/g, "")
    .trim()
    .replace(/\s+/g, "_");

const textoRegimeMateriais = (cliente = {}) => {
  if (cliente.regimeMateriais === "cliente") {
    return "Não está incluso o fornecimento dos materiais, que ficará por conta do contratante.";
  }
  if (cliente.regimeMateriais === "faturamentoDireto") {
    return "Os materiais indicados no ANEXO I serão faturados diretamente ao contratante.";
  }
  return "O fornecimento dos materiais está incluso conforme especificado no ANEXO I.";
};

const dadosProposta = ({ projeto, cliente = {}, totalGeral }) => {
  const nomeProjeto = projeto?.nome || "serviços contratados";
  const percentualSinal = Math.min(
    100,
    Math.max(0, Number(cliente.percentualSinalCollem ?? 20) || 0)
  );
  const valorSinal = totalGeral * (percentualSinal / 100);
  const nomeCliente = cliente.nome || "Cliente";
  const contato = cliente.contato || nomeCliente;
  const enderecoCliente = String(cliente.endereco || "").trim();
  const localCliente = String(cliente.local || "").trim();
  const introducao =
    String(cliente.textoApresentacaoCollem || "").trim() ||
    `Atendendo à vossa solicitação, estamos encaminhando nossa proposta para ${nomeProjeto.toLowerCase()}, conforme serviços especificados no ANEXO I.`;
  const condicoesPagamento =
    String(cliente.condicoesPagamento || "").trim() ||
    `Sinal de negócio: ${formatarPercentual(percentualSinal)}% do valor do item A acima (R$ ${formatarMoeda(valorSinal)} - ${moedaPorExtenso(valorSinal)}) na assinatura do contrato. O restante em pagamentos conforme medições mensais.`;
  const prazo =
    String(cliente.prazoExecucaoCollem || "").trim() ||
    String(cliente.prazoExecucao || "").trim() ||
    "A definir conforme cronograma aprovado entre as partes.";
  const naoInclusos =
    String(cliente.naoInclusosCollem || "").trim() || textoRegimeMateriais(cliente);
  const condicoesEspeciais =
    String(cliente.condicoesEspeciaisCollem || "").trim() ||
    String(cliente.observacoes || "").trim() ||
    "As condições de acesso, apoio e execução serão alinhadas entre as partes antes do início dos serviços.";

  return {
    nomeProjeto,
    nomeCliente,
    contato,
    enderecoCliente,
    localCliente,
    introducao,
    condicoesPagamento,
    prazo,
    naoInclusos,
    condicoesEspeciais,
    totalGeral,
    totalExtenso: moedaPorExtenso(totalGeral),
    data: formatarDataLonga(),
    responsavel: String(cliente.responsavelCollem || "").trim() || COLLEM.responsavel,
  };
};

const baixarBlob = (blob, nomeArquivo) => {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = nomeArquivo;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
};

const run = (texto, opcoes = {}) =>
  new TextRun({
    text: String(texto || ""),
    font: "Arial Narrow",
    size: 24,
    bold: !!opcoes.bold,
  });

const paragrafo = (filhos, opcoes = {}) =>
  new Paragraph({
    children: Array.isArray(filhos) ? filhos : [run(filhos)],
    alignment: opcoes.alignment || AlignmentType.JUSTIFIED,
    indent: opcoes.firstLine === false ? undefined : { firstLine: 720 },
    spacing: {
      line: 360,
      before: opcoes.before ?? 0,
      after: opcoes.after ?? 120,
    },
    keepNext: !!opcoes.keepNext,
  });

const tituloSecao = (letra, titulo) =>
  paragrafo([run(`${letra}) ${titulo}`, { bold: true })], {
    firstLine: false,
    before: 240,
    after: 120,
    keepNext: true,
  });

const criarRodape = () =>
  new Footer({
    children: [
      new Paragraph({
        children: [
          new TextRun({ text: COLLEM.rodapeEndereco, font: "Arial", size: 18 }),
        ],
        alignment: AlignmentType.CENTER,
        border: {
          top: {
            color: "000000",
            space: 6,
            style: BorderStyle.SINGLE,
            size: 6,
          },
        },
        spacing: { before: 80, after: 0 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: `e-mail: ${COLLEM.email}   site: ${COLLEM.site}`,
            font: "Arial",
            size: 18,
            color: "0563C1",
            underline: {},
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
      }),
    ],
  });

export const gerarPropostaCollemDocx = async ({ projeto, cliente, totalGeral }) => {
  const dados = dadosProposta({ projeto, cliente, totalGeral });
  const respostaLogo = await fetch(collemLogo);
  if (!respostaLogo.ok) throw new Error("Não foi possível carregar a logo da COLLEM.");
  const logo = new Uint8Array(await respostaLogo.arrayBuffer());
  const destinatario = [
    dados.nomeCliente,
    ...dados.enderecoCliente.split(/\r?\n/).filter(Boolean),
    dados.localCliente,
  ].filter(Boolean);

  const documento = new Document({
    styles: {
      default: {
        document: {
          run: { font: "Arial Narrow", size: 24 },
          paragraph: { spacing: { line: 360 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: {
              top: 1417,
              right: 1701,
              bottom: 1417,
              left: 1701,
              header: 708,
              footer: 708,
            },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                children: [
                  new ImageRun({
                    data: logo,
                    transformation: { width: 120, height: 47 },
                    type: "png",
                  }),
                ],
                indent: { left: -1300 },
                spacing: { after: 0 },
              }),
            ],
          }),
        },
        footers: { default: criarRodape() },
        children: [
          paragrafo(`Belo Horizonte, ${dados.data}`, {
            firstLine: false,
            alignment: AlignmentType.LEFT,
            after: 480,
          }),
          ...destinatario.map((linha) =>
            paragrafo(linha, {
              firstLine: false,
              alignment: AlignmentType.RIGHT,
              after: 40,
            })
          ),
          paragrafo(
            [run("Ref.: ", { bold: true }), run(`Contrato de empreitada global de ${dados.nomeProjeto.toLowerCase()}.`)],
            { firstLine: false, alignment: AlignmentType.LEFT, before: 360, after: 240 }
          ),
          paragrafo(
            [run("Atte.: ", { bold: true }), run(`${dados.contato};`)],
            { firstLine: false, alignment: AlignmentType.LEFT, after: 240 }
          ),
          paragrafo("Prezados senhores,", { firstLine: false, after: 240 }),
          paragrafo(dados.introducao, { after: 240 }),
          tituloSecao("A", "Preço"),
          paragrafo(
            `O preço total para a execução dos trabalhos é de R$ ${formatarMoeda(
              dados.totalGeral
            )} (${dados.totalExtenso}) sendo especificados no ANEXO I;`,
            { after: 240 }
          ),
          tituloSecao("B", "Condições de pagamento"),
          paragrafo(dados.condicoesPagamento, { after: 0 }),
          new Paragraph({ children: [new PageBreak()] }),
          tituloSecao("C", "Prazo de Execução"),
          paragrafo(`Estimamos executar todos os trabalhos em ${dados.prazo}.`, {
            after: 180,
          }),
          tituloSecao("D", "Não inclusos"),
          paragrafo(dados.naoInclusos, { after: 180 }),
          tituloSecao("E", "Condições especiais"),
          paragrafo(dados.condicoesEspeciais, { after: 80 }),
          paragrafo("Colocamo-nos à sua inteira disposição para qualquer esclarecimento.", {
            after: 480,
          }),
          paragrafo("Atenciosamente.", {
            firstLine: false,
            alignment: AlignmentType.LEFT,
            after: 540,
          }),
          new Paragraph({
            children: [new TextRun({ text: dados.responsavel, font: "Arial", size: 18 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 80 },
          }),
          new Paragraph({
            children: [new TextRun({ text: `Contratada: ${COLLEM.nome}`, font: "Arial", size: 18 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 80 },
          }),
          new Paragraph({
            children: [new TextRun({ text: `CNPJ/MF: ${COLLEM.cnpj}`, font: "Arial", size: 18 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 80 },
          }),
          new Paragraph({
            children: [new TextRun({ text: `Endereço: ${COLLEM.endereco}`, font: "Arial", size: 18 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 80 },
          }),
          new Paragraph({
            children: [new TextRun({ text: `Telefone: ${COLLEM.telefone}`, font: "Arial", size: 18 })],
            alignment: AlignmentType.CENTER,
          }),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(documento);
  baixarBlob(blob, `${nomeArquivoSeguro(dados.nomeProjeto)}_Proposta_COLLEM.docx`);
};

export const gerarPropostaCollemPdf = ({ projeto, cliente, totalGeral }) => {
  const dados = dadosProposta({ projeto, cliente, totalGeral });
  const destinatario = [
    dados.nomeCliente,
    ...dados.enderecoCliente.split(/\r?\n/).filter(Boolean),
    dados.localCliente,
  ].filter(Boolean);
  const linhasDestinatario = destinatario
    .map((linha) => `<p class="destinatario">${escaparHtml(linha)}</p>`)
    .join("");
  const texto = (valor) => escaparHtml(valor).replace(/\r?\n/g, "<br/>");
  const html = `<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"/>
<title>${escaparHtml(nomeArquivoSeguro(dados.nomeProjeto))}_Proposta_COLLEM</title>
<style>
@page { size: A4; margin: 25mm 30mm; }
* { box-sizing: border-box; }
body { margin: 0; color: #111; font-family: "Arial Narrow", Arial, sans-serif; font-size: 12pt; }
.page { min-height: 247mm; page-break-after: always; display: flex; flex-direction: column; }
.page:last-child { page-break-after: auto; }
header { height: 22mm; display: flex; align-items: flex-start; }
header img { width: 30mm; height: auto; }
p { margin: 0 0 5mm; line-height: 1.5; text-align: justify; }
.data { text-align: left; margin-bottom: 11mm; }
.destinatario { text-align: right; margin: 0 0 2mm; line-height: 1.25; }
.referencia { margin-top: 9mm; }
h2 { margin: 6mm 0 3mm; font-size: 12pt; line-height: 1.3; }
.assinatura { margin-top: 16mm; text-align: center; font-family: Arial, sans-serif; font-size: 9pt; line-height: 1.7; }
footer { margin-top: auto; padding-top: 2mm; border-top: .6pt solid #111; text-align: center; font-family: Arial, sans-serif; font-size: 9pt; line-height: 1.4; }
footer a { color: #0563c1; text-decoration: underline; }
@media screen { body { background: #eee; padding: 20px; } .page { width: 210mm; min-height: 297mm; padding: 25mm 30mm; margin: 0 auto 20px; background: white; box-shadow: 0 3px 14px rgba(0,0,0,.14); } }
</style></head><body>
<section class="page">
<header><img src="${escaparHtml(collemLogo)}" alt="COLLEM Construtora"/></header>
<p class="data">Belo Horizonte, ${escaparHtml(dados.data)}</p>
${linhasDestinatario}
<p class="referencia"><strong>Ref.:</strong> Contrato de empreitada global de ${escaparHtml(dados.nomeProjeto.toLowerCase())}.</p>
<p><strong>Atte.:</strong> ${escaparHtml(dados.contato)};</p>
<p>Prezados senhores,</p>
<p>${texto(dados.introducao)}</p>
<h2>A) Preço</h2>
<p>O preço total para a execução dos trabalhos é de R$ ${formatarMoeda(dados.totalGeral)} (${escaparHtml(dados.totalExtenso)}) sendo especificados no ANEXO I;</p>
<h2>B) Condições de pagamento</h2>
<p>${texto(dados.condicoesPagamento)}</p>
<footer>${escaparHtml(COLLEM.rodapeEndereco)}<br/>e-mail: <a>${escaparHtml(COLLEM.email)}</a>&nbsp;&nbsp; site: <a>${escaparHtml(COLLEM.site)}</a></footer>
</section>
<section class="page">
<header><img src="${escaparHtml(collemLogo)}" alt="COLLEM Construtora"/></header>
<h2>C) Prazo de Execução</h2>
<p>Estimamos executar todos os trabalhos em ${texto(dados.prazo)}.</p>
<h2>D) Não inclusos</h2>
<p>${texto(dados.naoInclusos)}</p>
<h2>E) Condições especiais</h2>
<p>${texto(dados.condicoesEspeciais)}</p>
<p>Colocamo-nos à sua inteira disposição para qualquer esclarecimento.</p>
<p style="margin-top:16mm">Atenciosamente.</p>
<div class="assinatura">${escaparHtml(dados.responsavel)}<br/>Contratada: ${escaparHtml(COLLEM.nome)}<br/>CNPJ/MF: ${escaparHtml(COLLEM.cnpj)}<br/>Endereço: ${escaparHtml(COLLEM.endereco)}<br/>Telefone: ${escaparHtml(COLLEM.telefone)}</div>
<footer>${escaparHtml(COLLEM.rodapeEndereco)}<br/>e-mail: <a>${escaparHtml(COLLEM.email)}</a>&nbsp;&nbsp; site: <a>${escaparHtml(COLLEM.site)}</a></footer>
</section>
<script>window.onload=()=>{const imagens=Array.from(document.images||[]);Promise.all(imagens.map(img=>img.complete?Promise.resolve():new Promise(resolve=>{img.onload=resolve;img.onerror=resolve;}))).finally(()=>setTimeout(()=>window.print(),350));};</script>
</body></html>`;

  const janela = window.open("", "_blank");
  if (!janela) {
    alert("Não foi possível abrir a proposta. Verifique se o navegador bloqueou pop-ups.");
    return;
  }
  janela.document.open();
  janela.document.write(html);
  janela.document.close();
};
