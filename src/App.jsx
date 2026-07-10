import React, { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx-js-style";
import {
  Plus, Trash2, Pencil, X, Search, Upload, Download,
  ChevronDown, ChevronRight, Database, Calculator, Copy, Save, Percent, TrendingUp, RefreshCw,
  Tags, AlertTriangle, Check, FolderKanban, HardHat, User, LogIn, MapPin, Phone, Mail, Building2, FileText
} from "lucide-react";
import { FONTES_PADRAO, TIPOS, createDefaultProject, seedCpus } from "./data/defaultData";
import {
  applyCatalogToInsumos,
  buildCatalog,
  cpuValorUnit,
  findSubCpu,
  insumoValorUnitario,
  precoKey,
} from "./utils/calculos";
import { fmt, norm, num, uid } from "./utils/format";
import { loadGoogleDriveSnapshot, requestGoogleDriveAccess, saveGoogleDriveSnapshot } from "./services/googleDriveStore";
import {
  loadLocalSnapshot,
  loadOrcamentoData,
  saveLocalSnapshot,
} from "./services/orcamentoStore";

const BDI_PADRAO = {
  custoInicial: 0,
  admCentral: 0.04,
  contabilidade: 0.01,
  contingenciamento: 0.02,
  custoFinanceiro: 0.03,
  dasAnexoIV: 0.13,
  art: 0,
  lucro: 0.42,
};

const CLIENTE_PADRAO = {
  nome: "",
  local: "",
  contato: "",
  telefone: "",
  email: "",
  documento: "",
  endereco: "",
  observacoes: "",
};

const clienteDoProjeto = (projeto) => ({
  ...CLIENTE_PADRAO,
  ...(projeto?.clienteCadastro || {}),
  nome: projeto?.clienteCadastro?.nome || projeto?.cliente || "",
});

const clienteEstaCompleto = (cliente) =>
  Boolean(String(cliente?.nome || "").trim() && String(cliente?.local || "").trim());

const nomeArquivoSeguro = (valor) =>
  String(valor || "Orcamento")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const escapeHtml = (valor) =>
  String(valor ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

const itemVendaResumo = (item, bdiCalc, cpus, catalogMap) => {
  const quantidade = num(item.quantidade);
  let total = 0;

  (item.insumos || []).forEach((ins) => {
    const tipo = String(ins.tipo || "").toUpperCase().trim();
    const custoBase = num(ins.coeficiente) * quantidade * insumoValorUnitario(ins, cpus, catalogMap);
    const isMat =
      bdiCalc.faturamentoDireto &&
      (tipo === "MAT" ||
        tipo === "MATERIAL" ||
        (!tipo.includes("MO") && !tipo.includes("MÃO") && !tipo.includes("MAO") && !tipo.includes("EQUIP")));
    total += custoBase * (isMat ? bdiCalc.FatorBdiMateriais : bdiCalc.FatorBdi);
  });

  return {
    quantidade,
    total,
    unitario: quantidade > 0 ? total / quantidade : 0,
  };
};

const montarItensProposta = (etapas, bdiCalc, cpus, catalogMap) =>
  (etapas || []).map((etapa, idxEtapa) => {
    const itens = (etapa.itens || []).map((item, idxItem) => ({
      numero: `${idxEtapa + 1}.${idxItem + 1}`,
      descricao: item.servico || item.descricao || "",
      unidade: item.unidade || "",
      ...itemVendaResumo(item, bdiCalc, cpus, catalogMap),
    }));

    return {
      numero: String(idxEtapa + 1),
      nome: etapa.nome || `Etapa ${idxEtapa + 1}`,
      total: itens.reduce((s, item) => s + item.total, 0),
      itens,
    };
  });

const XLSX_MOEDA = '_-"R$"\\ * #,##0.00_-;\\-"R$"\\ * #,##0.00_-;_-"R$"\\ * "-"??_-;_-@';
const XLSX_NUMERO = "###,###,##0.00";

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

const criarAbaVendaModelo = (grupos, fatorVenda = 1) => {
  const rows = [[], [null, "PLANILHA DE MATERIAL"], [null, "ITEM", "DESCRIÇÃO DOS SERVIÇOS", "UNID.", "QUANT.", "VALOR UNIT.", "VALOR TOTAL", "TOTAL DO ITEM", fatorVenda, 250, 150]];
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
  const totalRow = rows.length + 1;
  rows.push([null, "TOTAL GERAL", null, null, null, null, null, { f: `SUM(H4:H${totalRow - 2})` }]);

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!merges"] = [
    { s: { r: 1, c: 1 }, e: { r: 1, c: 7 } },
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
  aplicarEstiloLinha(ws, 2, 2, 8, estiloVendaTitulo);
  aplicarEstiloLinha(ws, 3, 2, 8, estiloVendaCabecalho);
  aplicarFormatoNumerico(ws, 3, [8], XLSX_MOEDA);
  aplicarFormatoNumerico(ws, 3, [9], "0.00");

  groupRows.forEach(({ row }) => {
    aplicarEstiloLinha(ws, row, 2, 8, estiloVendaGrupo);
    aplicarFormatoNumerico(ws, row, [8], XLSX_MOEDA);
  });

  itemRows.forEach((row) => {
    aplicarEstiloLinha(ws, row, 2, 8, estiloVendaBase);
    aplicarFormatoNumerico(ws, row, [5, 6, 7], XLSX_NUMERO);
  });

  aplicarEstiloLinha(ws, totalRow, 2, 8, estiloVendaTotal);
  aplicarFormatoNumerico(ws, totalRow, [8], XLSX_MOEDA);

  return ws;
};

const exportarPropostaXlsx = ({ projeto, etapas, bdiCalc, cpus, catalogMap }) => {
  const grupos = montarItensProposta(etapas, bdiCalc, cpus, catalogMap);
  const wb = XLSX.utils.book_new();
  const wsValores = criarAbaVendaModelo(grupos, bdiCalc?.FatorBdi || 1);
  XLSX.utils.book_append_sheet(wb, wsValores, "VENDA");
  XLSX.writeFile(wb, `${nomeArquivoSeguro(projeto.nome)}_Proposta.xlsx`);
};

const gerarPropostaPdf = ({ projeto, cliente, etapas, bdiCalc, cpus, catalogMap }) => {
  const grupos = montarItensProposta(etapas, bdiCalc, cpus, catalogMap);
  const totalGeral = grupos.reduce((s, grupo) => s + grupo.total, 0);
  const hoje = new Date();
  const dataHoje = hoje.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit", year: "numeric" });
  const numeroProposta = `PROP - ${String(hoje.getMonth() + 1).padStart(2, "0")}/${String(hoje.getFullYear()).slice(-2)}`;
  const nomeProjeto = projeto?.nome || "Orçamento";
  const nomeCliente = cliente?.nome || "Cliente";
  const localObra = cliente?.local || cliente?.endereco || "";
  const contato = cliente?.contato || "";
  const observacoes = cliente?.observacoes || "";

  const linhasEscopo = grupos
    .map((grupo) => `
      <tr class="grupo">
        <td>${escapeHtml(grupo.numero)}.</td>
        <td>${escapeHtml(grupo.nome)}</td>
        <td></td>
        <td></td>
      </tr>
      ${grupo.itens.map((item) => `
        <tr>
          <td>${escapeHtml(item.numero)}</td>
          <td>${escapeHtml(item.descricao)}</td>
          <td>${escapeHtml(item.unidade)}</td>
          <td>${fmt(item.quantidade)}</td>
        </tr>
      `).join("")}
    `)
    .join("");

  const linhasValores = grupos
    .map((grupo) => `
      <tr class="grupo">
        <td>${escapeHtml(grupo.numero)}.</td>
        <td>${escapeHtml(grupo.nome)}</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td>R$ ${fmt(grupo.total)}</td>
      </tr>
      ${grupo.itens.map((item) => `
        <tr>
          <td>${escapeHtml(item.numero)}</td>
          <td>${escapeHtml(item.descricao)}</td>
          <td>${escapeHtml(item.unidade)}</td>
          <td>${fmt(item.quantidade)}</td>
          <td>R$ ${fmt(item.unitario)}</td>
          <td>R$ ${fmt(item.total)}</td>
          <td></td>
        </tr>
      `).join("")}
    `)
    .join("");

  const html = `
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <title>${escapeHtml(nomeArquivoSeguro(nomeProjeto))}_Proposta</title>
  <style>
    @page { size: A4; margin: 16mm 14mm 14mm; }
    * { box-sizing: border-box; }
    body { margin: 0; font-family: Arial, Helvetica, sans-serif; color: #111; font-size: 11px; }
    .page { min-height: 267mm; page-break-after: always; position: relative; padding-bottom: 18mm; }
    .page:last-child { page-break-after: auto; }
    header { display: grid; grid-template-columns: 1fr auto; gap: 16px; align-items: start; border-bottom: 1px solid #111; padding-bottom: 8px; margin-bottom: 18px; }
    .prop { font-weight: 700; font-size: 12px; margin-bottom: 4px; }
    .empresa { font-weight: 700; font-size: 13px; letter-spacing: .2px; }
    .dados { line-height: 1.35; }
    .pagina { text-align: right; line-height: 1.35; white-space: nowrap; }
    h1 { text-align: center; font-size: 15px; margin: 24px 0 20px; }
    h2 { font-size: 12px; margin: 16px 0 8px; }
    p { margin: 6px 0; line-height: 1.45; }
    table { width: 100%; border-collapse: collapse; table-layout: fixed; }
    th { background: #d8d8d8; font-weight: 700; text-align: left; }
    th, td { padding: 4px 5px; vertical-align: top; }
    .escopo th:nth-child(1), .escopo td:nth-child(1) { width: 9%; }
    .escopo th:nth-child(2), .escopo td:nth-child(2) { width: 67%; }
    .escopo th:nth-child(3), .escopo td:nth-child(3) { width: 10%; text-align: center; }
    .escopo th:nth-child(4), .escopo td:nth-child(4) { width: 14%; text-align: right; }
    .valores th:nth-child(1), .valores td:nth-child(1) { width: 8%; }
    .valores th:nth-child(2), .valores td:nth-child(2) { width: 42%; }
    .valores th:nth-child(3), .valores td:nth-child(3) { width: 8%; text-align: center; }
    .valores th:nth-child(4), .valores td:nth-child(4) { width: 10%; text-align: right; }
    .valores th:nth-child(5), .valores td:nth-child(5),
    .valores th:nth-child(6), .valores td:nth-child(6),
    .valores th:nth-child(7), .valores td:nth-child(7) { width: 11%; text-align: right; }
    .grupo td { background: #e2efd9; font-weight: 700; }
    .total td { background: #e2efd9; font-weight: 700; font-size: 12px; }
    .total td:first-child { text-align: center; }
    ul { margin: 6px 0 14px 18px; padding: 0; }
    li { margin: 5px 0; }
    .assinatura { margin-top: 48px; width: 260px; border-top: 1px solid #111; text-align: center; padding-top: 6px; }
    .footer { position: absolute; bottom: 0; left: 0; right: 0; font-size: 10px; color: #555; border-top: 1px solid #ddd; padding-top: 6px; display: flex; justify-content: space-between; }
    @media screen {
      body { background: #eee; padding: 20px; }
      .page { background: white; width: 210mm; margin: 0 auto 20px; padding: 16mm 14mm 14mm; box-shadow: 0 4px 16px rgba(0,0,0,.12); }
      .footer { left: 14mm; right: 14mm; bottom: 10mm; }
    }
  </style>
</head>
<body>
  <section class="page">
    <header>
      <div>
        <div class="prop">${escapeHtml(numeroProposta)}</div>
        <div class="empresa">ALPHA ENGENHARIA E SERVIÇOS</div>
        <div class="dados">Rua José Da Costa, 116 - São João Batista<br/>Belo Horizonte<br/>Telefone: 31 9 9203-1783</div>
      </div>
      <div class="pagina">Página 1 de 3</div>
    </header>
    <h1>PROPOSTA DE PRESTAÇÃO DE SERVIÇOS</h1>
    <p>Belo Horizonte, ${escapeHtml(dataHoje)}</p>
    <p>Aos cuidados de ${escapeHtml(nomeCliente)}${contato ? ` - ${escapeHtml(contato)}` : ""}.</p>
    <p><strong>Ref.</strong> ${escapeHtml(nomeProjeto)}</p>
    <p><strong>Endereço da Obra:</strong> ${escapeHtml(localObra)}</p>
    <h2>Escopo do Serviço:</h2>
    <table class="escopo">
      <thead><tr><th>ITEM</th><th>DESCRIÇÃO DOS SERVIÇOS</th><th>UNID.</th><th>QUANT.</th></tr></thead>
      <tbody>${linhasEscopo}</tbody>
    </table>
    <div class="footer"><span>ALPHA ENGENHARIA E SERVIÇOS</span><span>${escapeHtml(numeroProposta)}</span></div>
  </section>

  <section class="page">
    <header>
      <div>
        <div class="prop">${escapeHtml(numeroProposta)}</div>
        <div class="empresa">ALPHA ENGENHARIA E SERVIÇOS</div>
        <div class="dados">Rua José Da Costa, 116 - São João Batista<br/>Belo Horizonte<br/>Telefone: 31 9 9203-1783</div>
      </div>
      <div class="pagina">Página 2 de 3</div>
    </header>
    <table class="valores">
      <thead><tr><th>ITEM</th><th>DESCRIÇÃO DOS SERVIÇOS</th><th>UNID.</th><th>QUANT.</th><th>VALOR UNIT.</th><th>VALOR TOTAL</th><th>TOTAL DO ITEM</th></tr></thead>
      <tbody>
        ${linhasValores}
        <tr class="total"><td colspan="6">TOTAL GERAL</td><td>R$ ${fmt(totalGeral)}</td></tr>
      </tbody>
    </table>
    <h2>PLANILHA DE MATERIAL</h2>
    <h2>Responsabilidade da ALPHA ENGENHARIA:</h2>
    <ul>
      <li>Acompanhamento Técnico;</li>
      <li>Fornecimento de EPIs para execução das atividades;</li>
      <li>Fornecimento de mão de obra;</li>
      <li>Fornecimento de equipamentos;</li>
      <li>Fornecimento de almoço e transporte para funcionários;</li>
      <li>Fornecimento de material conforme composição do orçamento.</li>
    </ul>
    <h2>Responsabilidade do Cliente:</h2>
    <ul>
      <li>Fornecimento de acesso ao local de prestação de serviço;</li>
      <li>Permitir os funcionários a usarem as instalações sanitárias;</li>
      <li>Fornecimento de água potável.</li>
    </ul>
    <div class="footer"><span>ALPHA ENGENHARIA E SERVIÇOS</span><span>${escapeHtml(numeroProposta)}</span></div>
  </section>

  <section class="page">
    <header>
      <div>
        <div class="prop">${escapeHtml(numeroProposta)}</div>
        <div class="empresa">ALPHA ENGENHARIA E SERVIÇOS</div>
        <div class="dados">Rua José Da Costa, 116 - São João Batista<br/>Belo Horizonte<br/>Telefone: 31 9 9203-1783</div>
      </div>
      <div class="pagina">Página 3 de 3</div>
    </header>
    <h2>Condições de pagamento:</h2>
    <p>Entrada de 40% (R$ ${fmt(totalGeral * 0.4)}) e o restante (R$ ${fmt(totalGeral * 0.6)}) conforme avanço dos serviços em medições.</p>
    <p>Pagamento via PIX (52.903.822/0001-86) 5 dias após a emissão da NF.</p>
    <h2>Prazo para Execução:</h2>
    <ul><li>A definir conforme cronograma aprovado entre as partes.</li></ul>
    ${observacoes ? `<h2>Observações:</h2><p>${escapeHtml(observacoes).replace(/\n/g, "<br/>")}</p>` : ""}
    <div class="assinatura">ALPHA ENGENHARIA E SERVIÇOS</div>
    <div class="footer"><span>ALPHA ENGENHARIA E SERVIÇOS</span><span>${escapeHtml(numeroProposta)}</span></div>
  </section>
  <script>
    window.onload = () => {
      setTimeout(() => window.print(), 350);
    };
  </script>
</body>
</html>`;

  const printWindow = window.open("", "_blank", "noopener,noreferrer");
  if (!printWindow) {
    alert("Não foi possível abrir a janela da proposta. Verifique se o navegador bloqueou pop-ups.");
    return;
  }
  printWindow.document.open();
  printWindow.document.write(html);
  printWindow.document.close();
};

const calcularPrecoVendaProjeto = (etapas, bdi, cpus, catalogMap) => {
  const calcularFatorBdiQualquer = (t = {}) => {
    const ac = num(t.admCentral || t.adminCentral);
    const c = num(t.contabilidade);
    const co = num(t.contingenciamento);
    const cf = num(t.custoFinanceiro);
    const l = num(t.lucro);
    const das = num(t.dasAnexoIV || 0);
    const art = num(t.art);
    const pv = das + art;
    const numerador = (1 + ac) * (1 + c) * (1 + co) * (1 + cf) * (1 + l);
    const denominador = 1 - pv;
    return denominador <= 0 ? 1 : numerador / denominador;
  };

  const FatorBdiGeral = calcularFatorBdiQualquer(bdi || BDI_PADRAO);
  const faturamentoDireto = !!bdi?.faturamentoDireto;
  const FatorBdiMateriais =
    faturamentoDireto && bdi?.materiais
      ? calcularFatorBdiQualquer(bdi.materiais)
      : FatorBdiGeral;

  let totalCustoDireto = 0;
  let totalPrecoVenda = 0;

  (etapas || []).forEach((e) => {
    (e.itens || []).forEach((it) => {
      const qtdCpu = num(it.quantidade);
      (it.insumos || []).forEach((ins) => {
        const tipo = String(ins.tipo || "").toUpperCase().trim();
        const custoInsumoTotal = num(ins.coeficiente) * qtdCpu * insumoValorUnitario(ins, cpus, catalogMap);
        totalCustoDireto += custoInsumoTotal;

        if (faturamentoDireto && (tipo === "MAT" || tipo === "MATERIAL" || (!tipo.includes("MO") && !tipo.includes("MÃO") && !tipo.includes("MAO") && !tipo.includes("EQUIP")))) {
          totalPrecoVenda += custoInsumoTotal * FatorBdiMateriais;
        } else {
          totalPrecoVenda += custoInsumoTotal * FatorBdiGeral;
        }
      });
    });
  });

  const totalDiValor = Math.max(0, totalPrecoVenda - totalCustoDireto);
  const totalDiRate = totalCustoDireto > 0 ? totalDiValor / totalCustoDireto : 0;

  return {
    bdiRate: FatorBdiGeral - 1,
    bdiRateMateriais: FatorBdiMateriais - 1,
    FatorBdi: FatorBdiGeral,
    FatorBdiMateriais,
    faturamentoDireto,
    totalDiValor,
    totalDiRate,
    valorVenda: totalPrecoVenda,
  };
};

export default function App() {
  const [tab, setTab] = useState("projetos");
  const [cpus, setCpusState] = useState([]);
  const [projetos, setProjetos] = useState([]);
  const [projetoAtivoId, setProjetoAtivoId] = useState("");
  const [precos, setPrecos] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");
  const [driveConnected, setDriveConnected] = useState(false);
  const fileInputRef = useRef(null);
  const cpuHashesRef = useRef({});
  const [cpusDirty, setCpusDirty] = useState(false);
  // Novos estados para controle de recolhimento/expansão das camadas
  const [etapasExpandidas, setEtapasExpandidas] = useState({});
  const [cpusExpandidas, setCpusExpandidas] = useState({});

  const setCpus = (nextCpus) => {
    setCpusDirty(true);
    setCpusState(nextCpus);
  };

  const aplicarDadosCarregados = (data) => {
    if (data.empty) {
      const defaultProj = createDefaultProject();
      setCpusState(seedCpus());
      setCpusDirty(true);
      setProjetos([defaultProj]);
      setPrecos([]);
      setProjetoAtivoId(defaultProj.id);
      setStatus("Nenhum dado salvo no Firebase. Projeto inicial criado localmente.");
      return;
    }

    setCpusState(data.cpus || []);
    cpuHashesRef.current = data.cpuHashes || {};
    setCpusDirty(false);
    setProjetos(data.projetos || []);
    setPrecos(data.precos || []);
    setProjetoAtivoId(data.projetoAtivoId || "");
    setStatus("Dados carregados do Firebase.");
  };

  const carregarDados = async ({ usarDrive = true } = {}) => {
    setBusy(true);
    setStatus("Carregando...");
    try {
      const driveData = usarDrive ? await loadGoogleDriveSnapshot() : null;
      if (usarDrive && driveData) {
        aplicarDadosCarregados(driveData);
        await saveLocalSnapshot(driveData);
        setDriveConnected(true);
        setStatus("Dados carregados do Google Drive.");
      } else {
        const data = await loadOrcamentoData();
        aplicarDadosCarregados(data);
      }
    } catch (e) {
      console.error("Erro ao carregar Firestore:", e);
      try {
        const local = await loadLocalSnapshot();
        if (local) {
          aplicarDadosCarregados(local);
          setStatus("Firebase indisponivel. Dados carregados do backup local deste navegador.");
        } else {
          setStatus("Falha ao carregar: " + (e?.message || e));
        }
      } catch (localError) {
        console.error("Erro ao carregar backup local:", localError);
        setStatus("Falha ao carregar: " + (e?.message || e));
      }
    } finally {
      setLoaded(true);
      setBusy(false);
      setTimeout(() => setStatus(""), 5000);
    }
  };

  const salvarDados = async () => {
    setBusy(true);
    setStatus("Salvando backup local...");
    try {
      await saveLocalSnapshot({ cpus, projetos, precos, projetoAtivoId });
      setStatus("Backup local salvo. Sincronizando Google Drive...");
      await saveGoogleDriveSnapshot({
        cpus,
        projetos,
        precos,
        projetoAtivoId,
      });
      setDriveConnected(true);
      setCpusDirty(false);
      setStatus("Salvo localmente e no Google Drive.");
    } catch (e) {
      console.error("Erro ao salvar no Google Drive:", e);
      setStatus("Salvo localmente. Falha ao sincronizar Google Drive: " + (e?.message || e));
    } finally {
      setBusy(false);
      setTimeout(() => setStatus(""), 12000);
    }
  };

  const conectarGoogleDrive = async () => {
    setBusy(true);
    setStatus("Conectando Google Drive...");
    try {
      await requestGoogleDriveAccess();
      setDriveConnected(true);
      setStatus("Google Drive conectado.");
    } catch (e) {
      setStatus("Falha ao conectar Drive: " + (e?.message || e));
    } finally {
      setBusy(false);
      setTimeout(() => setStatus(""), 8000);
    }
  };

  // Carrega uma vez ao abrir. Depois disso, salvar/carregar sao acoes manuais.
  useEffect(() => {
    carregarDados({ usarDrive: false });
  }, []);
  
  // Projeto Corrente Detectado
  const projetoAtivo = useMemo(() => {
    return projetos.find((p) => p.id === projetoAtivoId) || projetos[0] || null;
  }, [projetos, projetoAtivoId]);

  const clienteAtivo = useMemo(() => clienteDoProjeto(projetoAtivo), [projetoAtivo]);
  const cadastroClienteOk = useMemo(() => clienteEstaCompleto(clienteAtivo), [clienteAtivo]);
  const etapas = useMemo(() => projetoAtivo?.etapas || [], [projetoAtivo]);
  const bdi = useMemo(() => projetoAtivo?.bdi || BDI_PADRAO, [projetoAtivo]);

  const setEtapas = (novasEtapas) => {
    if (!projetoAtivoId) return;
    setProjetos((prev) =>
      prev.map((p) => (p.id === projetoAtivoId ? { ...p, etapas: typeof novasEtapas === "function" ? novasEtapas(p.etapas) : novasEtapas } : p))
    );
  };

  const setBdi = (novoBdi) => {
    if (!projetoAtivoId) return;
    setProjetos((prev) =>
      prev.map((p) => (p.id === projetoAtivoId ? { ...p, bdi: typeof novoBdi === "function" ? novoBdi(p.bdi) : novoBdi } : p))
    );
  };

  const setClienteAtivo = (novoCliente) => {
    if (!projetoAtivoId) return;
    setProjetos((prev) =>
      prev.map((p) => {
        if (p.id !== projetoAtivoId) return p;
        const clienteAtual = clienteDoProjeto(p);
        const clienteCadastro = typeof novoCliente === "function" ? novoCliente(clienteAtual) : novoCliente;
        return {
          ...p,
          cliente: clienteCadastro.nome || "",
          clienteCadastro,
        };
      })
    );
  };

  const abrirAbaProjeto = (proximaTab) => {
    if (proximaTab !== "cliente" && !cadastroClienteOk) {
      setTab("cliente");
      setStatus("Preencha Nome do cliente e Local da obra para continuar.");
      setTimeout(() => setStatus(""), 5000);
      return;
    }
    setTab(proximaTab);
  };

  const catalog = useMemo(() => buildCatalog(cpus, projetos, projetoAtivoId, precos), [cpus, projetos, projetoAtivoId, precos]);
  const catalogMap = useMemo(() => new Map(catalog.map((c) => [c.key, c])), [catalog]);

  const upsertPreco = (descricao, tipo, unidade, valorUnitario) => {
    const key = precoKey(descricao);
    if (!key) return;
    setPrecos((prev) => {
      const idx = prev.findIndex((p) => precoKey(p.descricao) === key);
      if (idx >= 0) {
        const next = [...prev];
        next[idx] = { ...next[idx], descricao, tipo, unidade, valorUnitario };
        return next;
      }
      return [...prev, { id: uid(), descricao, tipo, unidade, valorUnitario }];
    });
  };

  const removePreco = (descricao) => {
    const key = precoKey(descricao);
    setPrecos((prev) => prev.filter((p) => precoKey(p.descricao) !== key));
  };

  // Melhoria Crítica solicitada: Altera APENAS os insumos associados Ã  aba CUSTOS do projeto ativo
  const aplicarPrecoNoOrcamentoAtivo = (descricao, valorUnitario) => {
    const key = precoKey(descricao);
    if (!projetoAtivoId) return;
    setProjetos((prev) =>
      prev.map((p) => {
        if (p.id !== projetoAtivoId) return p;
        return {
          ...p,
          etapas: p.etapas.map((e) => ({
            ...e,
            itens: e.itens.map((it) => ({
              ...it,
              insumos: it.insumos.map((i) => (precoKey(i.descricao) === key ? { ...i, valorUnitario } : i))
            }))
          }))
        };
      })
    );
  };

  const aplicarTodosPrecosNoOrcamentoAtivo = () => {
    if (!projetoAtivoId) return;
    setProjetos((prev) =>
      prev.map((p) => {
        if (p.id !== projetoAtivoId) return p;
        return {
          ...p,
          etapas: p.etapas.map((e) => ({
            ...e,
            itens: e.itens.map((it) => ({
              ...it,
              insumos: applyCatalogToInsumos(it.insumos, catalogMap)
            }))
          }))
        };
      })
    );
  };

// NOVO: Função para varrer e consolidar o quantitativo de materiais
  const processarMateriais = useMemo(() => {
    const resumoMAT = {};
    etapas.forEach((etapa) => {
      (etapa.itens || []).forEach((item) => {
        const qtdItem = num(item.quantidade);
        (item.insumos || []).forEach((insumo) => {
          // Filtra o que for do tipo "MAT" ou o que NÃƒO for Mão de Obra (MO) ou Equipamento (EQUIP)
          if (insumo.tipo === "MAT" || (insumo.tipo !== "MO" && insumo.tipo !== "EQUIP" && insumo.unidade?.toLowerCase() !== "h")) {
            const nomeMat = (insumo.descricao || "").toUpperCase().trim();
            if (!nomeMat) return;

            // Busca o valor unitário atualizado diretamente do catálogo/banco de preços de referência
            const entry = catalogMap.get(precoKey(insumo.descricao));
            const precoUnit = entry && entry.valorUnitario !== "" ? num(entry.valorUnitario) : num(insumo.valorUnitario);

            const qtdTotal = num(insumo.coeficiente) * qtdItem;
            const custoTotal = qtdTotal * precoUnit;

            if (!resumoMAT[nomeMat]) {
              resumoMAT[nomeMat] = {
                material: insumo.descricao,
                unidade: insumo.unidade || "un",
                quantidade: 0,
                valorUnitario: precoUnit,
                valorTotal: 0,
              };
            }
            resumoMAT[nomeMat].quantidade += qtdTotal;
            resumoMAT[nomeMat].valorTotal += custoTotal;
          }
        });
      });
    });
    return Object.values(resumoMAT).sort((a, b) => b.valorTotal - a.valorTotal); // Ordena do mais caro para o mais barato
  }, [etapas, catalogMap]);

  const grandTotal = useMemo(() => {
    return etapas.reduce(
      (s, e) => s + e.itens.reduce((s2, it) => s2 + num(it.quantidade) * cpuValorUnit(it.insumos, cpus, catalogMap), 0),
      0
    );
  }, [etapas, cpus, catalogMap]);

  const bdiCalc = useMemo(() => {
    return calcularPrecoVendaProjeto(etapas, bdi, cpus, catalogMap);
  }, [bdi, etapas, cpus, catalogMap]);

  // Abas disponíveis apenas dentro de um projeto ativo
  const abasProjeto = ["cliente", "custo", "planilha", "bdi", "precovenda", "maoobra", "materiais", "precos"];
  const tabEhDeProjeto = abasProjeto.includes(tab);

  return (
    <div className="min-h-screen bg-stone-50 text-stone-900">

      <div className="max-w-[1500px] mx-auto px-4 py-6 lg:flex lg:items-start lg:gap-5">
        <aside className="lg:sticky lg:top-4 lg:w-64 lg:shrink-0 mb-5 lg:mb-0">
          <div className="bg-white border border-stone-200 shadow-sm rounded-lg overflow-hidden">
            <div className="px-4 py-3 border-b border-stone-200">
              <p className="text-[10px] font-semibold text-stone-400 uppercase">Navegação</p>
              <p className="text-xs font-medium text-stone-700 truncate mt-0.5">
                {projetoAtivo?.nome || "Sem orçamento ativo"}
              </p>
            </div>

            <nav className="p-2 space-y-1 max-lg:flex max-lg:overflow-x-auto max-lg:space-y-0 max-lg:gap-1">
              <SideTabBtn active={tab === "projetos"} onClick={() => setTab("projetos")} icon={<FolderKanban size={15} />}>
                Orçamentos ({projetos.length})
              </SideTabBtn>
              <SideTabBtn active={tab === "cpus"} onClick={() => setTab("cpus")} icon={<Database size={15} />}>
                Base de CPUs ({cpus.length})
              </SideTabBtn>

              {projetoAtivo && (
                <>
                  <div className="max-lg:hidden border-t border-stone-200 my-2 pt-2">
                    <p className="px-2 text-[10px] font-semibold text-stone-400 uppercase truncate">
                      {projetoAtivo.nome}
                    </p>
                  </div>
                  <SideTabBtn active={tab === "cliente"} onClick={() => setTab("cliente")} icon={<User size={15} />}>
                    Cadastro Cliente
                    {!cadastroClienteOk && (
                      <span className="ml-auto w-1.5 h-1.5 rounded-full bg-amber-500 inline-block" />
                    )}
                  </SideTabBtn>
                  <SideTabBtn active={tab === "custo"} onClick={() => abrirAbaProjeto("custo")} icon={<Calculator size={15} />}>
                    Lançamento CPU
                  </SideTabBtn>
                  <SideTabBtn active={tab === "planilha"} onClick={() => abrirAbaProjeto("planilha")} icon={<FolderKanban size={15} />}>
                    Planilha de custo
                  </SideTabBtn>
                  <SideTabBtn active={tab === "bdi"} onClick={() => abrirAbaProjeto("bdi")} icon={<Percent size={15} />}>
                    BDI - {fmt(bdiCalc.bdiRate * 100)}%
                  </SideTabBtn>
                  <SideTabBtn active={tab === "precovenda"} onClick={() => abrirAbaProjeto("precovenda")} icon={<TrendingUp size={15} />}>
                    Venda - R$ {fmt(bdiCalc.valorVenda)}
                  </SideTabBtn>
                  <SideTabBtn active={tab === "maoobra"} onClick={() => abrirAbaProjeto("maoobra")} icon={<HardHat size={15} />}>
                    Mão de Obra
                  </SideTabBtn>
                  <SideTabBtn active={tab === "materiais"} onClick={() => abrirAbaProjeto("materiais")} icon={<Database size={15} />}>
                    Materiais
                  </SideTabBtn>
                  <SideTabBtn active={tab === "precos"} onClick={() => abrirAbaProjeto("precos")} icon={<Tags size={15} />}>
                    Banco de Preços ({catalog.length})
                    {catalog.some((c) => c.divergente) && (
                      <span className="ml-auto w-1.5 h-1.5 rounded-full bg-amber-500 inline-block" />
                    )}
                  </SideTabBtn>
                </>
              )}
            </nav>
          </div>
        </aside>

        <main className="min-w-0 flex-1">
        {/* HEADER */}
        <header className="flex items-center justify-between mb-6 gap-4">
          <div className="min-w-0">
            <h1 className="text-2xl font-semibold tracking-tight">Orçamentador por CPU</h1>
            <p className="text-sm text-stone-500 truncate">
              {projetoAtivo
                ? `Orçamento: ${projetoAtivo.nome}  -  ${clienteAtivo.nome || "Cliente não cadastrado"}`
                : "Crie ou selecione um orçamento para começar"}
            </p>
          </div>
          <div className="flex items-center gap-2 flex-wrap justify-end">
            <span className="text-xs text-stone-400 min-h-4">{status}</span>
            <button
              type="button"
              onClick={conectarGoogleDrive}
              disabled={busy}
              className={`flex items-center gap-1.5 px-3 py-1.5 text-xs border rounded-lg font-medium disabled:opacity-50 disabled:cursor-not-allowed ${
                driveConnected
                  ? "border-emerald-200 bg-emerald-50 text-emerald-700"
                  : "border-stone-300 bg-white hover:bg-stone-50 text-stone-700"
              }`}
              title="Conectar sua conta Google para salvar no Drive"
            >
              <LogIn size={13} /> {driveConnected ? "Drive conectado" : "Conectar Drive"}
            </button>
            <button
              type="button"
              onClick={() => carregarDados({ usarDrive: true })}
              disabled={busy}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs border border-stone-300 rounded-lg font-medium bg-white hover:bg-stone-50 text-stone-700 disabled:opacity-50 disabled:cursor-not-allowed"
              title="Carregar dados salvos no Google Drive"
            >
              <RefreshCw size={13} className={busy ? "animate-spin" : ""} /> Carregar
            </button>
            <button
              type="button"
              onClick={salvarDados}
              disabled={busy || !loaded}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs border border-stone-900 rounded-lg font-medium bg-stone-900 hover:bg-stone-800 text-white disabled:opacity-50 disabled:cursor-not-allowed"
              title="Salvar dados atuais no Google Drive"
            >
              <Save size={13} /> Salvar
            </button>
          </div>
        </header>

        <datalist id="insumos-catalogo">
          {catalog.map((c) => <option key={c.key} value={c.descricao} />)}
        </datalist>

        {/* â”€â”€ CONTEÃšDO DAS ABAS â”€â”€ */}
        {tab === "projetos" && (
          <div className="space-y-4">
            <div className="flex justify-between items-center bg-white border border-stone-200 rounded-lg p-4 shadow-xs">
              <div>
                <h2 className="text-base font-semibold text-stone-800">Seus Orçamentos</h2>
                <p className="text-xs text-stone-500">Gerencie, selecione ou crie novas pastas de projetos e fechamentos comerciais.</p>
              </div>
              <button
                onClick={() => {
                  const pId = uid();
                  setProjetos((prev) => [
                    ...prev,
                    {
                      id: pId,
                      nome: `Novo Orçamento - ${prev.length + 1}`,
                      cliente: "",
                      clienteCadastro: { ...CLIENTE_PADRAO },
                      etapas: [{ id: uid(), nome: "Etapa Inicial", itens: [] }],
                      bdi: {
                        custoInicial: 0,
                        admCentral: 0,
                        contabilidade: 0,
                        contingenciamento: 0,
                        custoFinanceiro: 0,
                        dasAnexoIV: 0,
                        art: 0,
                        lucro: 0,
                        faturamentoDireto: false,
                        materiais: { admCentral: 0, contabilidade: 0, contingenciamento: 0, custoFinanceiro: 0, lucro: 0, dasAnexoIV: 0, art: 0 }
                      }
                    }
                  ]);
                  setProjetoAtivoId(pId);
                  setTab("cliente");
                }}
                className="flex items-center gap-1.5 px-3 py-1.5 bg-stone-900 text-white rounded-lg text-xs font-medium hover:bg-stone-800"
              >
                <Plus size={14} /> Novo Orçamento
              </button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {projetos.map((p) => {
                const isActive = p.id === projetoAtivoId;
                
                // Calcula o custo direto acumulado deste projeto específico
                const cDiretoTotal = (p.etapas || []).reduce(
                  (s, e) => s + (e.itens || []).reduce((s2, it) => s2 + num(it.quantidade) * cpuValorUnit(it.insumos, cpus, catalogMap), 0),
                  0
                );

                const valorVendaCalculado = calcularPrecoVendaProjeto(p.etapas || [], p.bdi || BDI_PADRAO, cpus, catalogMap).valorVenda;

                return (
                  (() => {
                    const clienteCard = clienteDoProjeto(p);
                    const clienteOk = clienteEstaCompleto(clienteCard);
                    return (
                  <div
                    key={p.id}
                    className={`bg-white border rounded-xl p-4 shadow-xs space-y-3 cursor-pointer transition-all ${
                      isActive ? "border-stone-900 ring-1 ring-stone-900 bg-stone-50/20" : "border-stone-200 hover:border-stone-400"
                    }`}
                    onClick={() => { setProjetoAtivoId(p.id); setTab(clienteOk ? "custo" : "cliente"); }}
                  >
                    <div className="flex justify-between items-start">
                      <div>
                        <h3 className="font-semibold text-stone-800 text-sm flex items-center gap-1.5 uppercase">
                          <input
                            value={p.nome}
                            onClick={(e) => e.stopPropagation()}
                            onChange={(e) => setProjetos(prev => prev.map(x => x.id === p.id ? { ...x, nome: e.target.value } : x))}
                            className="bg-transparent border-b border-transparent hover:border-stone-300 focus:border-stone-600 outline-none w-full uppercase font-semibold text-stone-800 text-sm"
                          />
                        </h3>
                        <input
                          value={clienteCard.nome || ""}
                          onClick={(e) => e.stopPropagation()}
                          onChange={(e) => setProjetos(prev => prev.map(x => {
                            if (x.id !== p.id) return x;
                            const cadastro = { ...clienteDoProjeto(x), nome: e.target.value };
                            return { ...x, cliente: cadastro.nome, clienteCadastro: cadastro };
                          }))}
                          placeholder="Nome do cliente"
                          className="text-xs text-stone-400 bg-transparent border-b border-transparent hover:border-stone-300 focus:border-stone-500 outline-none w-full mt-0.5"
                        />
                        <div className="text-[11px] text-stone-400 mt-1 flex items-center gap-1">
                          <MapPin size={11} />
                          <span>{clienteCard.local || "Local da obra pendente"}</span>
                        </div>
                      </div>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          if (projetos.length <= 1) return alert("Não é possível apagar todos os orçamentos.");
                          if (confirm(`Tem certeza que deseja apagar o orçamento "${p.nome}"?`)) {
                            setProjetos((prev) => prev.filter((item) => item.id !== p.id));
                            if (isActive) setProjetoAtivoId(projetos.find((item) => item.id !== p.id)?.id || "");
                          }
                        }}
                        className="text-stone-400 hover:text-red-600 p-1 rounded-md transition-colors"
                        title="Excluir Orçamento"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>

                    <div className="grid grid-cols-2 gap-2 pt-2 border-t border-stone-100 text-xs font-mono">
                      <div>
                        <span className="text-stone-400 block font-sans text-[10px] uppercase">Custo Direto:</span>
                        <span className="text-stone-600 font-medium">R$ {fmt(cDiretoTotal)}</span>
                      </div>
                      <div className="text-right">
                        <span className="text-stone-400 block font-sans text-[10px] uppercase">Preço Estimado (Venda):</span>
                        <span className="text-stone-900 font-bold text-sm">R$ {fmt(valorVendaCalculado)}</span>
                      </div>
                    </div>

                    <div className="flex justify-between items-center pt-1 text-[11px]">
                      <span className="text-stone-400">{(p.etapas || []).length} etapa(s) cadastrada(s)</span>
                      {!clienteOk && (
                        <span className="text-amber-700 font-semibold bg-amber-50 border border-amber-200 px-2 py-0.5 rounded text-[10px] uppercase tracking-wider">
                          Cadastro pendente
                        </span>
                      )}
                      {isActive && (
                        <span className="text-stone-800 font-semibold bg-stone-200 px-2 py-0.5 rounded text-[10px] uppercase tracking-wider">
                          Selecionado ativo
                        </span>
                      )}
                    </div>
                  </div>
                    );
                  })()
                );
              })}
            </div>
          </div>
        )}
        {tab === "cpus" && (
          <CpuLibrary cpus={cpus} setCpus={setCpus} fileInputRef={fileInputRef} catalogMap={catalogMap} />
        )}

        {/* Abas de projeto - só renderizam se houver projeto ativo */}
        {tabEhDeProjeto && !projetoAtivo && (
          <div className="text-center py-20 text-stone-400">
            <FolderKanban size={40} className="mx-auto mb-3 opacity-30" />
            <p className="text-sm font-medium">Nenhum orçamento selecionado.</p>
            <button onClick={() => setTab("projetos")} className="mt-3 text-xs underline">
              Criar ou selecionar um orçamento
            </button>
          </div>
        )}
        {tab === "cliente" && projetoAtivo && (
          <CadastroCliente
            projeto={projetoAtivo}
            cliente={clienteAtivo}
            setProjetos={setProjetos}
            setCliente={setClienteAtivo}
            completo={cadastroClienteOk}
            onContinuar={() => abrirAbaProjeto("custo")}
          />
        )}
        {tab === "custo" && projetoAtivo && (
          <Orcamento 
            etapas={etapas} 
            setEtapas={setEtapas} 
            cpus={cpus} 
            grandTotal={grandTotal} 
            catalogMap={catalogMap} 
            onUpsertPreco={upsertPreco}
          />
        )}
{tab === "bdi" && projetoAtivo && (
          <BdiTab bdi={bdi} setBdi={setBdi} bdiCalc={bdiCalc} grandTotal={grandTotal} />
        )}

        {tab === "planilha" && projetoAtivo && (
          <div className="bg-white border border-stone-200 shadow-sm rounded-lg overflow-hidden p-5 space-y-4">
            <div className="flex justify-between items-center">
              <div>
                <h2 className="text-base font-semibold text-stone-800">Planilha de Exploração de Custos Diretos</h2>
                <p className="text-xs text-stone-500">Visualização hierárquica completa: Etapa / CPU / Insumos associados.</p>
              </div>
              <div className="flex gap-2">
                <button 
                  onClick={() => {
                  const objEtapas = {};
                  const objCpus = {};
                  etapas.forEach((etapa) => {
                    // Usa o ID real da etapa ou o índice fallback
                    const eId = etapa.id || `etapa-${etapas.indexOf(etapa)}`;
                    objEtapas[eId] = true;
                    (etapa.itens || []).forEach((item) => {
                      // Usa o ID real do item
                      objCpus[item.id] = true;
                    });
                  });
                  setEtapasExpandidas(objEtapas);
                  setCpusExpandidas(objCpus);
                }}
                  className="px-2 py-1 text-[11px] font-medium border border-stone-200 rounded hover:bg-stone-50 text-stone-600 flex items-center gap-1"
                >
                  Expandir Tudo
                </button>
                <button 
                  onClick={() => { setEtapasExpandidas({}); setCpusExpandidas({}); }}
                  className="px-2 py-1 text-[11px] font-medium border border-stone-200 rounded hover:bg-stone-50 text-stone-600"
                >
                  Recolher Tudo
                </button>

                {/* EXCEL DA PLANILHA DE CUSTO */}
                <button 
                  onClick={() => {
                    const data = [];
                    data.push(["ESTRUTURA", "DESCRIÇÃO", "UND", "QTD PROP.", "CUSTO UNIT", "CUSTO TOTAL"]);
                    etapas.forEach((etapa, idxE) => {
                      data.push([`${idxE + 1}`, etapa.nome, "", "", "", (etapa.itens || []).reduce((acc, it) => acc + (num(it.quantidade) * cpuValorUnit(it.insumos, cpus, catalogMap)), 0)]);
                      (etapa.itens || []).forEach((item, idxI) => {
                        const numCpu = `${idxE + 1}.${idxI + 1}`;
                        data.push([numCpu, item.servico || item.descricao, item.unidade, num(item.quantidade), cpuValorUnit(item.insumos, cpus, catalogMap), num(item.quantidade) * cpuValorUnit(item.insumos, cpus, catalogMap)]);
                        (item.insumos || []).forEach((ins, idxIn) => {
                          const pUnit = insumoValorUnitario(ins, cpus, catalogMap);
                          data.push([`${numCpu}.${idxIn + 1}`, `[${ins.tipo}] ${ins.descricao}`, ins.unidade || "un", num(ins.coeficiente) * num(item.quantidade), pUnit, (num(ins.coeficiente) * num(item.quantidade)) * pUnit]);
                        });
                      });
                    });
                    const ws = XLSX.utils.aoa_to_sheet(data);
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, "Planilha de Custo");
                    XLSX.writeFile(wb, `${projetoAtivo.nome || "Orcamento"}_Planilha_Custo.xlsx`);
                  }}
                  className="px-2 py-1 text-[11px] font-medium border border-emerald-200 text-emerald-700 bg-emerald-50/50 rounded hover:bg-emerald-50 flex items-center gap-1"
                >
                  <Download size={12} /> Excel (.xlsx)
                </button>

                {/* PDF LIMPO DA PLANILHA DE CUSTO */}
                <button 
                  onClick={() => {
                    const tituloOriginal = document.title;
                    document.title = `${projetoAtivo.nome || "Orcamento"}_Planilha_Custo`;
                    const estiloPrint = document.createElement("style");
                    estiloPrint.innerHTML = `
                      @media print {
                        body * { visibility: hidden; }
                        #area-planilha-custo, #area-planilha-custo * { visibility: visible; }
                        #area-planilha-custo { position: absolute; left: 0; top: 0; width: 100%; background: white !important; }
                      }
                    `;
                    document.head.appendChild(estiloPrint);
                    window.print();
                    document.head.removeChild(estiloPrint);
                    document.title = tituloOriginal;
                  }}
                  className="px-2 py-1 text-[11px] font-medium border border-red-200 text-red-700 bg-red-50/50 rounded hover:bg-red-50 flex items-center gap-1"
                >
                  <Download size={12} /> PDF (.pdf)
                </button>
              </div>
            </div>

            <div id="area-planilha-custo" className="border border-stone-200 rounded-lg overflow-hidden bg-white">
              <div className="grid grid-cols-12 gap-2 px-4 py-2.5 bg-stone-100 border-b border-stone-200 text-stone-500 font-semibold text-[11px] uppercase tracking-wider">
                <span className="col-span-6">Estrutura (Etapa / CPU / Insumo)</span>
                <span className="col-span-1 text-center">Und</span>
                <span className="col-span-1.5 text-right">Qtd Prop.</span>
                <span className="col-span-1.5 text-right">Custo Unit</span>
                <span className="col-span-2 text-right">Custo Total</span>
              </div>

              <div className="divide-y divide-stone-200 max-h-[600px] overflow-y-auto">
                {etapas.length === 0 ? (
                  <div className="p-8 text-center text-stone-400 italic text-xs">
                    Nenhuma etapa cadastrada neste orçamento.
                  </div>
                ) : (
                  etapas.map((etapa, idxEtapa) => {
                    const numEtapa = idxEtapa + 1;
                    const etapaId = etapa.id || `etapa-${idxEtapa}`;
                    const isEtapaAberta = !!etapasExpandidas[etapaId];

                    return (
                      <div key={etapaId} className="bg-stone-50/30">
                        <div 
                          className="grid grid-cols-12 gap-2 px-4 py-2 bg-stone-200/60 text-stone-800 text-xs font-bold items-center uppercase tracking-wide cursor-pointer hover:bg-stone-200 select-none"
                          onClick={() => setEtapasExpandidas(p => ({ ...p, [etapaId]: !isEtapaAberta }))}
                        >
                          <span className="col-span-10 flex items-center gap-1.5">
                            {isEtapaAberta ? <ChevronDown size={14} className="text-stone-500 shrink-0" /> : <ChevronRight size={14} className="text-stone-500 shrink-0" />}
                            <span className="truncate">{numEtapa}. {etapa.nome}</span>
                          </span>
                          <span className="col-span-2 text-right font-mono">
                            R$ {fmt((etapa.itens || []).reduce((acc, it) => acc + (num(it.quantidade) * cpuValorUnit(it.insumos, cpus, catalogMap)), 0))}
                          </span>
                        </div>

                        {isEtapaAberta && (etapa.itens || []).map((item, idxItem) => {
                          const numCpu = `${numEtapa}.${idxItem + 1}`;
                          const itemId = item.id || `item-${numCpu}`;
                          const isCpuAberta = !!cpusExpandidas[itemId];
                          const qtdItem = num(item.quantidade);
                          const custoUnitCpu = cpuValorUnit(item.insumos, cpus, catalogMap);

                          return (
                            <div key={itemId} className="border-b border-stone-100">
                              <div 
                                onClick={() => setCpusExpandidas(p => ({ ...p, [itemId]: !isCpuAberta }))}
                                className="grid grid-cols-12 gap-2 px-4 py-2 bg-white text-xs items-center font-semibold text-stone-700 pl-8 cursor-pointer hover:bg-stone-50 select-none"
                              >
                                <span className="col-span-6 truncate text-stone-900 flex items-center gap-1">
                                  {isCpuAberta ? <ChevronDown size={13} className="text-stone-400 shrink-0" /> : <ChevronRight size={13} className="text-stone-400 shrink-0" />}
                                  {numCpu}. {item.codigo ? `[${item.codigo}] ` : ""}{item.servico || item.descricao}
                                </span>
                                <span className="col-span-1 text-center font-mono text-stone-400">{item.unidade}</span>
                                <span className="col-span-1.5 text-right font-mono">{qtdItem.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</span>
                                <span className="col-span-1.5 text-right font-mono text-stone-400">R$ {fmt(custoUnitCpu)}</span>
                                <span className="col-span-2 text-right font-mono text-stone-800">
                                  R$ {fmt(qtdItem * custoUnitCpu)}
                                </span>
                              </div>

                              {isCpuAberta && (item.insumos || []).length > 0 && (
                                <div className="bg-stone-50/50 divide-y divide-stone-100/60 border-t border-b border-stone-100">
                                  {(item.insumos || []).map((insumo, idxInsumo) => {
                                    const numInsumo = `${numCpu}.${idxInsumo + 1}`;
                                    const precoUnit = insumoValorUnitario(insumo, cpus, catalogMap);
                                    const qtdCalculada = num(insumo.coeficiente) * qtdItem;
                                    const custoTotalInsumo = qtdCalculada * precoUnit;

                                    return (
                                      <div key={insumo.id || idxInsumo} className="grid grid-cols-12 gap-2 px-4 py-1.5 text-[11px] items-center text-stone-600 pl-14 hover:bg-stone-100/40">
                                        <span className="col-span-6 truncate uppercase font-sans text-stone-500">
                                          {numInsumo}. <span className="text-[9px] font-mono font-bold text-stone-400 border border-stone-200 px-1 py-0.5 rounded bg-white mr-1">{insumo.tipo}</span> {insumo.descricao}
                                        </span>
                                        <span className="col-span-1 text-center font-mono text-stone-400 uppercase text-[10px]">{insumo.unidade || "un"}</span>
                                        <span className="col-span-1.5 text-right font-mono text-stone-600">
                                          {qtdCalculada.toLocaleString("pt-BR", { minimumFractionDigits: 3, maximumFractionDigits: 3 })}
                                        </span>
                                        <span className="col-span-1.5 text-right font-mono text-stone-400">R$ {fmt(precoUnit)}</span>
                                        <span className="col-span-2 text-right font-mono font-medium text-stone-700">
                                          R$ {fmt(custoTotalInsumo)}
                                        </span>
                                      </div>
                                    );
                                  })}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    );
                  })
                )}

                {/* LINHA DE TOTAL GERAL DA PLANILHA DE CUSTO */}
                <div className="grid grid-cols-12 gap-2 px-4 py-3 bg-stone-900 text-white text-sm font-semibold uppercase tracking-wider">
                  <span className="col-span-6">CUSTO DIRETO TOTAL</span>
                  <span className="col-span-1"></span>
                  <span className="col-span-1.5"></span>
                  <span className="col-span-1.5"></span>
                  <span className="col-span-2 text-right font-mono text-amber-400">
                    R$ {fmt(grandTotal)}
                  </span>
                </div>
              </div>
            </div>
          </div>
        )}

        {tab === "precovenda" && projetoAtivo && (
          <div className="bg-white border border-stone-200 shadow-sm rounded-lg overflow-hidden p-5 space-y-4">
            <div className="flex justify-between items-center">
              <div>
                <h2 className="text-base font-semibold text-stone-800">Planilha de Preço de Venda (Custo + BDI)</h2>
                <p className="text-xs text-stone-500">
                  Visualização hierárquica por Etapa / CPU / Insumos aplicando BDI Geral de {fmt(bdiCalc.bdiRate * 100)}% {bdiCalc.faturamentoDireto && `e BDI de Materiais de ${fmt(bdiCalc.bdiRateMateriais * 100)}%`}. Valor Comercial Fechado: <span className="font-bold text-stone-800 font-mono">R$ {fmt(bdiCalc.valorVenda)}</span>
                </p>
              </div>
              <div className="flex gap-2">
                <button 
                  onClick={() => {
                    const obj = {};
                    etapas.forEach((e, idxE) => { 
                      const etapaId = e.id || `etapa-${idxE}`;
                      obj[etapaId] = true; 
                      (e.itens || []).forEach((it, idxIt) => { 
                        const numCpu = `${idxE + 1}.${idxIt + 1}`;
                        const itemId = it.id || `item-${numCpu}`;
                        obj[itemId] = true; 
                      }); 
                    });
                    setEtapasExpandidas(obj); setCpusExpandidas(obj);
                  }}
                  className="px-2 py-1 text-[11px] font-medium border border-stone-200 rounded hover:bg-stone-50 text-stone-600"
                >
                  Expandir Tudo
                </button>
                <button 
                  onClick={() => { setEtapasExpandidas({}); setCpusExpandidas({}); }}
                  className="px-2 py-1 text-[11px] font-medium border border-stone-200 rounded hover:bg-stone-50 text-stone-600"
                >
                  Recolher Tudo
                </button>

                {/* EXCEL DA PLANILHA DE VENDA (CORRIGIDO PARA LER O BDI DE MATERIAIS) */}
                <button 
                  onClick={() => {
                    const data = [];
                    data.push(["ESTRUTURA", "DESCRIÇÃO", "UND", "QTD PROP.", "PREÇO UNIT VENDA", "TOTAL VENDA"]);
                    etapas.forEach((etapa, idxE) => {
                      let totalEtapaVenda = 0;
                      (etapa.itens || []).forEach(it => {
                        const qCpu = num(it.quantidade);
                        (it.insumos || []).forEach(ins => {
                          const tIn = String(ins.tipo || "").toUpperCase().trim();
                          const cIn = num(ins.coeficiente) * qCpu * insumoValorUnitario(ins, cpus, catalogMap);
                          const isMat = bdiCalc.faturamentoDireto && (tIn === "MAT" || tIn === "MATERIAL" || (!tIn.includes("MO") && !tIn.includes("MÃO") && !tIn.includes("MAO") && !tIn.includes("EQUIP")));
                          totalEtapaVenda += cIn * (isMat ? bdiCalc.FatorBdiMateriais : bdiCalc.FatorBdi);
                        });
                      });

                      data.push([`${idxE + 1}`, etapa.nome, "", "", "", totalEtapaVenda]);
                      
                      (etapa.itens || []).forEach((item, idxI) => {
                        const numCpu = `${idxE + 1}.${idxI + 1}`;
                        let totalItemVenda = 0;
                        (item.insumos || []).forEach(ins => {
                          const tIn = String(ins.tipo || "").toUpperCase().trim();
                          const cIn = num(ins.coeficiente) * num(item.quantidade) * insumoValorUnitario(ins, cpus, catalogMap);
                          const isMat = bdiCalc.faturamentoDireto && (tIn === "MAT" || tIn === "MATERIAL" || (!tIn.includes("MO") && !tIn.includes("MÃO") && !tIn.includes("MAO") && !tIn.includes("EQUIP")));
                          totalItemVenda += cIn * (isMat ? bdiCalc.FatorBdiMateriais : bdiCalc.FatorBdi);
                        });

                        data.push([numCpu, item.servico || item.descricao, item.unidade, num(item.quantidade), totalItemVenda / num(item.quantidade), totalItemVenda]);
                        
                        (item.insumos || []).forEach((ins, idxIn) => {
                          const tIn = String(ins.tipo || "").toUpperCase().trim();
                          const custoUnit = insumoValorUnitario(ins, cpus, catalogMap);
                          const isMat = bdiCalc.faturamentoDireto && (tIn === "MAT" || tIn === "MATERIAL" || (!tIn.includes("MO") && !tIn.includes("MÃO") && !tIn.includes("MAO") && !tIn.includes("EQUIP")));
                          const fatBdi = isMat ? bdiCalc.FatorBdiMateriais : bdiCalc.FatorBdi;
                          
                          data.push([`${numCpu}.${idxIn + 1}`, `[${ins.tipo}] ${ins.descricao}`, ins.unidade || "un", num(ins.coeficiente) * num(item.quantidade), custoUnit * fatBdi, (num(ins.coeficiente) * num(item.quantidade)) * custoUnit * fatBdi]);
                        });
                      });
                    });
                    const ws = XLSX.utils.aoa_to_sheet(data);
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, "Preço de Venda");
                    XLSX.writeFile(wb, `${projetoAtivo.nome || "Orcamento"}_Preco_Venda.xlsx`);
                  }}
                  className="px-2 py-1 text-[11px] font-medium border border-emerald-200 text-emerald-700 bg-emerald-50/50 rounded hover:bg-emerald-50 flex items-center gap-1"
                >
                  <Download size={12} /> Excel (.xlsx)
                </button>

                <button
                  onClick={() =>
                    exportarPropostaXlsx({
                      projeto: projetoAtivo,
                      etapas,
                      bdiCalc,
                      cpus,
                      catalogMap,
                    })
                  }
                  className="px-2 py-1 text-[11px] font-medium border border-stone-900 text-white bg-stone-900 rounded hover:bg-stone-800 flex items-center gap-1"
                >
                  <FileText size={12} /> Proposta .xlsx
                </button>

                <button
                  onClick={() =>
                    gerarPropostaPdf({
                      projeto: projetoAtivo,
                      cliente: clienteAtivo,
                      etapas,
                      bdiCalc,
                      cpus,
                      catalogMap,
                    })
                  }
                  className="px-2 py-1 text-[11px] font-medium border border-red-200 text-red-700 bg-red-50/50 rounded hover:bg-red-50 flex items-center gap-1"
                >
                  <Download size={12} /> Proposta PDF
                </button>

                {/* PDF LIMPO DA PLANILHA DE VENDA */}
                <button 
                  onClick={() => {
                    const tituloOriginal = document.title;
                    document.title = `${projetoAtivo.nome || "Orcamento"}_Preco_Venda`;
                    const estiloPrint = document.createElement("style");
                    estiloPrint.innerHTML = `
                      @media print {
                        body * { visibility: hidden; }
                        #area-planilha-venda, #area-planilha-venda * { visibility: visible; }
                        #area-planilha-venda { position: absolute; left: 0; top: 0; width: 100%; background: white !important; }
                      }
                    `;
                    document.head.appendChild(estiloPrint);
                    window.print();
                    document.head.removeChild(estiloPrint);
                    document.title = tituloOriginal;
                  }}
                  className="px-2 py-1 text-[11px] font-medium border border-red-200 text-red-700 bg-red-50/50 rounded hover:bg-red-50 flex items-center gap-1"
                >
                  <Download size={12} /> PDF (.pdf)
                </button>
              </div>
            </div>

            <div id="area-planilha-venda" className="border border-stone-200 rounded-lg overflow-hidden bg-white">
              <div className="grid grid-cols-12 gap-2 px-4 py-2.5 bg-stone-100 border-b border-stone-200 text-stone-500 font-semibold text-[11px] uppercase tracking-wider">
                <span className="col-span-6">Estrutura (Etapa / CPU / Insumo)</span>
                <span className="col-span-1 text-center">Und</span>
                <span className="col-span-1.5 text-right">Qtd Prop.</span>
                <span className="col-span-1.5 text-right">Preço Unit Venda</span>
                <span className="col-span-2 text-right">Total Venda</span>
              </div>

              <div className="divide-y divide-stone-200 max-h-[600px] overflow-y-auto">
                {etapas.length === 0 ? (
                  <div className="p-8 text-center text-stone-400 italic text-xs">
                    Nenhuma etapa cadastrada neste orçamento.
                  </div>
                ) : (
                  etapas.map((etapa, idxEtapa) => {
                    const numEtapa = idxEtapa + 1;
                    const etapaId = etapa.id || `etapa-${idxEtapa}`;
                    const isEtapaAberta = !!etapasExpandidas[etapaId];

                    let totalEtapaComBdi = 0;
                    (etapa.itens || []).forEach(it => {
                      const qCpu = num(it.quantidade);
                      (it.insumos || []).forEach(ins => {
                        const tIn = String(ins.tipo || "").toUpperCase().trim();
                        const cIn = num(ins.coeficiente) * qCpu * insumoValorUnitario(ins, cpus, catalogMap);
                        const isMat = bdiCalc.faturamentoDireto && (tIn === "MAT" || tIn === "MATERIAL" || (!tIn.includes("MO") && !tIn.includes("MÃO") && !tIn.includes("MAO") && !tIn.includes("EQUIP")));
                        totalEtapaComBdi += cIn * (isMat ? bdiCalc.FatorBdiMateriais : bdiCalc.FatorBdi);
                      });
                    });

                    return (
                      <div key={etapaId} className="bg-stone-50/30">
                        <div 
                          onClick={() => setEtapasExpandidas(p => ({ ...p, [etapaId]: !isEtapaAberta }))}
                          className="grid grid-cols-12 gap-2 px-4 py-2 bg-stone-200/60 text-stone-800 text-xs font-bold items-center uppercase tracking-wide cursor-pointer hover:bg-stone-200 select-none"
                        >
                          <span className="col-span-10 flex items-center gap-1.5">
                            {isEtapaAberta ? <ChevronDown size={14} className="text-stone-500" /> : <ChevronRight size={14} className="text-stone-500" />}
                            {numEtapa}. {etapa.nome}
                          </span>
                          <span className="col-span-2 text-right font-mono text-stone-900">
                            R$ {fmt(totalEtapaComBdi)}
                          </span>
                        </div>

                        {isEtapaAberta && (etapa.itens || []).map((item, idxItem) => {
                          const numCpu = `${numEtapa}.${idxItem + 1}`;
                          const itemId = item.id || `item-${numCpu}`;
                          const isCpuAberta = !!cpusExpandidas[itemId];
                          const qtdItem = num(item.quantidade);

                          let totalCpuComBdi = 0;
                          (item.insumos || []).forEach(ins => {
                            const tIn = String(ins.tipo || "").toUpperCase().trim();
                            const cIn = num(ins.coeficiente) * num(ins.valorUnitario);
                            const isMat = bdiCalc.faturamentoDireto && (tIn === "MAT" || tIn === "MATERIAL" || (!tIn.includes("MO") && !tIn.includes("MÃO") && !tIn.includes("MAO") && !tIn.includes("EQUIP")));
                            totalCpuComBdi += cIn * (isMat ? bdiCalc.FatorBdiMateriais : bdiCalc.FatorBdi);
                          });

                          return (
                            <div key={itemId} className="border-b border-stone-100">
                              <div 
                                onClick={() => setCpusExpandidas(p => ({ ...p, [itemId]: !isCpuAberta }))}
                                className="grid grid-cols-12 gap-2 px-4 py-2 bg-white text-xs items-center font-semibold text-stone-700 pl-8 cursor-pointer hover:bg-stone-50 select-none"
                              >
                                <span className="col-span-6 truncate text-stone-900 flex items-center gap-1">
                                  {isCpuAberta ? <ChevronDown size={13} className="text-stone-400 shrink-0" /> : <ChevronRight size={13} className="text-stone-400 shrink-0" />}
                                  {numCpu}. {item.codigo ? `[${item.codigo}] ` : ""}{item.servico || item.descricao}
                                </span>
                                <span className="col-span-1 text-center font-mono text-stone-400">{item.unidade}</span>
                                <span className="col-span-1.5 text-right font-mono">{qtdItem.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</span>
                                <span className="col-span-1.5 text-right font-mono text-stone-500">R$ {fmt(totalCpuComBdi)}</span>
                                <span className="col-span-2 text-right font-mono text-stone-900">
                                  R$ {fmt(qtdItem * totalCpuComBdi)}
                                </span>
                              </div>

                              {isCpuAberta && (item.insumos || []).length > 0 && (
                                <div className="bg-stone-50/50 divide-y divide-stone-100/60 border-t border-b border-stone-100">
                                  {(item.insumos || []).map((insumo, idxInsumo) => {
                                    const numInsumo = `${numCpu}.${idxInsumo + 1}`;
                                    const tIn = String(insumo.tipo || "").toUpperCase().trim();
                                    const entry = catalogMap.get(precoKey(insumo.descricao));
                                    const custoUnit = entry && entry.valorUnitario !== "" ? num(entry.valorUnitario) : num(insumo.valorUnitario);
                                    
                                    const isMat = bdiCalc.faturamentoDireto && (tIn === "MAT" || tIn === "MATERIAL" || (!tIn.includes("MO") && !tIn.includes("MÃO") && !tIn.includes("MAO") && !tIn.includes("EQUIP")));
                                    const precoVendaInsumo = custoUnit * (isMat ? bdiCalc.FatorBdiMateriais : bdiCalc.FatorBdi);
                                    
                                    const qtdCalculada = num(insumo.coeficiente) * qtdItem;
                                    const vendaTotalInsumo = qtdCalculada * precoVendaInsumo;

                                    return (
                                      <div key={insumo.id || idxInsumo} className="grid grid-cols-12 gap-2 px-4 py-1.5 text-[11px] items-center text-stone-600 pl-14 hover:bg-stone-100/40">
                                        <span className="col-span-6 truncate uppercase font-sans text-stone-500">
                                          {numInsumo}. <span className="text-[9px] font-mono font-bold text-stone-400 border border-stone-200 px-1 py-0.5 rounded bg-white mr-1">{insumo.tipo}</span> {insumo.descricao}
                                        </span>
                                        <span className="col-span-1 text-center font-mono text-stone-400 uppercase text-[10px]">{insumo.unidade || "un"}</span>
                                        <span className="col-span-1.5 text-right font-mono text-stone-600">
                                          {qtdCalculada.toLocaleString("pt-BR", { minimumFractionDigits: 3, maximumFractionDigits: 3 })}
                                        </span>
                                        <span className="col-span-1.5 text-right font-mono text-stone-400">R$ {fmt(precoVendaInsumo)}</span>
                                        <span className="col-span-2 text-right font-mono font-medium text-blue-700">
                                          R$ {fmt(vendaTotalInsumo)}
                                        </span>
                                      </div>
                                    );
                                  })}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    );
                  })
                )}
                
                {/* LINHA DE TOTAIS GERAIS DA PLANILHA DE VENDA */}
                <div className="grid grid-cols-12 gap-2 px-4 py-3 bg-stone-900 text-white text-sm font-semibold uppercase tracking-wider">
                  <span className="col-span-6">VALOR FINAL DE VENDA COM BDI</span>
                  <span className="col-span-1"></span>
                  <span className="col-span-1.5"></span>
                  <span className="col-span-1.5"></span>
                  <span className="col-span-2 text-right font-mono text-amber-400">
                    R$ {fmt(bdiCalc.valorVenda)}
                  </span>
                </div>
              </div>
            </div>
          </div>
        )}

        {tab === "maoobra" && projetoAtivo && (
          <div className="bg-white border border-stone-200 shadow-sm rounded-lg overflow-hidden p-5 space-y-4">
            <div>
              <h2 className="text-base font-semibold text-stone-800">Consolidado Qualitativo de Mão de Obra</h2>
              <p className="text-xs text-stone-500">Visualização agrupada de todas as horas e custos de mão de obra alocados no orçamento.</p>
            </div>

            <div className="border border-stone-200 rounded-lg overflow-hidden">
              <div className="grid grid-cols-12 gap-2 px-4 py-2.5 bg-stone-100 border-b border-stone-200 text-stone-500 font-semibold text-[11px] uppercase tracking-wider">
                <span className="col-span-6">Descrição do Profissional</span>
                <span className="col-span-1 text-center">Und</span>
                <span className="col-span-1.5 text-right">Horas Totais</span>
                <span className="col-span-1.5 text-right">Valor Unit.</span>
                <span className="col-span-2 text-right">Subtotal Direto</span>
              </div>

              <div className="divide-y divide-stone-200 max-h-[500px] overflow-y-auto">
                {(() => {
                  const mos = new Map();
                  (etapas || []).forEach(e => {
                    (e.itens || []).forEach(it => {
                      const qtdCpu = num(it.quantidade);
                      (it.insumos || []).forEach(ins => {
                        const tipo = String(ins.tipo || "").toUpperCase().trim();
                        if (tipo === "MO" || tipo.includes("MÃO") || tipo.includes("MAO")) {
                          if (!String(ins.descricao || "").trim()) return;
                          
                          const chave = ins.descricao.trim().toLowerCase();
                          const qtdCalc = num(ins.coeficiente) * qtdCpu;
                          
                          const entry = catalogMap.get(precoKey(ins.descricao));
                          const vUnit = entry && entry.valorUnitario !== "" ? num(entry.valorUnitario) : num(ins.valorUnitario);
                          
                          if (mos.has(chave)) {
                            const existente = mos.get(chave);
                            existente.qtd += qtdCalc;
                            existente.total += qtdCalc * vUnit;
                          } else {
                            mos.set(chave, {
                              descricao: ins.descricao,
                              unidade: ins.unidade || "h",
                              qtd: qtdCalc,
                              valorUnit: vUnit,
                              total: qtdCalc * vUnit
                            });
                          }
                        }
                      });
                    });
                  });

                  const listaMo = Array.from(mos.values()).sort((a, b) => b.total - a.total);
                  if (listaMo.length === 0) {
                    return <div className="p-8 text-center text-stone-400 italic text-xs">Nenhuma mão de obra localizada no orçamento.</div>;
                  }

                  return (
                    <>
                      {listaMo.map((r, idx) => (
                        <div key={idx} className="grid grid-cols-12 gap-2 px-4 py-2 text-xs items-center hover:bg-stone-50/60 uppercase">
                          <span className="col-span-6 font-medium text-stone-800 truncate">{r.descricao}</span>
                          <span className="col-span-1 text-center font-mono text-stone-400">{r.unidade}</span>
                          <span className="col-span-1.5 text-right font-mono text-stone-900">{r.qtd.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
                          <span className="col-span-1.5 text-right font-mono text-stone-400">R$ {fmt(r.valorUnit)}</span>
                          <span className="col-span-2 text-right font-mono font-semibold text-stone-700">R$ {fmt(r.total)}</span>
                        </div>
                      ))}
                      
                      <div className="grid grid-cols-12 gap-2 px-4 py-3 bg-stone-900 text-white text-sm font-semibold">
                        <span className="col-span-6">TOTAL GERAL EM MÃO DE OBRA</span>
                        <span className="col-span-1"></span>
                        <span className="col-span-1.5 text-right font-mono text-stone-300">
                          {listaMo.reduce((acc, curr) => acc + curr.qtd, 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                        </span>
                        <span className="col-span-1.5"></span>
                        <span className="col-span-2 text-right font-mono text-amber-400">
                          R$ {fmt(listaMo.reduce((acc, curr) => acc + curr.total, 0))}
                        </span>
                      </div>
                    </>
                  );
                })()}
              </div>
            </div>
          </div>
        )}

        {tab === "materiais" && projetoAtivo && (
          <div className="bg-white border border-stone-200 shadow-sm rounded-lg overflow-hidden p-5 space-y-4">
            <div>
              <h2 className="text-base font-semibold text-stone-800">Quantitativo de Materiais</h2>
              <p className="text-xs text-stone-500">Consolidação de todos os materiais físicos consumidos nas CPUs do orçamento ativo.</p>
            </div>

            <div className="border border-stone-200 rounded-lg overflow-hidden">
              <div className="grid grid-cols-12 gap-2 px-4 py-2.5 bg-stone-100 border-b border-stone-200 text-stone-500 font-semibold text-[11px] uppercase tracking-wider">
                <span className="col-span-6">Material</span>
                <span className="col-span-1 text-center">Und</span>
                <span className="col-span-1.5 text-right">Qtd Total</span>
                <span className="col-span-1.5 text-right">Preço Unit.</span>
                <span className="col-span-2 text-right">Total Bruto</span>
              </div>

              <div className="divide-y divide-stone-200 max-h-[500px] overflow-y-auto">
                {(() => {
                  const mats = new Map();
                  (etapas || []).forEach(e => {
                    (e.itens || []).forEach(it => {
                      const qtdCpu = num(it.quantidade);
                      (it.insumos || []).forEach(ins => {
                        const tipo = String(ins.tipo || "").toUpperCase().trim();
                        if (tipo === "MAT" || tipo === "MATERIAL" || (!tipo.includes("MO") && !tipo.includes("MÃO") && !tipo.includes("MAO") && !tipo.includes("EQUIP"))) {
                          if (!String(ins.descricao || "").trim()) return; 
                          
                          const chave = ins.descricao.trim().toLowerCase();
                          const qtdCalc = num(ins.coeficiente) * qtdCpu;
                          
                          const entry = catalogMap.get(precoKey(ins.descricao));
                          const vUnit = entry && entry.valorUnitario !== "" ? num(entry.valorUnitario) : num(ins.valorUnitario);
                          
                          if (mats.has(chave)) {
                            const existente = mats.get(chave);
                            existente.qtd += qtdCalc;
                            existente.total += qtdCalc * vUnit;
                          } else {
                            mats.set(chave, {
                              descricao: ins.descricao,
                              unidade: ins.unidade || "un",
                              qtd: qtdCalc,
                              valorUnit: vUnit,
                              total: qtdCalc * vUnit
                            });
                          }
                        }
                      });
                    });
                  });

                  const listaMats = Array.from(mats.values()).sort((a, b) => b.total - a.total);
                  if (listaMats.length === 0) {
                    return <div className="p-8 text-center text-stone-400 italic text-xs">Nenhum material localizado neste orçamento.</div>;
                  }

                  return (
                    <>
                      {listaMats.map((r, idx) => (
                        <div key={idx} className="grid grid-cols-12 gap-2 px-4 py-2 text-xs items-center hover:bg-stone-50/60 uppercase">
                          <span className="col-span-6 font-medium text-stone-800 truncate">{r.descricao}</span>
                          <span className="col-span-1 text-center font-mono text-stone-400">{r.unidade}</span>
                          <span className="col-span-1.5 text-right font-mono text-stone-900">{r.qtd.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 3 })}</span>
                          <span className="col-span-1.5 text-right font-mono text-stone-400">R$ {fmt(r.valorUnit)}</span>
                          <span className="col-span-2 text-right font-mono font-semibold text-emerald-700">R$ {fmt(r.total)}</span>
                        </div>
                      ))}
                      
                      <div className="grid grid-cols-12 gap-2 px-4 py-3 bg-stone-900 text-white text-sm font-semibold">
                        <span className="col-span-6">TOTAL GERAL EM MATERIAIS</span>
                        <span className="col-span-1"></span>
                        <span className="col-span-1.5"></span>
                        <span className="col-span-1.5"></span>
                        <span className="col-span-2 text-right font-mono text-amber-400">
                          R$ {fmt(listaMats.reduce((acc, curr) => acc + curr.total, 0))}
                        </span>
                      </div>
                    </>
                  );
                })()}
              </div>
            </div>
          </div>
        )}
        {tab === "precos" && projetoAtivo && (
          <PrecosTab
            catalog={catalog}
            onUpsert={upsertPreco}
            onRemove={removePreco}
            onApplyToCpus={aplicarPrecoNoOrcamentoAtivo}
            onApplyAllToCpus={aplicarTodosPrecosNoOrcamentoAtivo}
          />
        )}
        </main>
      </div>
    </div>
  );
}

function CadastroCliente({ projeto, cliente, setProjetos, setCliente, completo, onContinuar }) {
  const atualizarCampo = (campo, valor) => {
    setCliente((prev) => ({ ...prev, [campo]: valor }));
  };

  const atualizarNomeProjeto = (valor) => {
    setProjetos((prev) => prev.map((p) => (p.id === projeto.id ? { ...p, nome: valor } : p)));
  };

  const campoBase =
    "w-full border rounded-lg px-3 py-2 text-sm outline-none bg-white focus:ring-1";
  const campoObrigatorio = (valor) =>
    `${campoBase} ${
      String(valor || "").trim()
        ? "border-stone-300 focus:border-stone-700 focus:ring-stone-700"
        : "border-amber-300 bg-amber-50/40 focus:border-amber-500 focus:ring-amber-500"
    }`;

  return (
    <div className="bg-white border border-stone-200 shadow-sm rounded-lg overflow-hidden">
      <div className="px-5 py-4 border-b border-stone-200 flex items-start justify-between gap-4">
        <div>
          <h2 className="text-base font-semibold text-stone-800 flex items-center gap-2">
            <User size={17} /> Cadastro do Cliente
          </h2>
          <p className="text-xs text-stone-500 mt-1">
            Dados vinculados a este orçamento para uso em planilhas, propostas e documentos comerciais.
          </p>
        </div>
        <span
          className={`text-[11px] font-semibold px-2 py-1 rounded border ${
            completo
              ? "bg-emerald-50 border-emerald-200 text-emerald-700"
              : "bg-amber-50 border-amber-200 text-amber-700"
          }`}
        >
          {completo ? "Cadastro completo" : "Nome e local obrigatórios"}
        </span>
      </div>

      <div className="p-5 space-y-5">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <CampoCliente
            label="Nome do orçamento"
            value={projeto.nome || ""}
            onChange={atualizarNomeProjeto}
            icon={<FileText size={14} />}
          />
          <CampoCliente
            label="Nome do cliente"
            value={cliente.nome || ""}
            onChange={(valor) => atualizarCampo("nome", valor)}
            icon={<User size={14} />}
            inputClassName={campoObrigatorio(cliente.nome)}
            required
          />
          <CampoCliente
            label="Local da obra"
            value={cliente.local || ""}
            onChange={(valor) => atualizarCampo("local", valor)}
            icon={<MapPin size={14} />}
            inputClassName={campoObrigatorio(cliente.local)}
            required
          />
        </div>

        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <CampoCliente
            label="Contato"
            value={cliente.contato || ""}
            onChange={(valor) => atualizarCampo("contato", valor)}
            icon={<Building2 size={14} />}
          />
          <CampoCliente
            label="Telefone"
            value={cliente.telefone || ""}
            onChange={(valor) => atualizarCampo("telefone", valor)}
            icon={<Phone size={14} />}
          />
          <CampoCliente
            label="E-mail"
            value={cliente.email || ""}
            onChange={(valor) => atualizarCampo("email", valor)}
            icon={<Mail size={14} />}
            type="email"
          />
          <CampoCliente
            label="CPF/CNPJ"
            value={cliente.documento || ""}
            onChange={(valor) => atualizarCampo("documento", valor)}
            icon={<FileText size={14} />}
          />
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <CampoCliente
            label="Endereço"
            value={cliente.endereco || ""}
            onChange={(valor) => atualizarCampo("endereco", valor)}
            icon={<MapPin size={14} />}
          />
          <div>
            <label className="text-xs font-medium text-stone-500 mb-1 flex items-center gap-1.5">
              <FileText size={14} /> Observações
            </label>
            <textarea
              value={cliente.observacoes || ""}
              onChange={(e) => atualizarCampo("observacoes", e.target.value)}
              rows={3}
              className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm outline-none bg-white focus:border-stone-700 focus:ring-1 focus:ring-stone-700 resize-none"
              placeholder="Informações adicionais para proposta ou visita técnica"
            />
          </div>
        </div>

        <div className="pt-3 border-t border-stone-200 flex justify-end">
          <button
            type="button"
            onClick={onContinuar}
            className={`px-4 py-2 rounded-lg text-sm font-medium ${
              completo
                ? "bg-stone-900 text-white hover:bg-stone-800"
                : "bg-stone-200 text-stone-500 cursor-not-allowed"
            }`}
          >
            Continuar para Lançamento CPU
          </button>
        </div>
      </div>
    </div>
  );
}

function CampoCliente({ label, value, onChange, icon, required, type = "text", inputClassName }) {
  return (
    <div>
      <label className="text-xs font-medium text-stone-500 mb-1 flex items-center gap-1.5">
        {icon} {label} {required && <span className="text-amber-600">*</span>}
      </label>
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className={inputClassName || "w-full border border-stone-300 rounded-lg px-3 py-2 text-sm outline-none bg-white focus:border-stone-700 focus:ring-1 focus:ring-stone-700"}
      />
    </div>
  );
}

function TabBtn({ active, onClick, icon, children, disabled }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-medium border-b-2 -mb-px transition-colors ${
        disabled ? "opacity-40 cursor-not-allowed" : ""
      } ${active ? "border-stone-900 text-stone-900" : "border-transparent text-stone-400 hover:text-stone-600"}`}
    >
      {icon}
      {children}
    </button>
  );
}

function SideTabBtn({ active, onClick, icon, children, disabled }) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      className={`flex items-center gap-2 w-full px-3 py-2 text-sm font-medium rounded-md transition-colors text-left whitespace-nowrap ${
        disabled ? "opacity-40 cursor-not-allowed" : ""
      } ${
        active
          ? "bg-stone-900 text-white"
          : "text-stone-500 hover:text-stone-900 hover:bg-stone-100"
      }`}
    >
      <span className={active ? "text-white" : "text-stone-400"}>{icon}</span>
      <span className="truncate">{children}</span>
    </button>
  );
}

function CpuLibrary({ cpus, setCpus, fileInputRef, catalogMap }) {
  const [query, setQuery] = useState("");
  const [fonteFiltro, setFonteFiltro] = useState("Todas");
  const [editing, setEditing] = useState(null);
  const [expanded, setExpanded] = useState({});
  const [importMsg, setImportMsg] = useState("");
  
  // NOVO: Controla qual linha do resultado filtrado está focada pelo teclado
  const [activeIndex, setActiveIndex] = useState(-1);

  const fontes = useMemo(
    () => ["Todas", ...Array.from(new Set(cpus.map((c) => c.fonte).filter(Boolean)))],
    [cpus]
  );

  const cpuSearchIndex = useMemo(
    () =>
      cpus.map((c) => ({
        cpu: c,
        haystack: norm(
          `${c.codigo || ""} ${c.descricao || ""} ${c.fonte || ""} ${(c.insumos || []).map((i) => i.descricao).join(" ")}`
        ),
      })),
    [cpus]
  );

  const queryTokens = useMemo(() => {
    const tokens = [];
    const re = /"([^"]+)"|(\S+)/g;
    let m;
    while ((m = re.exec(query)) !== null) {
      const t = norm((m[1] || m[2] || "").trim());
      if (t) tokens.push(t);
    }
    return tokens;
  }, [query]);

  const filtered = useMemo(
    () =>
      cpuSearchIndex
        .filter(({ cpu, haystack }) => {
          const matchesFonte = fonteFiltro === "Todas" || cpu.fonte === fonteFiltro;
          if (!matchesFonte) return false;
          if (queryTokens.length === 0) return true;
          return queryTokens.every((t) => haystack.includes(t));
        })
        .map(({ cpu }) => cpu),
    [cpuSearchIndex, fonteFiltro, queryTokens]
  );

  const [confirmingDelete, setConfirmingDelete] = useState(null);

  const removeCpu = (id) => {
    setCpus(cpus.filter((c) => c.id !== id));
    setConfirmingDelete(null);
  };

  const duplicateCpu = (c) => {
    setCpus([...cpus, { ...c, id: uid(), codigo: c.codigo + " (cópia)", insumos: c.insumos.map((i) => ({ ...i, id: uid() })) }]);
  };

  const saveCpu = (cpu) => {
    if (cpus.find((c) => c.id === cpu.id)) {
      setCpus(cpus.map((c) => (c.id === cpu.id ? cpu : c)));
    } else {
      setCpus([...cpus, cpu]);
    }
    setEditing(null);
  };

  const cpuImportKey = (cpu) => norm(cpu.codigo || cpu.descricao);

  const insumoImportKey = (insumo) => norm(insumo.codigo || insumo.descricao);

  const normalizedInsumos = (insumos = []) =>
    insumos.map((i) => ({
      codigo: norm(i.codigo || ""),
      tipo: String(i.tipo || "").toUpperCase().trim(),
      descricao: norm(i.descricao || ""),
      unidade: norm(i.unidade || ""),
      coeficiente: num(i.coeficiente),
    }));

  const cpuMudou = (atual, importada) => {
    const baseAtual = {
      codigo: norm(atual.codigo || ""),
      descricao: norm(atual.descricao || ""),
      unidade: norm(atual.unidade || ""),
      insumos: normalizedInsumos(atual.insumos),
    };
    const baseImportada = {
      codigo: norm(importada.codigo || ""),
      descricao: norm(importada.descricao || ""),
      unidade: norm(importada.unidade || ""),
      insumos: normalizedInsumos(importada.insumos),
    };
    return JSON.stringify(baseAtual) !== JSON.stringify(baseImportada);
  };

  const mesclarInsumosImportados = (atuais = [], importados = []) => {
    const atuaisPorChave = new Map(atuais.map((i) => [insumoImportKey(i), i]));

    return importados.map((insumo) => {
      const existente = atuaisPorChave.get(insumoImportKey(insumo));
      return {
        ...insumo,
        id: existente?.id || insumo.id || uid(),
        valorUnitario:
          insumo.valorUnitario !== "" && insumo.valorUnitario !== null && insumo.valorUnitario !== undefined
            ? insumo.valorUnitario
            : existente?.valorUnitario ?? "",
      };
    });
  };

  const mesclarCpusImportadas = (atuais, importadas) => {
    const existentesPorChave = new Map(atuais.map((cpu) => [cpuImportKey(cpu), cpu]));
    let adicionadas = 0;
    let atualizadas = 0;
    let semMudanca = 0;

    const importadasPorChave = new Map();
    importadas.forEach((cpu) => {
      const chave = cpuImportKey(cpu);
      if (chave) importadasPorChave.set(chave, cpu);
    });

    const proximas = atuais.map((cpuAtual) => {
      const chave = cpuImportKey(cpuAtual);
      const cpuImportada = importadasPorChave.get(chave);
      if (!cpuImportada) return cpuAtual;

      importadasPorChave.delete(chave);
      if (!cpuMudou(cpuAtual, cpuImportada)) {
        semMudanca += 1;
        return cpuAtual;
      }

      atualizadas += 1;
      return {
        ...cpuAtual,
        codigo: cpuImportada.codigo || cpuAtual.codigo,
        fonte: cpuImportada.fonte || cpuAtual.fonte,
        descricao: cpuImportada.descricao || cpuAtual.descricao,
        unidade: cpuImportada.unidade || cpuAtual.unidade,
        insumos: mesclarInsumosImportados(cpuAtual.insumos, cpuImportada.insumos),
      };
    });

    importadasPorChave.forEach((cpuImportada) => {
      adicionadas += 1;
      proximas.push({
        ...cpuImportada,
        id: cpuImportada.id || uid(),
        insumos: (cpuImportada.insumos || []).map((i) => ({ ...i, id: i.id || uid() })),
      });
    });

    return { cpus: proximas, adicionadas, atualizadas, semMudanca };
  };

  // NOVO: Gerencia a navegação por setas e Enter na listagem
  const handleKeyDown = (evt) => {
    if (filtered.length === 0) return;

    if (evt.key === "ArrowDown") {
      evt.preventDefault();
      setActiveIndex((prev) => (prev + 1) % filtered.length);
    } else if (evt.key === "ArrowUp") {
      evt.preventDefault();
      setActiveIndex((prev) => (prev - 1 + filtered.length) % filtered.length);
    } else if (evt.key === "Enter") {
      if (activeIndex >= 0 && activeIndex < filtered.length) {
        evt.preventDefault();
        const targetCpu = filtered[activeIndex];
        setExpanded((prev) => ({ ...prev, [targetCpu.id]: !prev[targetCpu.id] }));
      }
    } else if (evt.key === "Escape") {
      setQuery("");
      setActiveIndex(-1);
    }
  };

  const handleImport = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setImportMsg("Lendo planilha...");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      const getField = (row, names) => {
        for (const key of Object.keys(row)) {
          if (names.includes(norm(key))) return row[key];
        }
        return "";
      };

      const headers = rows.length ? Object.keys(rows[0]).map(norm) : [];
      const hasInsumoColumn = headers.some((h) => ["insumo", "item", "insumo_descricao"].includes(h));

      const inferTipo = (desc) => {
        const d = norm(desc);
        if (/^chp\/|^chi\/|caminhao|trator|escavadeira|pa carregadeira|guindaste|compactador|motoniveladora|retroescavadeira/.test(d)) return "EQUIP";
        if (/servente|pedreiro|oficial|ajudante|encarregado|mestre de obras|carpinteiro|armador|eletricista|pintor/.test(d)) return "MO";
        return "MAT";
      };

      const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      const headerText = rawRows.slice(0, 5).flat().join(" ").toLowerCase();
      const isRelatorioSudecap = headerText.includes("relatório de composiç") || headerText.includes("relatorio de composic");
      const isEstruturaCodigoDescricao =
        norm(headerText).includes("codigo") &&
        norm(headerText).includes("descricao") &&
        norm(headerText).includes("consumo");

      const cell = (row, idx) => String(row[idx] ?? "").trim();
      const parseConsumo = (value) => {
        if (typeof value === "number") return value;
        const raw = String(value ?? "").trim();
        if (!raw) return 0;
        const normalized = raw.includes(",") ? raw.replace(/\./g, "").replace(",", ".") : raw;
        const parsed = Number(normalized);
        return Number.isFinite(parsed) ? parsed : 0;
      };

      const splitCodigoDescricao = (texto) => {
        const raw = String(texto ?? "").trim();
        const match = raw.match(/^([0-9]+(?:\.[0-9]+)+)\s*-\s*(.+)$/);
        if (!match) return { codigo: "", descricao: raw };
        return { codigo: match[1].trim(), descricao: match[2].trim() };
      };

      const isLinhaSeparadora = (texto) => {
        const t = String(texto ?? "").trim();
        return !t || /^[-–—]+$/.test(t);
      };

      const novas = [];

      if (isEstruturaCodigoDescricao) {
        const fonteNome = file.name.toLowerCase().includes("der") ? "DER-MG" : "SUDECAP";
        let atual = null;

        rawRows.slice(1).forEach((row) => {
          const codigo = cell(row, 0);
          const descricaoCompleta = cell(row, 1);
          const unidade = cell(row, 2);
          const consumoRaw = row[3];
          const temConsumo = String(consumoRaw ?? "").trim() !== "";

          if (isLinhaSeparadora(codigo) && isLinhaSeparadora(descricaoCompleta)) return;

          if (codigo && descricaoCompleta && unidade && !temConsumo) {
            atual = {
              id: uid(),
              codigo,
              fonte: fonteNome,
              descricao: descricaoCompleta.replace(/\s*-\s*$/, "").trim(),
              unidade: unidade || "un",
              insumos: [],
            };
            novas.push(atual);
            return;
          }

          if (atual && descricaoCompleta && unidade && temConsumo && !isLinhaSeparadora(descricaoCompleta)) {
            const { codigo: codigoNaDescricao, descricao } = splitCodigoDescricao(descricaoCompleta);
            const codigoInsumo = codigo || codigoNaDescricao;
            atual.insumos.push({
              id: uid(),
              tipo: inferTipo(descricao),
              codigo: codigoInsumo,
              descricao,
              unidade: unidade || "un",
              coeficiente: parseConsumo(consumoRaw),
              valorUnitario: "",
            });
          }
        });
      } else if (isRelatorioSudecap) {
        const fonteNome = file.name.toLowerCase().includes("der") ? "DER-MG" : "SUDECAP";
        let atual = null;
        rawRows.slice(3).forEach((row) => {
          const c0 = String(row[0] ?? "").trim();
          const c1 = String(row[1] ?? "").trim();
          const c2 = String(row[2] ?? "").trim();
          const und = String(row[7] ?? "").trim();
          const consumoRaw = row[9];
          if (c0) {
            if (und) {
              atual = { id: uid(), codigo: c0, fonte: fonteNome, descricao: c1, unidade: und, insumos: [] };
              novas.push(atual);
            } else {
              atual = null;
            }
          } else if (c1 && atual) {
            atual.insumos.push({
              id: uid(), tipo: inferTipo(c2), descricao: c2, unidade: und || "un", coeficiente: consumoRaw ? num(consumoRaw) : 0, valorUnitario: ""
            });
          }
        });
      } else if (hasInsumoColumn) {
        const grouped = {};
        rows.forEach((row) => {
          const codigo = String(getField(row, ["codigo", "código", "code"])).trim();
          if (!codigo) return;
          if (!grouped[codigo]) {
            grouped[codigo] = {
              id: uid(), codigo,
              fonte: String(getField(row, ["fonte", "tabela", "origem"])) || "Própria",
              descricao: String(getField(row, ["descricao", "descrição", "servico", "serviço"])),
              unidade: String(getField(row, ["unidade", "un", "unid"])) || "un",
              insumos: []
            };
          }
          const insumoDesc = String(getField(row, ["insumo", "item", "insumo_descricao"]));
          if (insumoDesc) {
            const rawValor = getField(row, ["valor_unitario", "valor unitário", "valor", "preco", "preço"]);
            grouped[codigo].insumos.push({
              id: uid(),
              tipo: (String(getField(row, ["tipo", "tipo_insumo"])).toUpperCase().includes("MAT") && "MAT") ||
                    (String(getField(row, ["tipo", "tipo_insumo"])).toUpperCase().includes("EQUIP") && "EQUIP") ||
                    (String(getField(row, ["tipo", "tipo_insumo"])).toUpperCase().includes("MO") && "MO") || inferTipo(insumoDesc),
              descricao: insumoDesc,
              unidade: String(getField(row, ["unidade_insumo", "un_insumo", "unidade insumo"])) || "un",
              coeficiente: num(getField(row, ["coeficiente", "coef", "indice", "índice", "produtividade"])),
              valorUnitario: rawValor === "" ? "" : num(rawValor),
            });
          }
        });
        novas.push(...Object.values(grouped));
      } else {
        let atual = null;
        rows.forEach((row) => {
          const codigo = String(getField(row, ["codigo", "código", "code"])).trim();
          const descricao = String(getField(row, ["descricao", "descrição", "servico", "serviço", "item"])).trim();
          const unidade = String(getField(row, ["unidade", "un", "unid"])).trim();
          const coefRaw = getField(row, ["coeficiente", "coef", "indice", "índice", "produtividade"]);
          if (!codigo && !descricao) return;
          if (coefRaw === "" || coefRaw === undefined || coefRaw === null) {
            atual = { id: uid(), codigo: codigo || "(sem código)", fonte: "Própria", descricao, unidade: unidade || "un", insumos: [] };
            novas.push(atual);
          } else if (atual) {
            atual.insumos.push({ id: uid(), tipo: inferTipo(descricao), descricao, unidade: unidade || "un", coeficiente: num(coefRaw), valorUnitario: "" });
          }
        });
      }

      if (novas.length === 0) {
        setImportMsg("Nenhuma composição reconhecida.");
      } else {
        const resultado = mesclarCpusImportadas(cpus, novas);
        setCpus(resultado.cpus);
        setImportMsg(
          `Importação concluída: ${resultado.adicionadas} nova(s), ${resultado.atualizadas} atualizada(s), ${resultado.semMudanca} sem mudança. Clique em Salvar para gravar na nuvem.`
        );
      }
    } catch (err) {
      setImportMsg("Erro ao ler: " + err.message);
    }
    e.target.value = "";
    setTimeout(() => setImportMsg(""), 4000);
  };

  return (
    <div>
      <div className="flex flex-wrap items-center gap-2 mb-4">
        <div className="relative flex-1 min-w-[200px]">
          <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-stone-400" />
          <input
            value={query}
            onChange={(e) => {
              setQuery(e.target.value);
              setActiveIndex(-1); // Reseta a linha ativa ao digitar
            }}
            onKeyDown={handleKeyDown} // NOVO: Gatilho para monitorar as setas do teclado
            placeholder='Buscar na biblioteca... ex: "alvenaria" "bloco"'
            className="w-full pl-8 pr-3 py-2 text-sm border border-stone-300 rounded-lg bg-white focus:outline-none focus:ring-1 focus:ring-stone-500"
          />
        </div>
        <select value={fonteFiltro} onChange={(e) => { setFonteFiltro(e.target.value); setActiveIndex(-1); }} className="px-3 py-2 text-sm border border-stone-300 rounded-lg bg-white">
          {fontes.map((f) => <option key={f}>{f}</option>)}
        </select>
        <label className="flex items-center gap-1.5 px-3 py-2 text-sm border border-stone-300 rounded-lg bg-white cursor-pointer hover:bg-stone-100">
          <Upload size={15} /> Importar/Atualizar Base
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleImport} />
        </label>
        <button onClick={() => setEditing("new")} className="flex items-center gap-1.5 px-3 py-2 text-sm bg-stone-900 text-white rounded-lg hover:bg-stone-700">
          <Plus size={15} /> Nova CPU Base
        </button>
      </div>

      {importMsg && <div className="mb-4 text-xs px-3 py-2 rounded-lg bg-amber-50 text-amber-800 border border-amber-200">{importMsg}</div>}

      <div className="space-y-2">
        {filtered.map((c, index) => (
          <div 
            key={c.id} 
            className={`border rounded-lg bg-white transition-all ${
              index === activeIndex ? "border-stone-500 ring-1 ring-stone-500 bg-stone-50/40" : "border-stone-200"
            }`}
          >
            <div className="flex items-center gap-3 px-4 py-3 cursor-pointer" onClick={() => { setExpanded({ ...expanded, [c.id]: !expanded[c.id] }); setActiveIndex(index); }}>
              {expanded[c.id] ? <ChevronDown size={16} className="text-stone-400 shrink-0" /> : <ChevronRight size={16} className="text-stone-400 shrink-0" />}
              <span className="text-[11px] font-mono px-1.5 py-0.5 bg-stone-100 rounded text-stone-500 shrink-0">{c.fonte}</span>
              <span className="text-xs font-mono text-stone-500 shrink-0">{c.codigo}</span>
              <span className={`text-sm flex-1 truncate ${index === activeIndex ? "font-medium text-stone-900" : "text-stone-800"}`}>{c.descricao}</span>
              <span className="text-xs text-stone-400 shrink-0">/{c.unidade}</span>
              <div className="flex gap-1 shrink-0" onClick={(e) => e.stopPropagation()}>
                <IconBtn onClick={() => duplicateCpu(c)} title="Duplicar"><Copy size={14} /></IconBtn>
                <IconBtn onClick={() => setEditing(c)} title="Editar"><Pencil size={14} /></IconBtn>
                {confirmingDelete === c.id ? (
                  <span className="flex items-center gap-1 text-xs">
                    <button onClick={() => removeCpu(c.id)} className="px-1.5 py-0.5 bg-red-600 text-white rounded">Sim</button>
                    <button onClick={() => setConfirmingDelete(null)} className="px-1.5 py-0.5 border border-stone-300 rounded">Não</button>
                  </span>
                ) : (
                  <IconBtn onClick={() => setConfirmingDelete(c.id)} title="Excluir"><Trash2 size={14} /></IconBtn>
                )}
              </div>
            </div>
            {expanded[c.id] && (
              <div className="px-4 pb-3 border-t border-stone-100 pt-2 bg-stone-50/50">
                <InsumoTable insumos={c.insumos} readOnly />
              </div>
            )}
          </div>
        ))}
      </div>

      {editing && <CpuEditor cpu={editing === "new" ? null : editing} onCancel={() => setEditing(null)} onSave={saveCpu} catalogMap={catalogMap} />}
    </div>
  );
}

function IconBtn({ onClick, title, children }) {
  return <button onClick={onClick} title={title} className="p-1.5 text-stone-400 hover:text-stone-700 hover:bg-stone-100 rounded">{children}</button>;
}

/* ---------------- TABELA DE INSUMOS PADRONIZADA ---------------- */
function InsumoTable({ insumos, readOnly, onChange, catalogMap, cpus = [], onUpsertPreco }) {
  const [subCpusExpandidas, setSubCpusExpandidas] = useState({});
  const setMany = (id, patch) => onChange(insumos.map((i) => (i.id === id ? { ...i, ...patch } : i)));
  const set = (id, field, value) => setMany(id, { [field]: value });
  const remove = (id) => onChange(insumos.filter((i) => i.id !== id));

  const handleDescricaoBlur = (i) => {
    if (!catalogMap) return;
    const entry = catalogMap.get(norm(i.descricao));
    if (!entry) return;
    const semValor = i.valorUnitario === "" || i.valorUnitario === null || i.valorUnitario === undefined;
    if (semValor && entry.valorUnitario !== "" && entry.valorUnitario !== null) {
      setMany(i.id, { valorUnitario: entry.valorUnitario, tipo: i.tipo || entry.tipo, unidade: i.unidade || entry.unidade });
    }
  };

  const renderSubCpuTree = (cpu, nivel = 0, visited = new Set()) => {
    if (!cpu || visited.has(cpu.id)) {
      return null;
    }

    const nextVisited = new Set(visited);
    nextVisited.add(cpu.id);

    return (
      <div className={`border rounded-md overflow-hidden bg-white ${nivel === 0 ? "border-amber-100" : "border-stone-200 mt-1"}`}>
        <div className={`grid grid-cols-12 gap-2 px-3 py-1.5 text-[10px] font-semibold uppercase ${nivel === 0 ? "bg-amber-50 text-amber-800" : "bg-stone-50 text-stone-600"}`}>
          <span className="col-span-6">{cpu.codigo} - {cpu.descricao}</span>
          <span className="col-span-1 text-center">Un.</span>
          <span className="col-span-2 text-right">Coef. interno</span>
          <span className="col-span-1.5 text-right">Valor unit.</span>
          <span className="col-span-1.5 text-right">Subtotal</span>
        </div>
        {(cpu.insumos || []).map((subInsumo) => {
          const rowKey = `${cpu.id}-${subInsumo.id || subInsumo.codigo || subInsumo.descricao}-${nivel}`;
          const subSubCpu = findSubCpu(subInsumo, cpus);
          const estaAberta = !!subCpusExpandidas[rowKey];
          const subValor = insumoValorUnitario(subInsumo, cpus, catalogMap, nextVisited);
          const subTotal = num(subInsumo.coeficiente) * subValor;
          const podeEditarValor = !readOnly && !subSubCpu && onUpsertPreco;

          return (
            <React.Fragment key={rowKey}>
              <div className="grid grid-cols-12 gap-2 px-3 py-1.5 text-[11px] border-t border-stone-100 items-center">
                <span className="col-span-6 truncate text-stone-700 flex items-center gap-1">
                  {subSubCpu && (
                    <button
                      type="button"
                      onClick={() => setSubCpusExpandidas((prev) => ({ ...prev, [rowKey]: !estaAberta }))}
                      className="text-stone-400 hover:text-stone-800"
                      title={estaAberta ? "Recolher sub-CPU interna" : "Expandir sub-CPU interna"}
                    >
                      {estaAberta ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                    </button>
                  )}
                  <span className={`mr-1 text-[9px] px-1 py-0.5 rounded ${subSubCpu ? "bg-amber-100 text-amber-800" : "bg-stone-100 text-stone-500"}`}>
                    {subSubCpu ? "CPU" : subInsumo.tipo}
                  </span>
                  <span className="truncate">{subInsumo.codigo ? `${subInsumo.codigo} - ` : ""}{subInsumo.descricao}</span>
                </span>
                <span className="col-span-1 text-center font-mono text-stone-400">{subInsumo.unidade || "un"}</span>
                <span className="col-span-2 text-right font-mono text-stone-600">{fmt(subInsumo.coeficiente)}</span>
                <span className="col-span-1.5 text-right font-mono text-stone-500">
                  {podeEditarValor ? (
                    <input
                      type="number"
                      step="any"
                      value={subValor || ""}
                      onChange={(e) =>
                        onUpsertPreco(
                          subInsumo.descricao,
                          subInsumo.tipo || "MAT",
                          subInsumo.unidade || "un",
                          e.target.value === "" ? "" : num(e.target.value)
                        )
                      }
                      className="w-24 border border-amber-200 rounded px-1 py-0.5 text-right font-mono bg-white"
                      placeholder="0,00"
                    />
                  ) : (
                    <>R$ {fmt(subValor)}</>
                  )}
                </span>
                <span className="col-span-1.5 text-right font-mono font-medium text-stone-700">R$ {fmt(subTotal)}</span>
              </div>
              {subSubCpu && estaAberta && (
                <div className="border-t border-stone-100 bg-stone-50/30 pl-6 pr-2 py-2">
                  {nextVisited.has(subSubCpu.id) ? (
                    <div className="text-[11px] text-amber-700 px-3 py-2">Ciclo de sub-CPU detectado; expansão interrompida.</div>
                  ) : (
                    renderSubCpuTree(subSubCpu, nivel + 1, nextVisited)
                  )}
                </div>
              )}
            </React.Fragment>
          );
        })}
      </div>
    );
  };

  return (
    <table className="w-full text-xs">
      <thead>
        <tr className="text-stone-400 text-left">
          <th className="font-normal py-1 pr-2 w-24">Tipo</th>
          <th className="font-normal py-1 pr-2">Insumo</th>
          <th className="font-normal py-1 pr-2 w-16">Un.</th>
          <th className="font-normal py-1 pr-2 w-24 text-right">Coeficiente</th>
          <th className="font-normal py-1 pr-2 w-28 text-right">Valor Unit. (R$)</th>
          <th className="font-normal py-1 pr-2 w-24 text-right">Subtotal</th>
          {!readOnly && <th className="w-7"></th>}
        </tr>
      </thead>
      <tbody>
        {insumos.map((i) => {
          const valorEfetivo = insumoValorUnitario(i, cpus, catalogMap);
          const subCpu = findSubCpu(i, cpus);
          const subCpuAberta = !!subCpusExpandidas[i.id];
          return (
          <React.Fragment key={i.id}>
          <tr className={`border-t border-stone-100 ${subCpu ? "bg-amber-50/40" : ""}`}>
            <td className="py-1 pr-2">
              {readOnly ? (
                <span className={`text-[10px] px-1 py-0.5 rounded font-medium ${subCpu ? "bg-amber-100 text-amber-800" : "bg-stone-100 text-stone-600"}`}>
                  {subCpu ? "CPU" : i.tipo}
                </span>
              ) : subCpu ? (
                <span className="text-[10px] px-1 py-0.5 rounded font-medium bg-amber-100 text-amber-800">CPU</span>
              ) : (
                <select value={i.tipo || "MAT"} onChange={(e) => set(i.id, "tipo", e.target.value)} className="w-full border border-stone-200 rounded p-0.5 bg-white">
                  {TIPOS.map((t) => <option key={t.v} value={t.v}>{t.label}</option>)}
                </select>
              )}
            </td>
            <td className="py-1 pr-2">
              {readOnly ? (
                <span className="text-stone-700 flex items-center gap-1.5">
                  {subCpu && (
                    <button
                      type="button"
                      onClick={() => setSubCpusExpandidas((prev) => ({ ...prev, [i.id]: !subCpuAberta }))}
                      className="text-stone-400 hover:text-stone-800"
                      title={subCpuAberta ? "Recolher insumos da sub-CPU" : "Expandir insumos da sub-CPU"}
                    >
                      {subCpuAberta ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                    </button>
                  )}
                  <span className="truncate">
                    {i.codigo ? `${i.codigo} - ` : ""}{i.descricao}
                    {subCpu && <span className="ml-2 text-[10px] text-amber-700 font-semibold">sub-CPU</span>}
                  </span>
                </span>
              ) : (
                <div className="flex items-center gap-1.5">
                  {subCpu && (
                    <button
                      type="button"
                      onClick={() => setSubCpusExpandidas((prev) => ({ ...prev, [i.id]: !subCpuAberta }))}
                      className="text-stone-400 hover:text-stone-800"
                      title={subCpuAberta ? "Recolher insumos da sub-CPU" : "Expandir insumos da sub-CPU"}
                    >
                      {subCpuAberta ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                    </button>
                  )}
                  <input
                    value={i.descricao || ""}
                    onChange={(e) => set(i.id, "descricao", e.target.value)}
                    onBlur={() => handleDescricaoBlur(i)}
                    list="insumos-catalogo"
                    className={`w-full border rounded px-1 py-0.5 ${subCpu ? "border-amber-200 bg-amber-50/40" : "border-stone-200"}`}
                  />
                  {subCpu && <span className="text-[10px] text-amber-700 font-semibold whitespace-nowrap">sub-CPU</span>}
                </div>
              )}
            </td>
            <td className="py-1 pr-2">
              {readOnly ? (
                <span className="text-stone-500">{i.unidade}</span>
              ) : (
                <input value={i.unidade || ""} onChange={(e) => set(i.id, "unidade", e.target.value)} className="w-full border border-stone-200 rounded px-1 py-0.5" />
              )}
            </td>
            <td className="py-1 pr-2 text-right">
              {readOnly ? (
                <span className="font-mono">{i.coeficiente}</span>
              ) : (
                <input type="number" step="any" value={i.coeficiente ?? ""} onChange={(e) => set(i.id, "coeficiente", e.target.value)} className="w-20 border border-stone-200 rounded px-1 py-0.5 text-right font-mono" />
              )}
            </td>
            <td className="py-1 pr-2 text-right">
              {readOnly ? (
                <span className="font-mono text-stone-600">{valorEfetivo ? `R$ ${fmt(valorEfetivo)}` : "-"}</span>
              ) : (
                <input type="number" step="any" value={valorEfetivo || ""} onChange={(e) => set(i.id, "valorUnitario", e.target.value)} placeholder="0,00" className="w-24 border border-stone-200 rounded px-1 py-0.5 text-right font-mono" />
              )}
            </td>
            <td className="py-1 pr-2 text-right font-mono text-stone-600">
              R$ {fmt(num(i.coeficiente) * valorEfetivo)}
            </td>
            {!readOnly && (
              <td className="py-1 text-center">
                <button onClick={() => remove(i.id)} className="text-stone-300 hover:text-red-500"><X size={13} /></button>
              </td>
            )}
          </tr>
          {subCpu && subCpuAberta && (
            <tr className="bg-amber-50/20 border-t border-amber-100">
              <td colSpan={readOnly ? 6 : 7} className="py-2 pl-10 pr-2">
                {renderSubCpuTree(subCpu)}
              </td>
            </tr>
          )}
          </React.Fragment>
          );
        })}
        {!readOnly && (
          <tr>
            <td colSpan="7" className="py-2">
              <button type="button" onClick={() => onChange([...insumos, { id: uid(), tipo: "MAT", descricao: "", unidade: "un", coeficiente: 1, valorUnitario: "" }])} className="text-stone-500 hover:text-stone-900 font-medium flex items-center gap-1">
                <Plus size={12} /> Adicionar Insumo
              </button>
            </td>
          </tr>
        )}
      </tbody>
    </table>
  );
}

/* ---------------- ABA BANCO DE PREÇOS (PRECOS) ---------------- */
function PrecosTab({ catalog, onUpsert, onRemove, onApplyToCpus, onApplyAllToCpus }) {
  const [editing, setEditing] = useState(null);
  const [query, setQuery] = useState("");

  const exportarXls = () => {
    const wb = XLSX.utils.book_new();
    const rows = [
      ["Banco de Preços - Catálogo de Insumos"],
      [],
      ["Tipo", "Descrição", "Unidade", "Valor Unitário (R$)", "Ocorrências na Planilha"],
      ...catalog.map((c) => [c.tipo, c.descricao, c.unidade, c.valorUnitario !== "" ? c.valorUnitario : "", c.ocorrencias]),
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 12 }, { wch: 45 }, { wch: 10 }, { wch: 20 }, { wch: 22 }];
    XLSX.utils.book_append_sheet(wb, ws, "Banco de Preços");
    XLSX.writeFile(wb, "banco_de_precos.xlsx");
  };

  const filtered = catalog.filter((c) => {
    // Divide o texto digitado por espaços e remove itens vazios
    const searchTerms = norm(query).split(/\s+/).filter(Boolean);
    const targetText = norm(c.descricao);
    
    // Verifica se TODAS as palavras buscadas estão presentes na descrição do insumo
    return searchTerms.every((term) => targetText.includes(term));
  });

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4">
      <div className="flex flex-wrap items-center justify-between gap-4 mb-4">
        <div className="relative w-72">
          <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-stone-400" />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Filtrar catálogo de preços..." className="w-full pl-8 pr-3 py-2 text-sm border border-stone-300 rounded-lg" />
        </div>
        <div className="flex gap-2">
          <button onClick={onApplyAllToCpus} className="flex items-center gap-1.5 px-3 py-1.5 text-xs border border-stone-300 rounded-lg font-medium bg-stone-50 hover:bg-stone-100 text-stone-700">
            <RefreshCw size={13} /> Sincronizar Tudo na Planilha de Custos
          </button>
          <button onClick={exportarXls} className="flex items-center gap-1.5 px-3 py-1.5 text-xs border border-stone-300 rounded-lg font-medium bg-white hover:bg-stone-50 text-stone-700">
            <Download size={13} /> Exportar .xlsx
          </button>
          <button onClick={() => setEditing({ id: null, descricao: "", tipo: "MAT", unidade: "un", valorUnitario: "" })} className="flex items-center gap-1.5 px-3 py-1.5 text-xs bg-stone-900 text-white rounded-lg font-medium hover:bg-stone-700">
            <Plus size={13} /> Novo Insumo Manual
          </button>
        </div>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-left text-xs border-collapse">
          <thead>
            <tr className="border-b border-stone-200 text-stone-400 font-normal">
              <th className="py-2 pr-3 w-28">Tipo</th>
              <th className="py-2 pr-3">Descrição Única do Insumo</th>
              <th className="py-2 pr-3 w-20">Un.</th>
              <th className="py-2 pr-3 w-32 text-right">Preço Padrão (R$)</th>
              <th className="py-2 pr-3 w-28 text-center">Na Planilha Ativa</th>
              <th className="py-2 w-24 text-center">Ações</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((c) => (
              <tr key={c.key} className="border-b border-stone-100 hover:bg-stone-50/50">
                <td className="py-2 pr-3">
                  <span className="px-1.5 py-0.5 bg-stone-100 text-stone-600 rounded font-medium text-[10px]">{c.tipo}</span>
                </td>
                <td className="py-2 pr-3 font-medium text-stone-800">
                  {c.descricao}
                  {c.divergente && (
                    <span className="ml-2 text-[10px] bg-amber-50 text-amber-700 border border-amber-200 px-1.5 py-0.2 rounded inline-flex items-center gap-1">
                      <AlertTriangle size={10} /> Preço Divergente
                    </span>
                  )}
                </td>
                <td className="py-2 pr-3 text-stone-500 font-mono">{c.unidade}</td>
                <td className="py-2 pr-3 text-right font-mono font-medium text-stone-900">
                  {c.valorUnitario !== "" ? `R$ ${fmt(c.valorUnitario)}` : <span className="text-stone-300">Não definido</span>}
                </td>
                <td className="py-2 pr-3 text-center text-stone-500">{c.ocorrencias} item(ns)</td>
                <td className="py-2 text-center flex justify-center gap-1">
                  <button onClick={() => setEditing(c)} className="p-1 border border-stone-200 rounded text-stone-600 hover:bg-stone-100" title="Editar Preço">
                    <Pencil size={12} />
                  </button>
                  {c.valorUnitario !== "" && c.ocorrencias > 0 && (
                    <button onClick={() => onApplyToCpus(c.descricao, c.valorUnitario)} className="p-1 border border-stone-200 rounded bg-stone-50 text-stone-700 hover:bg-stone-100" title="Forçar este valor nos Custos desta obra">
                      <Check size={12} />
                    </button>
                  )}
                  <button onClick={() => onRemove(c.descricao)} className="p-1 border border-stone-200 rounded text-stone-400 hover:text-red-600" title="Remover Referência">
                    <X size={12} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {editing && (
        <div className="fixed inset-0 bg-stone-900/40 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-white border border-stone-200 rounded-xl max-w-md w-full p-5 shadow-lg">
            <h3 className="font-semibold text-sm text-stone-900 mb-4">{editing.key ? "Editar Insumo do Catálogo" : "Novo Insumo no Banco"}</h3>
            <div className="space-y-3 text-xs">
              <div>
                <label className="block text-stone-500 mb-1">Descrição</label>
                <input disabled={!!editing.key} value={editing.descricao} onChange={(e) => setEditing({ ...editing, descricao: e.target.value })} className="w-full border border-stone-300 rounded-lg px-3 py-2 disabled:bg-stone-50" />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-stone-500 mb-1">Tipo</label>
                  <select value={editing.tipo} onChange={(e) => setEditing({ ...editing, tipo: e.target.value })} className="w-full border border-stone-300 rounded-lg px-2 py-2 bg-white">
                    {TIPOS.map((t) => <option key={t.v} value={t.v}>{t.label}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-stone-500 mb-1">Unidade</label>
                  <input value={editing.unidade} onChange={(e) => setEditing({ ...editing, unidade: e.target.value })} className="w-full border border-stone-300 rounded-lg px-3 py-2" />
                </div>
              </div>
              <div>
                <label className="block text-stone-500 mb-1">Valor Unitário Homologado (R$)</label>
                <input type="number" step="any" value={editing.valorUnitario} onChange={(e) => setEditing({ ...editing, valorUnitario: e.target.value })} placeholder="0,00" className="w-full border border-stone-300 rounded-lg px-3 py-2 font-mono text-sm" />
              </div>
            </div>
            <div className="mt-5 flex justify-end gap-2 text-xs">
              <button onClick={() => setEditing(null)} className="px-3 py-2 border border-stone-300 rounded-lg">Cancelar</button>
              <button
                onClick={() => {
                  onUpsert(editing.descricao, editing.tipo, editing.unidade, editing.valorUnitario === "" ? "" : num(editing.valorUnitario));
                  if (editing.key && editing.valorUnitario !== "") {
                    onApplyToCpus(editing.descricao, num(editing.valorUnitario));
                  }
                  setEditing(null);
                }}
                className="px-3 py-2 bg-stone-900 text-white rounded-lg hover:bg-stone-700"
              >
                Salvar e Replicar nos Custos
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------- EDITOR DE CPUS INDIVIDUAIS ---------------- */
function CpuEditor({ cpu, onCancel, onSave, catalogMap }) {
  const [codigo, setCodigo] = useState(cpu?.codigo || "");
  const [fonte, setFonte] = useState(cpu?.fonte || "Própria");
  const [descricao, setDescricao] = useState(cpu?.descricao || "");
  const [unidade, setUnidade] = useState(cpu?.unidade || "m²");
  const [insumos, setInsumos] = useState(cpu?.insumos ? JSON.parse(JSON.stringify(cpu.insumos)) : []);

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!descricao.trim()) return;
    onSave({ id: cpu?.id || uid(), codigo, fonte, descricao, unidade, insumos });
  };

  return (
    <div className="fixed inset-0 bg-stone-900/40 backdrop-blur-xs flex items-center justify-center p-4 z-40 overflow-y-auto">
      <form onSubmit={handleSubmit} className="bg-white border border-stone-200 rounded-xl max-w-2xl w-full p-5 shadow-lg my-8">
        <h3 className="font-semibold text-sm mb-4">{cpu ? "Editar Composição da Base" : "Nova Composição Técnica"}</h3>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs mb-4">
          <div>
            <label className="block text-stone-500 mb-1">Tabela / Fonte</label>
            <input value={fonte} onChange={(e) => setFonte(e.target.value)} placeholder="Ex: SINAPI, SUDECAP" className="w-full border border-stone-300 rounded-lg px-3 py-2" />
          </div>
          <div>
            <label className="block text-stone-500 mb-1">Código Identificador</label>
            <input value={codigo} onChange={(e) => setCodigo(e.target.value)} placeholder="Ex: 12.34.56" className="w-full border border-stone-300 rounded-lg px-3 py-2" />
          </div>
          <div>
            <label className="block text-stone-500 mb-1">Unidade Principal</label>
            <input value={unidade} onChange={(e) => setUnidade(e.target.value)} className="w-full border border-stone-300 rounded-lg px-3 py-2" />
          </div>
        </div>
        <div className="text-xs mb-4">
          <label className="block text-stone-500 mb-1">Descrição Técnica da Composição</label>
          <input value={descricao} onChange={(e) => setDescricao(e.target.value)} placeholder="Ex: Concreto armado fck=25mpa..." className="w-full border border-stone-300 rounded-lg px-3 py-2" />
        </div>

        <div className="border-t border-stone-200 pt-3">
          <h4 className="text-xs font-semibold text-stone-700 mb-2">Estrutura de Insumos da CPU</h4>
          <InsumoTable insumos={insumos} onChange={setInsumos} catalogMap={catalogMap} />
        </div>

        <div className="mt-6 pt-3 border-t border-stone-200 flex justify-end gap-2 text-xs">
          <button type="button" onClick={onCancel} className="px-4 py-2 border border-stone-300 rounded-lg">Cancelar</button>
          <button type="submit" className="px-4 py-2 bg-stone-900 text-white rounded-lg hover:bg-stone-700">Salvar na Biblioteca</button>
        </div>
      </form>
    </div>
  );
}

/* ---------------- PLANILHA DE ORÇAMENTO / CUSTO ---------------- */
function Orcamento({ etapas, setEtapas, cpus, grandTotal, catalogMap, onUpsertPreco }) {
  const [buscasPorEtapa, setBuscasPorEtapa] = useState({}); // Controla a busca de cada etapa individualmente
  const [editingEtapaId, setEditingEtapaId] = useState(null);
  const [editingEtapaNome, setEditingEtapaNome] = useState("");
  
  // Controla o índice do item selecionado via teclado para cada etapa
  const [activeIndices, setActiveIndices] = useState({}); 

  // NOVO: Controla quais itens da etapa estão expandidos (mostrando insumos)
  const [itensExpandidos, setItensExpandidos] = useState({});
  const [etapasRecolhidas, setEtapasRecolhidas] = useState({});

  const cpuSearchIndex = useMemo(
    () =>
      cpus.map((c) => ({
        cpu: c,
        haystack: norm(`${c.codigo || ""} ${c.descricao || ""}`),
      })),
    [cpus]
  );

  const adicionarEtapa = () => {
    setEtapas([...etapas, { id: uid(), nome: `Nova Etapa ${etapas.length + 1}`, itens: [] }]);
  };

  const removerEtapa = (id) => {
    if (etapas.length <= 1) return;
    setEtapas(etapas.filter((e) => e.id !== id));
  };

  const salvarNomeEtapa = (id) => {
    setEtapas(etapas.map((e) => (e.id === id ? { ...e, nome: editingEtapaNome } : e)));
    setEditingEtapaId(null);
  };

  const lancarCpuNaEtapa = (etapaId, cpu) => {
    const insumosAjustados = applyCatalogToInsumos(cpu.insumos, catalogMap);
    setEtapas(
      etapas.map((e) => {
        if (e.id !== etapaId) return e;
        return {
          ...e,
          itens: [
            ...e.itens,
            {
              id: uid(),
              cpuId: cpu.id,
              codigo: cpu.codigo,
              servico: cpu.descricao,
              unidade: cpu.unidade,
              quantidade: 1,
              insumos: insumosAjustados,
            },
          ],
        };
      })
    );
  };

  const mudarQuantidadeItem = (etapaId, itemId, Qtd) => {
    setEtapas(
      etapas.map((e) => {
        if (e.id !== etapaId) return e;
        return {
          ...e,
          itens: e.itens.map((it) => (it.id === itemId ? { ...it, quantidade: Qtd } : it)),
        };
      })
    );
  };

  const mudarInsumosDoItem = (etapaId, itemId, novosInsumos) => {
    setEtapas(
      etapas.map((e) => {
        if (e.id !== etapaId) return e;
        return {
          ...e,
          itens: e.itens.map((it) => (it.id === itemId ? { ...it, insumos: novosInsumos } : it)),
        };
      })
    );
  };

  const removerItemDaEtapa = (etapaId, itemId) => {
    setEtapas(
      etapas.map((e) => {
        if (e.id !== etapaId) return e;
        return { ...e, itens: e.itens.filter((it) => it.id !== itemId) };
      })
    );
  };

  const toggleExpandirItem = (itemId) => {
    setItensExpandidos((prev) => ({ ...prev, [itemId]: !prev[itemId] }));
  };

  const toggleRecolherEtapa = (etapaId) => {
    setEtapasRecolhidas((prev) => ({ ...prev, [etapaId]: !prev[etapaId] }));
  };

  const obterCpusFiltradas = (textoBusca) => {
    if (!textoBusca || !textoBusca.trim()) return [];
    const searchTerms = norm(textoBusca).split(/\s+/).filter(Boolean);
    const result = [];
    for (const item of cpuSearchIndex) {
      if (searchTerms.every((term) => item.haystack.includes(term))) {
        result.push(item.cpu);
        if (result.length >= 10) break;
      }
    }
    return result;
  };

  const handleKeyDown = (evt, etapaId, listaCpus) => {
    if (listaCpus.length === 0) return;
    
    const currentIndex = activeIndices[etapaId] !== undefined ? activeIndices[etapaId] : -1;

    if (evt.key === "ArrowDown") {
      evt.preventDefault();
      const nextIndex = (currentIndex + 1) % listaCpus.length;
      setActiveIndices({ ...activeIndices, [etapaId]: nextIndex });
    } else if (evt.key === "ArrowUp") {
      evt.preventDefault();
      const prevIndex = (currentIndex - 1 + listaCpus.length) % listaCpus.length;
      setActiveIndices({ ...activeIndices, [etapaId]: prevIndex });
    } else if (evt.key === "Enter") {
      if (currentIndex >= 0 && currentIndex < listaCpus.length) {
        evt.preventDefault();
        lancarCpuNaEtapa(etapaId, listaCpus[currentIndex]);
        setBuscasPorEtapa({ ...buscasPorEtapa, [etapaId]: "" });
        setActiveIndices({ ...activeIndices, [etapaId]: -1 });
      }
    } else if (evt.key === "Escape") {
      setBuscasPorEtapa({ ...buscasPorEtapa, [etapaId]: "" });
      setActiveIndices({ ...activeIndices, [etapaId]: -1 });
    }
  };
  
  return (
    <div className="space-y-4">
      {/* Topo da aba: apenas o botão de Adicionar Etapa */}
      <div className="flex justify-end items-center bg-white border border-stone-200 rounded-lg p-3">
        <button onClick={adicionarEtapa} className="flex items-center gap-1 px-3 py-1.5 text-xs bg-stone-900 text-white rounded-lg hover:bg-stone-700">
          <Plus size={14} /> Adicionar Nova Etapa
        </button>
      </div>

      {/* Listagem das Etapas */}
      <div className="space-y-4">
        {etapas.map((e) => {
          const termoBuscaEtapa = buscasPorEtapa[e.id] || "";
          const filtradasParaEstaEtapa = obterCpusFiltradas(termoBuscaEtapa);
          const activeIndex = activeIndices[e.id] !== undefined ? activeIndices[e.id] : -1;
          const etapaRecolhida = !!etapasRecolhidas[e.id];
          const totalEtapa = (e.itens || []).reduce(
            (s, it) => s + num(it.quantidade) * cpuValorUnit(it.insumos, cpus, catalogMap),
            0
          );

          return (
            <div key={e.id} className="bg-white border border-stone-200 rounded-lg overflow-visible">
              <div
                onClick={() => toggleRecolherEtapa(e.id)}
                className="bg-stone-200 px-4 py-2.5 flex justify-between items-center gap-3 border-b border-stone-300 cursor-pointer select-none hover:bg-stone-300"
              >
                <div className="min-w-0 flex items-center gap-2">
                  {etapaRecolhida ? (
                    <ChevronRight size={15} className="text-stone-400 shrink-0" />
                  ) : (
                    <ChevronDown size={15} className="text-stone-400 shrink-0" />
                  )}
                  {editingEtapaId === e.id ? (
                    <div className="flex items-center gap-2" onClick={(evt) => evt.stopPropagation()}>
                      <input value={editingEtapaNome} onChange={(e) => setEditingEtapaNome(e.target.value)} className="border border-stone-300 text-xs rounded px-2 py-1 bg-white" />
                      <button onClick={() => salvarNomeEtapa(e.id)} className="text-stone-800 font-bold text-xs">Salvar</button>
                    </div>
                  ) : (
                    <h3 className="font-medium text-sm text-stone-800 flex items-center gap-2 min-w-0">
                      <span className="truncate">{e.nome}</span>
                      <button
                        onClick={(evt) => {
                          evt.stopPropagation();
                          setEditingEtapaId(e.id);
                          setEditingEtapaNome(e.nome);
                        }}
                        className="text-stone-400 hover:text-stone-700 shrink-0"
                      >
                        <Pencil size={12} />
                      </button>
                    </h3>
                  )}
                </div>
                <div className="flex items-center gap-4 shrink-0">
                  <div className="text-right">
                    <span className="block text-[10px] text-stone-400 font-mono uppercase">Total da etapa</span>
                    <span className="font-semibold text-stone-900 font-mono text-sm">R$ {fmt(totalEtapa)}</span>
                  </div>
                  {etapas.length > 1 && (
                    <button
                      onClick={(evt) => {
                        evt.stopPropagation();
                        removerEtapa(e.id);
                      }}
                      className="text-stone-400 hover:text-red-500"
                    >
                      <Trash2 size={14} />
                    </button>
                  )}
                </div>
              </div>

              {/* Contêiner expande dinamicamente ao digitar na busca */}
              {!etapaRecolhida && (
              <div className={`p-4 space-y-3 transition-all ${termoBuscaEtapa.trim() ? "min-h-[400px]" : "min-h-0"}`}>
                {/* Campo de busca exclusivo DESTA ETAPA - LARGURA TOTAL */}
                <div className="relative w-full mb-3">
                  <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-stone-400" />
                  <input 
                    value={termoBuscaEtapa} 
                    onChange={(evt) => {
                      setBuscasPorEtapa({ ...buscasPorEtapa, [e.id]: evt.target.value });
                      setActiveIndices({ ...activeIndices, [e.id]: -1 });
                    }}
                    onKeyDown={(evt) => handleKeyDown(evt, e.id, filtradasParaEstaEtapa)}
                    placeholder="Pesquisar CPU para lançar NESTA etapa..." 
                    className="w-full pl-8 pr-3 py-1.5 text-xs border border-stone-300 rounded-lg bg-stone-50/40 focus:bg-white" 
                  />
                  
                  {termoBuscaEtapa.trim() && (
                    <div className="absolute left-0 right-0 top-full bg-white border border-stone-200 rounded-b-lg shadow-xl mt-1 z-50 max-h-[350px] overflow-y-auto text-xs">
                      {filtradasParaEstaEtapa.length === 0 && <p className="p-3 text-stone-400">Nenhuma composição encontrada.</p>}
                      {filtradasParaEstaEtapa.map((c, index) => (
                        <div 
                          key={c.id} 
                          className={`p-2 border-b border-stone-100 last:border-0 cursor-pointer flex justify-between items-center transition-colors ${
                            index === activeIndex ? "bg-stone-100 font-medium" : "hover:bg-stone-50"
                          }`} 
                          onClick={() => {
                            lancarCpuNaEtapa(e.id, c);
                            setBuscasPorEtapa({ ...buscasPorEtapa, [e.id]: "" });
                            setActiveIndices({ ...activeIndices, [e.id]: -1 });
                          }}
                        >
                          <div className="flex-1 min-w-0 pr-2">
                            <span className="font-mono text-[10px] text-stone-400 block">{c.codigo}</span>
                            <p className="truncate text-stone-800">{c.descricao}</p>
                          </div>
                          <span className={`text-[10px] px-2 py-0.5 rounded shrink-0 transition-colors ${
                            index === activeIndex ? "bg-stone-900 text-white" : "bg-stone-200 text-stone-700"
                          }`}>
                            Lançar
                          </span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                {/* Exibição dos itens da Etapa */}
                {e.itens.length === 0 && !termoBuscaEtapa.trim() && (
                  <p className="text-xs text-stone-400 italic pt-1">Nenhuma CPU lançada nesta etapa.</p>
                )}
                {e.itens.map((it) => {
                  const estaExpandido = !!itensExpandidos[it.id]; // Por padrão, undefined avalia como falso (recolhido)

                  return (
                    <div key={it.id} className="border border-stone-100 rounded-lg p-3 bg-stone-50/30">
                      {/* Cabeçalho do item - Clicável para expandir/recolher */}
                      <div className="flex flex-wrap items-center justify-between gap-3 mb-1 pb-1">
                        <div 
                          className="flex-1 min-w-0 flex items-center gap-2 cursor-pointer select-none"
                          onClick={() => toggleExpandirItem(it.id)}
                          title="Clique para alternar entre nome principal e composição completa"
                        >
                          {estaExpandido ? (
                            <ChevronDown size={14} className="text-stone-400 shrink-0" />
                          ) : (
                            <ChevronRight size={14} className="text-stone-400 shrink-0" />
                          )}
                          <div className="min-w-0">
                            <span className="font-mono text-[10px] text-stone-400">{it.codigo}</span>
                            <h4 className="text-xs font-semibold text-stone-800 truncate">{it.servico}</h4>
                          </div>
                        </div>
                        
                        <div className="flex items-center gap-3 text-xs">
                          <div className="flex items-center gap-1.5">
                            <span className="text-stone-400">Qtd:</span>
                            <input type="number" step="any" value={it.quantidade} onChange={(evt) => mudarQuantidadeItem(e.id, it.id, evt.target.value)} className="w-16 border border-stone-200 rounded px-1.5 py-0.5 text-right font-mono bg-white" />
                            <span className="text-stone-500 font-medium">/{it.unidade}</span>
                          </div>
                          <div className="text-right">
                            <span className="text-[10px] block text-stone-400 font-mono">Unit: R$ {fmt(cpuValorUnit(it.insumos, cpus, catalogMap))}</span>
                            <span className="font-semibold text-stone-900 font-mono">Total: R$ {fmt(num(it.quantidade) * cpuValorUnit(it.insumos, cpus, catalogMap))}</span>
                          </div>
                          <button onClick={() => removerItemDaEtapa(e.id, it.id)} className="text-stone-300 hover:text-red-500 ml-2"><Trash2 size={14} /></button>
                        </div>
                      </div>

                      {/* Exibe a tabela de insumos apenas se o usuário expandir o item */}
                      {estaExpandido && (
                        <div className="mt-2 pt-2 border-t border-stone-100 transition-all">
                          <InsumoTable insumos={it.insumos} onChange={(novos) => mudarInsumosDoItem(e.id, it.id, novos)} catalogMap={catalogMap} cpus={cpus} onUpsertPreco={onUpsertPreco} />
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------------- ABA PLANILHA DE BDI ---------------- */
function BdiTab({ bdi, setBdi, bdiCalc, grandTotal }) {
  const faturamentoDireto = !!bdi.faturamentoDireto;

  // Inicializa taxas de materiais se não existirem
  const bdiMats = bdi.materiais || {
    admCentral: 0,
    contabilidade: 0,
    contingenciamento: 0,
    custoFinanceiro: 0,
    lucro: 0,
    dasAnexoIV: 0,
    art: 0
  };

  const handleGeralChange = (campo, valor) => {
    setBdi(prev => ({ ...prev, [campo]: valor }));
  };

  const handleMatChange = (campo, valor) => {
    setBdi(prev => ({
      ...prev,
      materiais: { ...bdiMats, [campo]: valor }
    }));
  };

  // Função auxiliar para calcular taxas somadas ou BDI para o painel resumo
  const calcularFatorQualquer = (t) => {
    const ac = num(t.admCentral);
    const c = num(t.contabilidade);
    const co = num(t.contingenciamento);
    const cf = num(t.custoFinanceiro);
    const l = num(t.lucro);
    const das = num(t.dasAnexoIV);
    const art = num(t.art);

    const pv = das + art;
    const numerador = (1 + ac) * (1 + c) * (1 + co) * (1 + cf) * (1 + l);
    const denominador = 1 - pv;
    return denominador <= 0 ? 1 : numerador / denominador;
  };

  const bdiGeralRate = calcularFatorQualquer(bdi) - 1;
  const bdiMatRate = faturamentoDireto ? (calcularFatorQualquer(bdiMats) - 1) : bdiGeralRate;

  return (
    <div className="space-y-6">
      {/* Barra de controle superior */}
      <div className="bg-white border border-stone-200 rounded-lg p-4 flex justify-between items-center flex-wrap gap-3">
        <div>
          <h3 className="font-semibold text-sm text-stone-800">Opções do Regime de Faturamento</h3>
          <p className="text-xs text-stone-400">Ative o BDI diferenciado se houver materiais faturados direto pelo fornecedor.</p>
        </div>
        <label className="flex items-center gap-2 bg-stone-50 border border-stone-200 px-3 py-1.5 rounded-md cursor-pointer select-none hover:bg-stone-100 transition-colors text-xs font-semibold text-stone-700">
          <input
            type="checkbox"
            checked={faturamentoDireto}
            onChange={(e) => setBdi(prev => ({ ...prev, faturamentoDireto: e.target.checked }))}
            className="w-4 h-4 accent-stone-900 rounded"
          />
          Habilitar Faturamento Direto (BDI Diferenciado para Materiais)
        </label>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className={`bg-white border border-stone-200 rounded-lg p-4 text-xs space-y-4 ${faturamentoDireto ? "lg:col-span-2" : "lg:col-span-2"}`}>
          <h3 className="font-semibold text-sm text-stone-800 border-b border-stone-100 pb-2">Composição Analítica do BDI</h3>
          
          <div className={`grid grid-cols-1 gap-6 ${faturamentoDireto ? "sm:grid-cols-2" : "sm:grid-cols-2"}`}>
            {/* GRUPO 1: BDI GERAL */}
            <div className="space-y-4">
              <h4 className="font-bold text-stone-700 uppercase text-[10px] bg-stone-100 px-2 py-1 rounded tracking-wide">
                {faturamentoDireto ? "1. Taxas Gerais (Serviços e MO)" : "Taxas Gerais / Padrão"}
              </h4>
              
              <div className="space-y-3">
                <h5 className="font-medium text-stone-400 uppercase text-[9px]">Administração e Riscos</h5>
                <BdiInput label="Administração Central" value={bdi.admCentral} onChange={(v) => handleGeralChange("admCentral", v)} />
                <BdiInput label="Contabilidade / Seguros" value={bdi.contabilidade} onChange={(v) => handleGeralChange("contabilidade", v)} />
                <BdiInput label="Contingenciamento" value={bdi.contingenciamento} onChange={(v) => handleGeralChange("contingenciamento", v)} />
                <BdiInput label="Custo Financeiro" value={bdi.custoFinanceiro} onChange={(v) => handleGeralChange("custoFinanceiro", v)} />
                
                <h5 className="font-medium text-stone-400 uppercase text-[9px] pt-1">Margem e Impostos</h5>
                <BdiInput label="Lucro Real de Venda" value={bdi.lucro} onChange={(v) => handleGeralChange("lucro", v)} />
                <BdiInput label="DAS / Tributos (Anexo IV)" value={bdi.dasAnexoIV} onChange={(v) => handleGeralChange("dasAnexoIV", v)} />
                <BdiInput label="ART / Encargos Contrato" value={bdi.art} onChange={(v) => handleGeralChange("art", v)} />
              </div>

              <div className="pt-2 border-t border-stone-100 flex justify-between items-center text-[11px] font-bold text-stone-700">
                <span>Taxa BDI Geral:</span>
                <span className="font-mono bg-stone-100 text-stone-800 px-1.5 py-0.5 rounded">{fmt(bdiGeralRate * 100)}%</span>
              </div>
            </div>

            {/* GRUPO 2: BDI MATERIAIS (Só renderiza se a caixa estiver marcada) */}
            {faturamentoDireto && (
              <div className="space-y-4 border-l border-stone-100 pl-4 sm:pl-6">
                <h4 className="font-bold text-emerald-800 uppercase text-[10px] bg-emerald-50 px-2 py-1 rounded tracking-wide">
                  2. Taxas Exclusivas para Materiais
                </h4>
                
                <div className="space-y-3">
                  <h5 className="font-medium text-emerald-600/70 uppercase text-[9px]">Administração e Riscos</h5>
                  <BdiInput label="Administração Central" value={bdiMats.admCentral} onChange={(v) => handleMatChange("admCentral", v)} />
                  <BdiInput label="Contabilidade / Seguros" value={bdiMats.contabilidade} onChange={(v) => handleMatChange("contabilidade", v)} />
                  <BdiInput label="Contingenciamento" value={bdiMats.contingenciamento} onChange={(v) => handleMatChange("contingenciamento", v)} />
                  <BdiInput label="Custo Financeiro" value={bdiMats.custoFinanceiro} onChange={(v) => handleMatChange("custoFinanceiro", v)} />
                  
                  <h5 className="font-medium text-emerald-600/70 uppercase text-[9px] pt-1">Margem e Impostos</h5>
                  <BdiInput label="Lucro Real de Venda" value={bdiMats.lucro} onChange={(v) => handleMatChange("lucro", v)} />
                  <BdiInput label="DAS / Tributos (Anexo IV)" value={bdiMats.dasAnexoIV} onChange={(v) => handleMatChange("dasAnexoIV", v)} />
                  <BdiInput label="ART / Encargos Contrato" value={bdiMats.art} onChange={(v) => handleMatChange("art", v)} />
                </div>

                <div className="pt-2 border-t border-stone-100 flex justify-between items-center text-[11px] font-bold text-emerald-800">
                  <span>Taxa BDI Materiais:</span>
                  <span className="font-mono bg-emerald-50 text-emerald-800 px-1.5 py-0.5 rounded">{fmt(bdiMatRate * 100)}%</span>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* PAINEL DA DIREITA: RESUMO TOTALIZADOR */}
        <div className="bg-stone-900 text-stone-100 rounded-lg p-5 flex flex-col justify-between h-full min-h-[320px]">
          <div>
            <h3 className="font-semibold text-xs uppercase tracking-wider text-stone-400 mb-4">Resumo Geral de Fechamento</h3>
            <div className="space-y-3 text-xs">
              <div className="flex justify-between"><span className="text-stone-400">Custo Direto Base:</span><span className="font-mono">R$ {fmt(grandTotal)}</span></div>
              <div className="flex justify-between">
                <span className="text-stone-400">BDI Geral Aplicado:</span>
                <span className="font-mono text-stone-300">{fmt(bdiCalc.bdiRate * 100)}%</span>
              </div>
              {faturamentoDireto && (
                <div className="flex justify-between">
                  <span className="text-emerald-400">BDI Materiais Aplicado:</span>
                  <span className="font-mono text-emerald-300">{fmt(bdiCalc.bdiRateMateriais * 100)}%</span>
                </div>
              )}
              <div className="flex justify-between border-t border-stone-800 pt-2">
                <span className="text-stone-400">Total BDI (Rateio):</span>
                <span className="font-mono">R$ {fmt(bdiCalc.totalDiValor)}</span>
              </div>
            </div>
          </div>
          <div className="mt-6 pt-4 border-t border-stone-800 text-right">
            <span className="text-[10px] text-stone-400 block uppercase font-medium">Preço Final de Venda</span>
            <span className="text-2xl font-bold font-mono text-white">R$ {fmt(bdiCalc.valorVenda)}</span>
          </div>
        </div>
      </div>
    </div>
  );
}

function BdiInput({ label, value, onChange }) {
  return (
    <div className="flex items-center justify-between gap-3">
      <span className="text-stone-600">{label}</span>
      <input type="number" step="any" value={value === 0 ? "" : num(value) * 100} onChange={(e) => onChange(e.target.value === "" ? 0 : num(e.target.value) / 100)} className="w-20 border border-stone-300 rounded px-2 py-1 text-right font-mono" placeholder="0.00" />
    </div>
  );
}

/* ---------------- ABA FECHAMENTO: PREÇO DE VENDA ---------------- */
function PrecoVenda({ etapas, FatorBdi, grandTotal, nomeProjeto, cpus, catalogMap }) {
  const exportarXls = () => {
    const wb = XLSX.utils.book_new();
    const rows = [
      [`Planilha de Preço de Venda - ${nomeProjeto || "Orçamento"}`],
      [`Fator BDI aplicado: ${FatorBdi.toLocaleString("pt-BR", { minimumFractionDigits: 4, maximumFractionDigits: 4 })}`],
      [],
      ["Etapa", "Serviço", "Qtd.", "Un.", "Custo Unit. (R$)", "Preço Venda Unit. (R$)", "Total Venda (R$)"],
    ];
    (etapas || []).forEach((e) => {
      (e.itens || []).forEach((it) => {
        const uCusto = cpuValorUnit(it.insumos, cpus, catalogMap);
        rows.push([e.nome, it.servico, num(it.quantidade), it.unidade, uCusto, uCusto * FatorBdi, num(it.quantidade) * uCusto * FatorBdi]);
      });
    });
    rows.push([]);
    rows.push(["", "", "", "", "", "TOTAL GERAL", grandTotal * FatorBdi]);
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 20 }, { wch: 45 }, { wch: 8 }, { wch: 6 }, { wch: 20 }, { wch: 22 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(wb, ws, "Preço de Venda");
    XLSX.writeFile(wb, "preco_de_venda.xlsx");
  };

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4 space-y-4">
      <div className="border-b border-stone-100 pb-2 flex justify-between items-center">
        <h3 className="font-semibold text-sm text-stone-800">Planilha Sintética de Fechamento (Preço de Venda)</h3>
        <button onClick={exportarXls} className="flex items-center gap-1.5 px-3 py-1.5 text-xs border border-stone-300 rounded-lg font-medium bg-white hover:bg-stone-50 text-stone-700">
          <Download size={13} /> Exportar .xlsx
        </button>
      </div>
      <div className="space-y-3">
        {etapas.map((e) => {
          const custoEtapa = e.itens.reduce((s, it) => s + num(it.quantidade) * cpuValorUnit(it.insumos, cpus, catalogMap), 0);
          return (
            <div key={e.id} className="border border-stone-100 rounded-lg overflow-hidden">
              <div className="bg-stone-50/50 px-4 py-2 flex justify-between text-xs font-semibold text-stone-700">
                <span>{e.nome}</span>
                <span className="font-mono">R$ {fmt(custoEtapa * FatorBdi)}</span>
              </div>
              <div className="divide-y divide-stone-50">
                {e.itens.map((it) => {
                  const uCusto = cpuValorUnit(it.insumos, cpus, catalogMap);
                  const totalVendaItem = num(it.quantidade) * (uCusto * FatorBdi);
                  return (
                    <div key={it.id} className="flex items-center justify-between gap-4 px-4 py-2 text-xs">
                      <span className="text-stone-700 truncate flex-1">{it.servico}</span>
                      <span className="text-stone-400 font-mono w-24 text-right">{it.quantidade} {it.unidade}</span>
                      <span className="text-stone-500 font-mono w-28 text-right">R$ {fmt(uCusto * FatorBdi)}/un.</span>
                      <span className="font-medium font-mono text-stone-900 w-28 text-right">R$ {fmt(totalVendaItem)}</span>
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
      <div className="pt-4 border-t border-stone-200 flex justify-end">
        <div className="text-right p-2">
          <span className="text-xs text-stone-400 block font-medium">Valor Total do Fechamento Comercial</span>
          <span className="text-xl font-bold font-mono text-stone-900">R$ {fmt(grandTotal * FatorBdi)}</span>
        </div>
      </div>
    </div>
  );
}



