import React, { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx-js-style";
import {
  Plus, Trash2, Pencil, X, Search, Upload, Download,
  ChevronDown, ChevronRight, Database, Calculator, Copy, Save, Percent, TrendingUp, RefreshCw,
  Tags, AlertTriangle, Check, FolderKanban, HardHat, User, LogIn, MapPin, Phone, Mail, Building2, FileText,
  ArrowDown, ArrowUp, ArrowUpDown, CalendarDays, BarChart3, ChevronLeft
} from "lucide-react";
import HistogramaMaoObra from "./HistogramaMaoObra";
import { FONTES_PADRAO, TIPOS, createDefaultProject, seedCpus } from "./data/defaultData";
import {
  applyCatalogToInsumos,
  buildCatalog,
  criarIndiceBuscaCpus,
  cpuValorUnit,
  findSubCpu,
  insumosResolvidosSubCpu,
  insumoValorUnitario,
  precoKey,
} from "./utils/calculos";
import {
  avaliarExpressaoNumerica,
  formatarCep,
  formatarCpfCnpj,
  formatarTelefone,
  fmt,
  norm,
  normalizarBusca,
  num,
  uid,
} from "./utils/format";
import { consolidarMaoDeObra } from "./utils/maoObra";
import {
  etapaComOpcaoAtiva,
  etapasComOpcaoAtiva,
  grupoAlternativaDoItem,
  gruposAlternativasDaEtapa,
  itemIncluidoNoCalculo,
  itensAtivosDaEtapa,
} from "./utils/alternativas";
import {
  deleteGoogleDriveProject,
  loadGoogleDriveSnapshot,
  requestGoogleDriveAccess,
  saveGoogleDriveSnapshot,
} from "./services/googleDriveStore";
import {
  loadLocalSnapshot,
  loadOrcamentoData,
  saveLocalSnapshot,
} from "./services/orcamentoStore";
import alphaLogo from "./assets/alpha-engenharia-logo.png";

const BDI_PADRAO = {
  custoInicial: 0,
  admCentral: 0.04,
  contabilidade: 0.01,
  contingenciamento: 0.02,
  custoFinanceiro: 0.03,
  dasAnexoIV: 0.13,
  art: 0,
  retencaoInss: 0,
  lucro: 0.42,
  collemAtivo: false,
  collemX: 1,
  collemY: 1,
};

const CLIENTE_PADRAO = {
  clienteId: "",
  nome: "",
  local: "",
  contato: "",
  telefone: "",
  email: "",
  documento: "",
  cep: "",
  endereco: "",
  numeroProposta: "",
  modeloProposta: "",
  regimeMateriais: "alpha",
  prazoExecucao: "",
  condicoesPagamento: "",
  percentualSinalCollem: 20,
  prazoExecucaoCollem: "",
  textoApresentacaoCollem: "",
  naoInclusosCollem: "",
  condicoesEspeciaisCollem: "",
  responsavelCollem: "Geraldo Belloni Perez",
  responsabilidadesAlpha: "",
  responsabilidadesCliente: "",
  observacoes: "",
};

const RESPONSABILIDADES_ALPHA_PADRAO = [
  "Acompanhamento Técnico;",
  "Fornecimento de EPIs para execução das atividades;",
  "Fornecimento de mão de obra;",
  "Fornecimento de equipamentos;",
  "Fornecimento de almoço e transporte para funcionários;",
  "Fornecimento de material conforme composição do orçamento.",
];

const RESPONSABILIDADES_CLIENTE_PADRAO = [
  "Fornecimento de acesso ao local de prestação de serviço;",
  "Permitir os funcionários a usarem as instalações sanitárias;",
  "Fornecimento de água potável, água bruta e energia;",
  "Local para armazenamento de materiais e equipamentos;",
];

const listaTextoOuPadrao = (texto, padrao) => {
  const linhas = String(texto || "")
    .split(/\r?\n/)
    .map((linha) => linha.trim())
    .filter(Boolean);
  return linhas.length ? linhas : padrao;
};

const clienteDoProjeto = (projeto) => ({
  ...CLIENTE_PADRAO,
  ...(projeto?.clienteCadastro || {}),
  nome: projeto?.clienteCadastro?.nome || projeto?.cliente || "",
});

const CAMPOS_CLIENTE_COMPARTILHADOS = [
  "nome",
  "contato",
  "telefone",
  "email",
  "documento",
  "cep",
  "endereco",
];

const dadosClienteCompartilhado = (cliente = {}) => ({
  nome: cliente.nome || "",
  contato: cliente.contato || "",
  telefone: formatarTelefone(cliente.telefone || ""),
  email: cliente.email || "",
  documento: formatarCpfCnpj(cliente.documento || ""),
  cep: formatarCep(cliente.cep || ""),
  endereco: cliente.endereco || "",
});

const chaveClienteCompartilhado = (cliente = {}) => {
  const documento = String(cliente.documento || "").replace(/\D/g, "");
  if (documento.length >= 11) return `doc:${documento}`;
  const email = normalizarBusca(cliente.email || "");
  if (email) return `email:${email}`;
  const nome = normalizarBusca(cliente.nome || "");
  return nome ? `nome:${nome}` : "";
};

const prepararClientesCompartilhados = (projetos = [], clientesSalvos = []) => {
  const clientes = [];
  const porId = new Map();
  const porChave = new Map();

  (clientesSalvos || []).forEach((clienteSalvo) => {
    const cliente = {
      id: clienteSalvo.id || uid(),
      ...dadosClienteCompartilhado(clienteSalvo),
      atualizadoEm: clienteSalvo.atualizadoEm || "",
    };
    clientes.push(cliente);
    porId.set(cliente.id, cliente);
    const chave = chaveClienteCompartilhado(cliente);
    if (chave && !porChave.has(chave)) porChave.set(chave, cliente);
  });

  const projetosPreparados = (projetos || []).map((projeto) => {
    const cadastro = clienteDoProjeto(projeto);
    if (!String(cadastro.nome || "").trim()) return projeto;

    const chave = chaveClienteCompartilhado(cadastro);
    let compartilhado = porId.get(cadastro.clienteId) || (chave ? porChave.get(chave) : null);
    if (!compartilhado) {
      compartilhado = {
        id: cadastro.clienteId || uid(),
        ...dadosClienteCompartilhado(cadastro),
        atualizadoEm: projeto.atualizadoEm || projeto.criadoEm || "",
      };
      clientes.push(compartilhado);
      porId.set(compartilhado.id, compartilhado);
      if (chave) porChave.set(chave, compartilhado);
    }

    const dadosCompartilhados = Object.fromEntries(
      CAMPOS_CLIENTE_COMPARTILHADOS.map((campo) => [
        campo,
        compartilhado[campo] || cadastro[campo] || "",
      ])
    );
    Object.assign(compartilhado, dadosCompartilhados);

    return {
      ...projeto,
      cliente: compartilhado.nome || cadastro.nome || "",
      clienteCadastro: {
        ...cadastro,
        ...dadosCompartilhados,
        clienteId: compartilhado.id,
      },
    };
  });

  clientes.sort((a, b) => String(a.nome || "").localeCompare(String(b.nome || ""), "pt-BR"));
  return { clientes, projetos: projetosPreparados };
};

const clienteEstaCompleto = (cliente) =>
  Boolean(String(cliente?.nome || "").trim() && String(cliente?.local || "").trim());

const obterStatusProjeto = (projeto) => {
  if (!clienteEstaCompleto(clienteDoProjeto(projeto))) {
    return {
      id: "cadastro_pendente",
      label: "Cadastro pendente",
      className: "border-amber-200 bg-amber-50 text-amber-700",
    };
  }

  const status = normalizarBusca(projeto?.status || "");
  if (status === "rascunho") {
    return {
      id: "rascunho",
      label: "Rascunho",
      className: "border-stone-200 bg-stone-100 text-stone-600",
    };
  }
  if (status === "concluido") {
    return {
      id: "concluido",
      label: "Concluído",
      className: "border-emerald-200 bg-emerald-50 text-emerald-700",
    };
  }
  if (status === "enviado cliente") {
    return {
      id: "enviado_cliente",
      label: "Enviado p/ cliente",
      className: "border-amber-400 bg-amber-300 text-amber-950",
    };
  }
  if (status === "cancelado") {
    return {
      id: "cancelado",
      label: "Cancelado",
      className: "border-stone-300 bg-stone-200 text-stone-700",
    };
  }
  if (status === "aprovado") {
    return {
      id: "aprovado",
      label: "Aprovado",
      className: "border-emerald-300 bg-emerald-100 text-emerald-800",
    };
  }
  if (status === "reprovado") {
    return {
      id: "reprovado",
      label: "Reprovado",
      className: "border-rose-200 bg-rose-50 text-rose-700",
    };
  }

  return {
    id: "em_elaboracao",
    label: "Em elaboração",
    className: "border-emerald-200 bg-emerald-50 text-emerald-700",
  };
};

const formatarAtualizacaoProjeto = (valor) => {
  if (!valor) return "Sem registro";
  const data = new Date(valor);
  if (Number.isNaN(data.getTime())) return "Sem registro";

  const agora = new Date();
  const inicioHoje = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate());
  const inicioData = new Date(data.getFullYear(), data.getMonth(), data.getDate());
  const dias = Math.round((inicioHoje - inicioData) / 86400000);
  const hora = data.toLocaleTimeString("pt-BR", {
    hour: "2-digit",
    minute: "2-digit",
  });

  if (dias === 0) return `Hoje, ${hora}`;
  if (dias === 1) return `Ontem, ${hora}`;
  return data.toLocaleDateString("pt-BR");
};

const aguardar = (tempoMs) => new Promise((resolve) => setTimeout(resolve, tempoMs));

const intervaloNovaTentativa = (tentativa) =>
  Math.min(30000, 2000 * (2 ** Math.min(Math.max(tentativa - 1, 0), 4)));

const prepararBancosDePrecosPorProjeto = (projetos = [], precosLegados = []) => {
  const precosAtuais = Array.isArray(precosLegados) ? precosLegados : [];
  const projetosSemBanco = projetos.filter(
    (projeto) =>
      !Object.prototype.hasOwnProperty.call(projeto, "precos") &&
      !projeto.bancoPrecosInicializado
  );

  if (projetosSemBanco.length === 0) {
    return {
      projetos: projetos.map((projeto) => ({
        ...projeto,
        precos: Array.isArray(projeto.precos) ? projeto.precos : [],
        bancoPrecosInicializado: true,
      })),
      projetoMigrado: null,
    };
  }

  const textoProjeto = (projeto) => {
    const cliente = clienteDoProjeto(projeto);
    return norm(
      `${projeto?.nome || ""} ${projeto?.cliente || ""} ${cliente.nome || ""} ${cliente.local || ""}`
    );
  };
  const projetoMigrado =
    projetosSemBanco.find((projeto) => {
      const texto = textoProjeto(projeto);
      return (
        texto.includes("reforma cozinha") &&
        texto.includes("santuario") &&
        texto.includes("caraca")
      );
    }) ||
    projetosSemBanco.find((projeto) => textoProjeto(projeto).includes("reforma cozinha")) ||
    null;

  return {
    projetos: projetos.map((projeto) => {
      const jaPossuiBanco =
        Object.prototype.hasOwnProperty.call(projeto, "precos") ||
        projeto.bancoPrecosInicializado;
      if (jaPossuiBanco) {
        return {
          ...projeto,
          precos: Array.isArray(projeto.precos) ? projeto.precos : [],
          bancoPrecosInicializado: true,
        };
      }

      return {
        ...projeto,
        precos:
          projeto.id === projetoMigrado?.id
            ? precosAtuais.map((preco) => ({ ...preco }))
            : [],
        bancoPrecosInicializado: true,
      };
    }),
    projetoMigrado,
  };
};

const materialPorContaCliente = (cliente) => cliente?.regimeMateriais === "cliente";
const materialFaturamentoDireto = (cliente) => cliente?.regimeMateriais === "faturamentoDireto";
const insumoEhMaterial = (tipo) => {
  const t = String(tipo || "").toUpperCase().trim();
  return t === "MAT" || t === "MATERIAL" || (!t.includes("MO") && !t.includes("MÃO") && !t.includes("MAO") && !t.includes("EQUIP"));
};

const insumoEhMaoDeObra = (tipo) => {
  const t = String(tipo || "").toUpperCase().trim();
  return t === "MO" || t.includes("MÃO DE OBRA") || t.includes("MAO DE OBRA");
};

const insumoComFaturamentoDireto = (insumo, configuracao = {}, cliente = {}) => {
  if (!insumoEhMaterial(insumo?.tipo)) return false;
  if (!configuracao?.faturamentoDireto && !materialFaturamentoDireto(cliente)) return false;

  const selecionados = configuracao?.materiaisFaturamentoDireto;
  if (!Array.isArray(selecionados)) return true;
  return selecionados.includes(precoKey(insumo?.descricao));
};

const fatorVendaInsumo = (insumo, bdiCalc = {}, cliente = {}) => {
  if (insumoComFaturamentoDireto(insumo, bdiCalc, cliente)) {
    return num(bdiCalc.FatorBdiMateriais) || 1;
  }
  if (insumoEhMaoDeObra(insumo?.tipo)) {
    return num(bdiCalc.FatorBdiMaoObra) || num(bdiCalc.FatorBdi) || 1;
  }
  return num(bdiCalc.FatorBdi) || 1;
};

const buscarInsumosCatalogo = (catalogMap, textoBusca, limite = 12) => {
  const busca = normalizarBusca(textoBusca).trim();
  const termos = busca.split(/\s+/).filter(Boolean);
  if (!busca || termos.length === 0 || !catalogMap) return [];

  return Array.from(catalogMap.values())
    .filter((insumo) => {
      const descricao = normalizarBusca(insumo.descricao);
      return termos.every((termo) => descricao.includes(termo));
    })
    .sort((a, b) => {
      const descricaoA = normalizarBusca(a.descricao);
      const descricaoB = normalizarBusca(b.descricao);
      const inicioA = descricaoA.startsWith(busca) ? 0 : 1;
      const inicioB = descricaoB.startsWith(busca) ? 0 : 1;
      return inicioA - inicioB || descricaoA.length - descricaoB.length || descricaoA.localeCompare(descricaoB, "pt-BR");
    })
    .slice(0, limite);
};

const nomeArquivoSeguro = (valor) =>
  String(valor || "Orcamento")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const proximoNumeroProposta = (projetos = [], data = new Date()) => {
  const anoCompleto = data.getFullYear();
  const anoCurto = String(anoCompleto).slice(-2);
  const sequenciaInicialPorAno = { 2026: 74 };
  let maiorSequencia = (sequenciaInicialPorAno[anoCompleto] || 1) - 1;

  (projetos || []).forEach((projeto) => {
    const numeroAtual = String(projeto?.clienteCadastro?.numeroProposta || "").trim();
    const match = numeroAtual.match(/(\d+)\s*[\/-]\s*(\d{2}|\d{4})\s*$/);
    if (!match) return;

    const anoNumero = match[2].length === 2 ? Number(`20${match[2]}`) : Number(match[2]);
    if (anoNumero !== anoCompleto) return;
    maiorSequencia = Math.max(maiorSequencia, Number(match[1]) || 0);
  });

  return `PROP - ${String(maiorSequencia + 1).padStart(2, "0")}/${anoCurto}`;
};

const numeroMoeda = (valor) => {
  if (typeof valor === "number") return Number.isFinite(valor) ? valor : 0;
  const texto = String(valor ?? "").trim();
  if (!texto) return 0;
  const normalizado = texto.includes(",")
    ? texto.replace(/\./g, "").replace(",", ".")
    : texto;
  const numero = Number(normalizado);
  return Number.isFinite(numero) ? numero : 0;
};

const descontoNegociacaoProjeto = (projeto) =>
  Math.max(0, numeroMoeda(projeto?.descontoNegociacao));

const aplicarDescontoNegociacao = (valorBruto, desconto) =>
  Math.max(0, num(valorBruto) - Math.max(0, numeroMoeda(desconto)));

const VERSAO_NUMERACAO_PROPOSTAS = 1;
const ANO_INICIAL_PROPOSTAS = 2026;
const SEQUENCIA_INICIAL_PROPOSTAS = 74;

const numerarProjetosExistentes = (projetos = []) => {
  const lista = Array.isArray(projetos) ? projetos : [];
  if (lista.length === 0) return { projetos: lista, idsMigrados: [] };

  const ordenados = lista
    .map((projeto, indiceOriginal) => ({ projeto, indiceOriginal }))
    .sort((a, b) => {
      const dataA = Date.parse(a.projeto?.criadoEm || "");
      const dataB = Date.parse(b.projeto?.criadoEm || "");
      const tempoA = Number.isFinite(dataA) ? dataA : Number.MAX_SAFE_INTEGER;
      const tempoB = Number.isFinite(dataB) ? dataB : Number.MAX_SAFE_INTEGER;
      return tempoA - tempoB || a.indiceOriginal - b.indiceOriginal;
    });

  const indiceAncora = ordenados.findIndex(({ projeto }) => {
    const cliente = clienteDoProjeto(projeto);
    const texto = normalizarBusca(
      `${projeto?.nome || ""} ${cliente.nome || ""} ${cliente.contato || ""}`
    );
    return texto.includes("supressao arborea") && texto.includes("fernanda");
  });
  const indiceNumero74 = ordenados.findIndex(({ projeto }) =>
    /(^|\D)74\s*[\/-]\s*26\s*$/.test(
      String(projeto?.clienteCadastro?.numeroProposta || "").trim()
    )
  );
  const indiceBase = indiceAncora >= 0 ? indiceAncora : indiceNumero74;
  if (indiceBase < 0) return { projetos: lista, idsMigrados: [] };

  const numeroPorId = new Map(
    ordenados.map(({ projeto }, indice) => [
      projeto.id,
      `PROP - ${String(SEQUENCIA_INICIAL_PROPOSTAS + indice - indiceBase).padStart(2, "0")}/${String(ANO_INICIAL_PROPOSTAS).slice(-2)}`,
    ])
  );
  const idsMigrados = [];
  const projetosNumerados = lista.map((projeto) => {
    const numeroAtual = String(projeto?.clienteCadastro?.numeroProposta || "").trim();
    if (
      Number(projeto?.numeracaoPropostaVersao || 0) >= VERSAO_NUMERACAO_PROPOSTAS &&
      numeroAtual
    ) {
      return projeto;
    }

    idsMigrados.push(projeto.id);
    return {
      ...projeto,
      numeracaoPropostaVersao: VERSAO_NUMERACAO_PROPOSTAS,
      clienteCadastro: {
        ...(projeto.clienteCadastro || {}),
        numeroProposta: numeroPorId.get(projeto.id),
      },
    };
  });

  return { projetos: projetosNumerados, idsMigrados };
};

const escapeHtml = (valor) =>
  String(valor ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

const itemVendaResumo = (item, bdiCalc, cpus, catalogMap, cliente = {}) => {
  const quantidade = num(item.quantidade);
  let total = 0;

  (item.insumos || []).forEach((ins) => {
    const isMaterial = insumoEhMaterial(ins.tipo);
    if (materialPorContaCliente(cliente) && isMaterial) return;

    const custoBase = num(ins.coeficiente) * quantidade * insumoValorUnitario(ins, cpus, catalogMap);
    total += custoBase * fatorVendaInsumo(ins, bdiCalc, cliente);
  });

  return {
    quantidade,
    total,
    unitario: quantidade > 0 ? total / quantidade : 0,
  };
};

const montarItensProposta = (etapas, bdiCalc, cpus, catalogMap, cliente = {}) =>
  (etapas || []).map((etapa, idxEtapa) => {
    const itens = itensAtivosDaEtapa(etapa).map((item, idxItem) => ({
      numero: `${idxEtapa + 1}.${idxItem + 1}`,
      descricao: item.servico || item.descricao || "",
      unidade: item.unidade || "",
      ...itemVendaResumo(item, bdiCalc, cpus, catalogMap, cliente),
    }));

    return {
      numero: String(idxEtapa + 1),
      nome: etapa.nome || `Etapa ${idxEtapa + 1}`,
      total: itens.reduce((s, item) => s + item.total, 0),
      itens,
    };
  });

const montarComparativosProposta = (
  etapas,
  bdiCalc,
  cpus,
  catalogMap,
  cliente = {}
) =>
  (etapas || []).flatMap((etapa) =>
    gruposAlternativasDaEtapa(etapa).map((grupo) => ({
      etapaNome: etapa.nome,
      grupoNome: grupo.nome,
      opcoes: (grupo.opcoes || []).map((opcao) => {
        const etapasDaOpcao = etapasComOpcaoAtiva(
          etapas,
          etapa.id,
          grupo.id,
          opcao.id
        );
        const valorVenda = etapasDaOpcao.reduce(
          (totalEtapas, etapaAtual) =>
            totalEtapas +
            itensAtivosDaEtapa(etapaAtual).reduce(
              (totalItens, item) =>
                totalItens +
                itemVendaResumo(
                  item,
                  bdiCalc,
                  cpus,
                  catalogMap,
                  cliente
                ).total,
              0
            ),
          0
        );
        return {
          ...opcao,
          selecionada: grupo.opcaoAtivaId === opcao.id,
          valorVenda: aplicarDescontoNegociacao(
            valorVenda,
            bdiCalc?.descontoNegociacao
          ),
        };
      }),
    }))
  );

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

const criarAbaVendaModelo = (
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
  const totalBruto = grupos.reduce((total, grupo) => total + num(grupo.total), 0);
  const descontoAplicado = Math.min(
    totalBruto,
    Math.max(0, num(descontoNegociacao))
  );
  let descontoRow = 0;
  if (descontoAplicado > 0) {
    descontoRow = rows.length + 1;
    rows.push([null, "DESCONTO DA NEGOCIAÇÃO", null, null, null, null, null, -descontoAplicado]);
  }
  const totalRow = rows.length + 1;
  const totalGeral = aplicarDescontoNegociacao(totalBruto, descontoAplicado);
  rows.push([null, "TOTAL GERAL", null, null, null, null, null, totalGeral]);

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const totalAddr = XLSX.utils.encode_cell({ r: totalRow - 1, c: 7 });
  const ultimaLinhaValores = Math.max(4, totalRow - 1);
  ws[totalAddr] = {
    t: "n",
    v: totalGeral,
    f: `SUM(H4:H${ultimaLinhaValores})`,
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
    aplicarFormatoNumerico(ws, row, [6, 7, 8], collem ? XLSX_MOEDA : XLSX_NUMERO);
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

const exportarPropostaXlsx = ({ projeto, cliente, etapas, bdiCalc, cpus, catalogMap, modelo }) => {
  const grupos = montarItensProposta(etapas, bdiCalc, cpus, catalogMap, cliente);
  const comparativos = montarComparativosProposta(
    etapas,
    bdiCalc,
    cpus,
    catalogMap,
    cliente
  );
  const wb = XLSX.utils.book_new();
  wb.Workbook = {
    ...(wb.Workbook || {}),
    CalcPr: {
      ...(wb.Workbook?.CalcPr || {}),
      calcMode: "auto",
      fullCalcOnLoad: true,
      forceFullCalc: true,
    },
  };
  const wsValores = criarAbaVendaModelo(
    grupos,
    bdiCalc?.FatorBdi || 1,
    modelo,
    bdiCalc?.descontoNegociacao || 0
  );
  XLSX.utils.book_append_sheet(wb, wsValores, "VENDA");
  if (comparativos.length > 0) {
    const linhas = [["ETAPA", "GRUPO", "ALTERNATIVA", "SELECIONADA", "TOTAL DA PROPOSTA"]];
    comparativos.forEach((grupo) =>
      grupo.opcoes.forEach((opcao) =>
        linhas.push([
          grupo.etapaNome,
          grupo.grupoNome,
          opcao.nome,
          opcao.selecionada ? "SIM" : "NÃO",
          opcao.valorVenda,
        ])
      )
    );
    const wsAlternativas = XLSX.utils.aoa_to_sheet(linhas);
    wsAlternativas["!cols"] = [
      { wch: 28 },
      { wch: 32 },
      { wch: 28 },
      { wch: 14 },
      { wch: 20 },
    ];
    XLSX.utils.book_append_sheet(wb, wsAlternativas, "ALTERNATIVAS");
  }
  const sufixo = modelo === "collem" ? "Anexo_I_COLLEM" : "Proposta_ALPHA";
  XLSX.writeFile(wb, `${nomeArquivoSeguro(projeto.nome)}_${sufixo}.xlsx`);
};

const gerarPropostaPdf = ({ projeto, cliente, etapas, bdiCalc, cpus, catalogMap, numeroPropostaAutomatico }) => {
  const grupos = montarItensProposta(etapas, bdiCalc, cpus, catalogMap, cliente);
  const comparativos = montarComparativosProposta(
    etapas,
    bdiCalc,
    cpus,
    catalogMap,
    cliente
  );
  const totalBruto = grupos.reduce((s, grupo) => s + grupo.total, 0);
  const descontoNegociacao = Math.min(
    totalBruto,
    Math.max(0, num(bdiCalc?.descontoNegociacao))
  );
  const totalGeral = aplicarDescontoNegociacao(
    totalBruto,
    descontoNegociacao
  );
  const hoje = new Date();
  const dataHoje = hoje.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit", year: "numeric" });
  const numeroProposta = cliente?.numeroProposta || numeroPropostaAutomatico || `PROP - 01/${String(hoje.getFullYear()).slice(-2)}`;
  const nomeProjeto = projeto?.nome || "Orçamento";
  const nomeCliente = cliente?.nome || "Cliente";
  const localObra = cliente?.local || cliente?.endereco || "";
  const contato = cliente?.contato || "";
  const observacoes = cliente?.observacoes || "";
  const prazoExecucao = cliente?.prazoExecucao || "A definir conforme cronograma aprovado entre as partes.";
  const condicoesPagamento =
    cliente?.condicoesPagamento ||
    `Entrada de 40% (R$ ${fmt(totalGeral * 0.4)}) e o restante (R$ ${fmt(totalGeral * 0.6)}) conforme avanço dos serviços em medições.`;
  const responsabilidadesAlpha = listaTextoOuPadrao(cliente?.responsabilidadesAlpha, RESPONSABILIDADES_ALPHA_PADRAO);
  const responsabilidadesCliente = listaTextoOuPadrao(cliente?.responsabilidadesCliente, RESPONSABILIDADES_CLIENTE_PADRAO);
  const descricaoRegimeMateriais = materialPorContaCliente(cliente)
    ? "Material por conta do cliente. A proposta considera somente os serviços, mão de obra, equipamentos e demais custos não classificados como material."
    : materialFaturamentoDireto(cliente) || bdiCalc.faturamentoDireto
      ? Array.isArray(bdiCalc.materiaisFaturamentoDireto)
        ? "Somente os materiais selecionados no orçamento são considerados com faturamento direto para o cliente, aplicando o BDI específico configurado."
        : "Materiais considerados com faturamento direto para o cliente, aplicando BDI específico de materiais quando configurado."
      : "Materiais inclusos no fornecimento da ALPHA ENGENHARIA conforme composição do orçamento.";

  const itensEscopo = grupos
    .map((grupo) => `<li>${escapeHtml(grupo.nome)}</li>`)
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

  const linhasAlternativas = comparativos
    .flatMap((grupo) =>
      grupo.opcoes.map(
        (opcao) => `
          <tr class="${opcao.selecionada ? "alternativa-selecionada" : ""}">
            <td>${escapeHtml(grupo.grupoNome)}</td>
            <td>${escapeHtml(opcao.nome)}</td>
            <td>${opcao.selecionada ? "Considerada no total" : "Alternativa"}</td>
            <td>R$ ${fmt(opcao.valorVenda)}</td>
          </tr>
        `
      )
    )
    .join("");

  const listaResponsabilidadesAlpha = responsabilidadesAlpha.map((item) => `<li>${escapeHtml(item)}</li>`).join("");
  const listaResponsabilidadesCliente = responsabilidadesCliente.map((item) => `<li>${escapeHtml(item)}</li>`).join("");

  const html = `
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <title>${escapeHtml(nomeArquivoSeguro(nomeProjeto))}_Proposta</title>
  <style>
    @page { size: A4; margin: 0; }
    * { box-sizing: border-box; }
    :root { --alpha: #7B9A56; --alpha-dark: #111111; --alpha-soft: #e2efd9; --line: #c6d4b1; --alpha-gray: #a6a6a6; }
    body { margin: 0; font-family: "Arial Narrow", Arial, sans-serif; color: #111; font-size: 12pt; }
    .page { width: 210mm; height: 297mm; padding: 12mm 12.5mm 15mm; page-break-after: always; break-after: page; position: relative; display: flex; flex-direction: column; }
    .page + .page { page-break-before: always; break-before: page; }
    .page:last-child { page-break-after: auto; }
    .page::before, .page::after { content: ""; position: absolute; left: 0; right: 0; height: 8mm; background: var(--alpha); }
    .page::before { top: 0; }
    .page::after { bottom: 0; }
    header { display: flex; flex: 0 0 auto; justify-content: space-between; align-items: flex-end; margin-bottom: 6mm; page-break-inside: avoid; break-inside: avoid; }
    .logo { width: 26mm; height: 20mm; object-fit: contain; display: block; }
    .prop { font-weight: 700; font-size: 12pt; text-align: right; color: var(--alpha-gray); transform: translateY(7mm); }
    .pagina-topo { display: none; }
    h1 { text-align: center; font-size: 16pt; line-height: 1.15; margin: 0 0 12mm; font-weight: 700; color: #111; }
    h2 { font-size: 12pt; line-height: 1.15; margin: 10px 0 6px; font-weight: 700; color: #111; }
    p { margin: 0 0 6px; line-height: 1.5; text-align: justify; }
    .data { text-align: right; line-height: 1.15; margin-bottom: 10px; }
    .ref { font-weight: 700; }
    table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 9pt; line-height: 1.15; margin-top: 4px; }
    th { background: var(--alpha); color: #fff; font-weight: 700; text-align: left; }
    th, td { padding: 3px 4px; vertical-align: middle; border-bottom: 0.4pt solid #ececec; }
    tbody tr { page-break-inside: avoid; }
    .escopo-lista { margin: 4px 0 12px 22px; padding: 0; font-size: 12pt; line-height: 1.35; }
    .escopo-lista li { margin: 0 0 5px; padding-left: 2px; font-weight: 600; text-align: left; }
    .valores th:nth-child(1), .valores td:nth-child(1) { width: 8%; }
    .valores th:nth-child(2), .valores td:nth-child(2) { width: 42%; }
    .valores th:nth-child(3), .valores td:nth-child(3) { width: 8%; text-align: center; }
    .valores th:nth-child(4), .valores td:nth-child(4) { width: 10%; text-align: right; }
    .valores th:nth-child(5), .valores td:nth-child(5),
    .valores th:nth-child(6), .valores td:nth-child(6),
    .valores th:nth-child(7), .valores td:nth-child(7) { width: 11%; text-align: right; }
    .grupo td { background: var(--alpha-soft); color: var(--alpha-dark); font-weight: 700; border-bottom: 0.6pt solid var(--line); }
    .total td { background: var(--alpha); color: #fff; font-weight: 700; font-size: 10pt; }
    .total td:first-child { text-align: center; }
    .alternativas th:nth-child(1), .alternativas td:nth-child(1) { width: 34%; }
    .alternativas th:nth-child(2), .alternativas td:nth-child(2) { width: 28%; }
    .alternativas th:nth-child(3), .alternativas td:nth-child(3) { width: 20%; }
    .alternativas th:nth-child(4), .alternativas td:nth-child(4) { width: 18%; text-align: right; }
    .alternativa-selecionada td { background: var(--alpha-soft); color: var(--alpha-dark); font-weight: 700; }
    ul { margin: 0 0 12px 18px; padding: 0; line-height: 1.15; }
    li { margin: 0 0 4px; text-align: justify; }
    .assinatura { margin-top: 48px; width: 260px; border-top: 1px solid var(--alpha-dark); text-align: center; padding-top: 6px; color: var(--alpha-dark); font-weight: 700; }
    .footer { flex: 0 0 auto; margin-top: auto; padding-top: 10mm; font-size: 10pt; color: var(--alpha-gray); display: grid; grid-template-columns: 1fr auto; gap: 12px; align-items: end; page-break-inside: avoid; break-inside: avoid; }
    .footer strong { display: block; font-weight: 700; margin-bottom: 0; color: var(--alpha-gray); }
    .footer .endereco { line-height: 1.15; }
    .footer .pagina { white-space: nowrap; color: var(--alpha-gray); font-weight: 400; }
    @media screen {
      body { background: #eee; padding: 20px; }
      .page { background: white; margin: 0 auto 20px; box-shadow: 0 4px 16px rgba(0,0,0,.12); }
    }
  </style>
</head>
<body>
  <section class="page">
    <header>
      <img class="logo" src="${escapeHtml(alphaLogo)}" alt="Alpha Engenharia" />
      <div class="prop">${escapeHtml(numeroProposta)}</div>
      <div class="pagina-topo">Página 1 de 3</div>
    </header>
    <h1>PROPOSTA DE PRESTAÇÃO DE SERVIÇOS</h1>
    <p class="data">Belo Horizonte ${escapeHtml(dataHoje)}</p>
    <p>Aos cuidados de ${escapeHtml(nomeCliente)}${contato ? ` - ${escapeHtml(contato)}` : ""}.</p>
    <p class="ref">Ref. ${escapeHtml(nomeProjeto)}</p>
    <p><strong>Endereço da Obra:</strong> ${escapeHtml(localObra)}</p>
    <h2>Escopo do Serviço:</h2>
    <ul class="escopo-lista">${itensEscopo}</ul>
    <div class="footer">
      <div class="endereco"><strong>ALPHA ENGENHARIA E SERVIÇOS</strong>Rua José Da Costa, 116 - São João Batista<br/>Belo Horizonte<br/>Telefone: 31 9 9203-1783</div>
      <div class="pagina">Página 1 de 3</div>
    </div>
  </section>

  <section class="page">
    <header>
      <img class="logo" src="${escapeHtml(alphaLogo)}" alt="Alpha Engenharia" />
      <div class="prop">${escapeHtml(numeroProposta)}</div>
      <div class="pagina-topo">Página 2 de 3</div>
    </header>
    <h2>Responsabilidade da ALPHA ENGENHARIA:</h2>
    <ul>${listaResponsabilidadesAlpha}</ul>
    <h2>Responsabilidade do Cliente:</h2>
    <ul>${listaResponsabilidadesCliente}</ul>
    <h2>Valores:</h2>
    <p>Segue relação da mão de obra especializada para execução e acompanhamento dos serviços apresentados em visita técnica, totalizando o valor de <strong>R$ ${fmt(totalGeral)}</strong>.</p>
    <p><strong>Condição dos materiais:</strong> ${escapeHtml(descricaoRegimeMateriais)}</p>
    ${comparativos.length > 0 ? `
      <h2>ALTERNATIVAS TÉCNICAS</h2>
      <table class="alternativas">
        <thead><tr><th>GRUPO</th><th>ALTERNATIVA</th><th>SITUAÇÃO</th><th>TOTAL DA PROPOSTA</th></tr></thead>
        <tbody>${linhasAlternativas}</tbody>
      </table>
    ` : ""}
    <h2>PLANILHA DE MATERIAL</h2>
    <table class="valores">
      <thead><tr><th>ITEM</th><th>DESCRIÇÃO DOS SERVIÇOS</th><th>UNID.</th><th>QUANT.</th><th>VALOR UNIT.</th><th>VALOR TOTAL</th><th>TOTAL DO ITEM</th></tr></thead>
      <tbody>
        ${linhasValores}
        ${descontoNegociacao > 0 ? `<tr class="grupo"><td colspan="6">DESCONTO DA NEGOCIAÇÃO</td><td>- R$ ${fmt(descontoNegociacao)}</td></tr>` : ""}
        <tr class="total"><td colspan="6">TOTAL GERAL</td><td>R$ ${fmt(totalGeral)}</td></tr>
      </tbody>
    </table>
    <div class="footer">
      <div class="endereco"><strong>ALPHA ENGENHARIA E SERVIÇOS</strong>Rua José Da Costa, 116 - São João Batista<br/>Belo Horizonte<br/>Telefone: 31 9 9203-1783</div>
      <div class="pagina">Página 2 de 3</div>
    </div>
  </section>

  <section class="page">
    <header>
      <img class="logo" src="${escapeHtml(alphaLogo)}" alt="Alpha Engenharia" />
      <div class="prop">${escapeHtml(numeroProposta)}</div>
      <div class="pagina-topo">Página 3 de 3</div>
    </header>
    <h2>Condições de pagamento:</h2>
    <p>${escapeHtml(condicoesPagamento).replace(/\n/g, "<br/>")}</p>
    <p>Pagamento via PIX (52.903.822/0001-86) 5 dias após a emissão da NF.</p>
    <h2>Prazo para Execução:</h2>
    <ul><li>${escapeHtml(prazoExecucao)}</li></ul>
    ${observacoes ? `<h2>Observações:</h2><p>${escapeHtml(observacoes).replace(/\n/g, "<br/>")}</p>` : ""}
    <div class="assinatura">ALPHA ENGENHARIA E SERVIÇOS</div>
    <div class="footer">
      <div class="endereco"><strong>ALPHA ENGENHARIA E SERVIÇOS</strong>Rua José Da Costa, 116 - São João Batista<br/>Belo Horizonte<br/>Telefone: 31 9 9203-1783</div>
      <div class="pagina">Página 3 de 3</div>
    </div>
  </section>
  <script>
    window.onload = () => {
      const images = Array.from(document.images || []);
      const loaded = images.map((img) => img.complete ? Promise.resolve() : new Promise((resolve) => {
        img.onload = resolve;
        img.onerror = resolve;
      }));
      Promise.all(loaded).finally(() => setTimeout(() => window.print(), 450));
    };
  </script>
</body>
</html>`;

  const printWindow = window.open("", "_blank");
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

  const FatorBdiGeralBase = calcularFatorBdiQualquer(bdi || BDI_PADRAO);
  const faturamentoDireto = !!bdi?.faturamentoDireto;
  const materiaisFaturamentoDireto = Array.isArray(bdi?.materiaisFaturamentoDireto)
    ? bdi.materiaisFaturamentoDireto
    : undefined;
  const FatorBdiMateriaisBase =
    faturamentoDireto && bdi?.materiais
      ? calcularFatorBdiQualquer(bdi.materiais)
      : FatorBdiGeralBase;
  const retencaoInss = Math.min(0.99, Math.max(0, num(bdi?.retencaoInss)));
  const FatorBdiMaoObraBase = retencaoInss > 0
    ? FatorBdiGeralBase / (1 - retencaoInss)
    : FatorBdiGeralBase;
  const collemAtivo = !!bdi?.collemAtivo;
  const collemX = num(bdi?.collemX) > 0 ? num(bdi.collemX) : 1;
  const collemY = num(bdi?.collemY) > 0 ? num(bdi.collemY) : 1;
  const divisorCollem = collemAtivo ? collemX * collemY : 1;
  const FatorBdiGeral = FatorBdiGeralBase / divisorCollem;
  const FatorBdiMateriais = FatorBdiMateriaisBase / divisorCollem;
  const FatorBdiMaoObra = FatorBdiMaoObraBase / divisorCollem;

  let totalCustoDireto = 0;
  let custoMaoObra = 0;
  let custoMateriaisFaturamentoDireto = 0;
  let totalPrecoVenda = 0;
  let totalPrecoVendaBase = 0;

  (etapas || []).forEach((e) => {
    itensAtivosDaEtapa(e).forEach((it) => {
      const qtdCpu = num(it.quantidade);
      (it.insumos || []).forEach((ins) => {
        const custoInsumoTotal = num(ins.coeficiente) * qtdCpu * insumoValorUnitario(ins, cpus, catalogMap);
        totalCustoDireto += custoInsumoTotal;

        if (insumoComFaturamentoDireto(ins, { faturamentoDireto, materiaisFaturamentoDireto })) {
          custoMateriaisFaturamentoDireto += custoInsumoTotal;
          totalPrecoVendaBase += custoInsumoTotal * FatorBdiMateriaisBase;
          totalPrecoVenda += custoInsumoTotal * FatorBdiMateriais;
        } else if (insumoEhMaoDeObra(ins.tipo)) {
          custoMaoObra += custoInsumoTotal;
          totalPrecoVendaBase += custoInsumoTotal * FatorBdiMaoObraBase;
          totalPrecoVenda += custoInsumoTotal * FatorBdiMaoObra;
        } else {
          totalPrecoVendaBase += custoInsumoTotal * FatorBdiGeralBase;
          totalPrecoVenda += custoInsumoTotal * FatorBdiGeral;
        }
      });
    });
  });

  const totalDiValor = totalPrecoVenda - totalCustoDireto;
  const totalDiRate = totalCustoDireto > 0 ? totalDiValor / totalCustoDireto : 0;
  const retencaoInssValorBase = custoMaoObra * (FatorBdiMaoObraBase - FatorBdiGeralBase);
  const retencaoInssValor = retencaoInssValorBase / divisorCollem;
  const custoBaseBdiGeral = totalCustoDireto - custoMateriaisFaturamentoDireto;

  return {
    bdiRate: FatorBdiGeralBase - 1,
    bdiRateMateriais: FatorBdiMateriaisBase - 1,
    bdiRateMaoObra: FatorBdiMaoObraBase - 1,
    FatorBdi: FatorBdiGeral,
    FatorBdiMateriais,
    FatorBdiMaoObra,
    FatorBdiBase: FatorBdiGeralBase,
    FatorBdiMateriaisBase,
    FatorBdiMaoObraBase,
    retencaoInss,
    custoMaoObra,
    custoBaseBdiGeral,
    custoMateriaisFaturamentoDireto,
    retencaoInssValor,
    retencaoInssValorBase,
    faturamentoDireto,
    materiaisFaturamentoDireto,
    collemAtivo,
    collemX,
    collemY,
    divisorCollem,
    totalDiValor,
    totalDiRate,
    custoDireto: totalCustoDireto,
    valorVendaBase: totalPrecoVendaBase,
    valorVenda: totalPrecoVenda,
  };
};

const calcularComparativosAlternativas = (etapas, bdi, cpus, catalogMap) =>
  (etapas || []).flatMap((etapa) =>
    gruposAlternativasDaEtapa(etapa).map((grupo) => ({
      etapaId: etapa.id,
      etapaNome: etapa.nome,
      grupoId: grupo.id,
      grupoNome: grupo.nome,
      opcaoAtivaId: grupo.opcaoAtivaId,
      opcoes: (grupo.opcoes || []).map((opcao) => {
        const calculo = calcularPrecoVendaProjeto(
          etapasComOpcaoAtiva(etapas, etapa.id, grupo.id, opcao.id),
          bdi,
          cpus,
          catalogMap
        );
        return {
          ...opcao,
          selecionada: opcao.id === grupo.opcaoAtivaId,
          custoDireto: calculo.custoDireto,
          valorVenda: calculo.valorVenda,
        };
      }),
    }))
  );

export default function App() {
  const [tab, setTab] = useState("projetos");
  const [cpus, setCpusState] = useState([]);
  const [projetos, setProjetos] = useState([]);
  const [clientes, setClientes] = useState([]);
  const [projetoAtivoId, setProjetoAtivoId] = useState("");
  const [loaded, setLoaded] = useState(false);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");
  const [driveConnected, setDriveConnected] = useState(false);
  const fileInputRef = useRef(null);
  const cpuHashesRef = useRef({});
  const projectHashesRef = useRef({});
  const legacyPrecosRef = useRef([]);
  const dadosAtuaisRef = useRef({ cpus: [], projetos: [], clientes: [], projetoAtivoId: "" });
  const [cpusDirty, setCpusDirty] = useState(false);
  const [abaPendenteAposSalvarCpus, setAbaPendenteAposSalvarCpus] = useState(null);
  // Novos estados para controle de recolhimento/expansão das camadas
  const [etapasExpandidas, setEtapasExpandidas] = useState({});
  const [cpusExpandidas, setCpusExpandidas] = useState({});
  const [buscaProjetos, setBuscaProjetos] = useState("");
  const [filtroStatusProjetos, setFiltroStatusProjetos] = useState("todos");
  const [ordenacaoProjetos, setOrdenacaoProjetos] = useState("recentes");
  const [paginaProjetos, setPaginaProjetos] = useState(1);
  const [projetoParaExcluir, setProjetoParaExcluir] = useState(null);
  const [textoConfirmacaoExclusao, setTextoConfirmacaoExclusao] = useState("");

  dadosAtuaisRef.current = { cpus, projetos, clientes, projetoAtivoId };

  const setCpus = (nextCpus) => {
    setCpusDirty(true);
    setCpusState(nextCpus);
  };

  const aplicarDadosCarregados = (data) => {
    if (data.empty) {
      const defaultBase = createDefaultProject();
      const defaultProj = {
        ...defaultBase,
        numeracaoPropostaVersao: VERSAO_NUMERACAO_PROPOSTAS,
        clienteCadastro: {
          ...(defaultBase.clienteCadastro || {}),
          numeroProposta: proximoNumeroProposta([]),
        },
      };
      const cpusIniciais = seedCpus();
      const snapshotInicial = {
        ...data,
        cpus: cpusIniciais,
        projetos: [defaultProj],
        clientes: [],
        projetoAtivoId: defaultProj.id,
      };
      setCpusState(cpusIniciais);
      setCpusDirty(true);
      setProjetos([defaultProj]);
      setClientes([]);
      projectHashesRef.current = {};
      legacyPrecosRef.current = [];
      setProjetoAtivoId(defaultProj.id);
      setStatus("Nenhum dado salvo no Firebase. Projeto inicial criado localmente.");
      return { snapshot: snapshotInicial, idsNumerados: [] };
    }

    const precosLegados = Array.isArray(data.precos) ? data.precos : [];
    const {
      projetos: projetosPreparados,
      projetoMigrado,
    } = prepararBancosDePrecosPorProjeto(data.projetos || [], precosLegados);
    const {
      projetos: projetosNumerados,
      idsMigrados: idsNumerados,
    } = numerarProjetosExistentes(projetosPreparados);
    const cadastroCompartilhado = prepararClientesCompartilhados(
      projetosNumerados,
      data.clientes || []
    );
    const snapshotPreparado = {
      ...data,
      projetos: cadastroCompartilhado.projetos,
      clientes: cadastroCompartilhado.clientes,
    };

    setCpusState(data.cpus || []);
    cpuHashesRef.current = data.cpuHashes || {};
    setCpusDirty(false);
    setProjetos(cadastroCompartilhado.projetos);
    setClientes(cadastroCompartilhado.clientes);
    projectHashesRef.current = Object.fromEntries(
      cadastroCompartilhado.projetos.map((project) => [project.id, JSON.stringify(project)])
    );
    legacyPrecosRef.current = precosLegados;
    setProjetoAtivoId(data.projetoAtivoId || "");
    setStatus(
      idsNumerados.length > 0
        ? `${idsNumerados.length} orçamento(s) receberam numeração sequencial de proposta.`
        : projetoMigrado && precosLegados.length > 0
        ? `Banco de Preços atual vinculado ao orçamento "${projetoMigrado.nome}".`
        : "Dados carregados do Firebase."
    );
    return { snapshot: snapshotPreparado, idsNumerados };
  };

  const carregarDados = async ({ usarDrive = true } = {}) => {
    setBusy(true);
    setStatus("Carregando...");
    try {
      const driveData = usarDrive ? await loadGoogleDriveSnapshot() : null;
      if (usarDrive && driveData) {
        const dadosAplicados = aplicarDadosCarregados(driveData);
        await saveLocalSnapshot(dadosAplicados.snapshot);
        if (dadosAplicados.idsNumerados.length > 0) {
          await executarSalvamentoAteConseguir("a numeração das propostas", () =>
            saveGoogleDriveSnapshot(dadosAplicados.snapshot, {
              includeBase: false,
              includeClients: true,
              projectIds: dadosAplicados.idsNumerados,
            })
          );
        }
        setDriveConnected(true);
        setStatus(
          dadosAplicados.idsNumerados.length > 0
            ? `${dadosAplicados.idsNumerados.length} proposta(s) numerada(s) e salva(s) no Google Drive.`
            : "Dados carregados do Google Drive."
        );
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

  const executarSalvamentoAteConseguir = async (descricao, operacao) => {
    let tentativa = 0;
    while (true) {
      tentativa += 1;
      if (tentativa > 1) {
        setStatus(`Tentativa automática ${tentativa}: salvando ${descricao}...`);
      }
      try {
        return await operacao();
      } catch (error) {
        const intervalo = intervaloNovaTentativa(tentativa);
        console.error(`Falha ao salvar ${descricao} (tentativa ${tentativa}):`, error);
        setStatus(
          `Falha ao salvar ${descricao}. Nova tentativa automática em ${Math.ceil(intervalo / 1000)}s (tentativa ${tentativa + 1}).`
        );
        await aguardar(intervalo);
      }
    }
  };

  const salvarProjeto = async (projectId) => {
    const projectInicial = dadosAtuaisRef.current.projetos.find(
      (item) => item.id === projectId
    );
    if (!projectInicial) return false;

    setBusy(true);
    setStatus(`Salvando o orçamento "${projectInicial.nome}"...`);
    try {
      const projectSalvo = await executarSalvamentoAteConseguir(
        `o orçamento "${projectInicial.nome}"`,
        async () => {
          const dados = dadosAtuaisRef.current;
          const projectAtual = dados.projetos.find((item) => item.id === projectId);
          if (!projectAtual) {
            throw new Error("O orçamento não está mais disponível.");
          }

          const atualizadoEm = new Date().toISOString();
          const projetosAtualizados = dados.projetos.map((item) =>
            item.id === projectId ? { ...item, atualizadoEm } : item
          );
          const projectAtualizado = projetosAtualizados.find(
            (item) => item.id === projectId
          );

          await saveLocalSnapshot({
            cpus: dados.cpus,
            projetos: projetosAtualizados,
            clientes: dados.clientes,
            precos: legacyPrecosRef.current,
            projetoAtivoId: dados.projetoAtivoId,
          });
          await saveGoogleDriveSnapshot(
            {
              cpus: dados.cpus,
              projetos: projetosAtualizados,
              clientes: dados.clientes,
              precos: legacyPrecosRef.current,
              projetoAtivoId: dados.projetoAtivoId,
            },
            {
              includeBase: false,
              includeClients: true,
              projectIds: [projectId],
            }
          );
          return projectAtualizado;
        }
      );
      setProjetos((prev) =>
        prev.map((item) =>
          item.id === projectId
            ? { ...item, atualizadoEm: projectSalvo.atualizadoEm }
            : item
        )
      );
      projectHashesRef.current[projectId] = JSON.stringify(projectSalvo);
      setDriveConnected(true);
      setStatus(
        cpusDirty
          ? `Orçamento "${projectSalvo.nome}" e seu Banco de Preços salvos. A Base de CPUs ainda possui alterações não salvas.`
          : `Orçamento "${projectSalvo.nome}" e seu Banco de Preços salvos no Google Drive.`
      );
      return true;
    } finally {
      setBusy(false);
      setTimeout(() => setStatus(""), 12000);
    }
  };

  const salvarBaseGeral = async () => {
    setBusy(true);
    setStatus("Salvando Base de CPUs no Google Drive...");
    try {
      await executarSalvamentoAteConseguir("a Base de CPUs", async () => {
        const dados = dadosAtuaisRef.current;
        await saveLocalSnapshot({
          cpus: dados.cpus,
          projetos: dados.projetos,
          clientes: dados.clientes,
          precos: legacyPrecosRef.current,
          projetoAtivoId: dados.projetoAtivoId,
        });
        await saveGoogleDriveSnapshot(
          {
            cpus: dados.cpus,
            projetos: dados.projetos,
            clientes: dados.clientes,
            precos: legacyPrecosRef.current,
            projetoAtivoId: dados.projetoAtivoId,
          },
          {
            includeBase: true,
            includeClients: true,
            projectIds: [],
          }
        );
      });
      setCpusDirty(false);
      setStatus("Base de CPUs salva no Google Drive.");
      return true;
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
      setStatus("Google Drive conectado. Carregando orçamentos...");
      const driveData = await loadGoogleDriveSnapshot();
      if (driveData) {
        const dadosAplicados = aplicarDadosCarregados(driveData);
        await saveLocalSnapshot(dadosAplicados.snapshot);
        if (dadosAplicados.idsNumerados.length > 0) {
          await executarSalvamentoAteConseguir("a numeração das propostas", () =>
            saveGoogleDriveSnapshot(dadosAplicados.snapshot, {
              includeBase: false,
              includeClients: true,
              projectIds: dadosAplicados.idsNumerados,
            })
          );
        }
        setStatus(
          dadosAplicados.idsNumerados.length > 0
            ? `Drive conectado. ${dadosAplicados.idsNumerados.length} proposta(s) numerada(s) e salva(s).`
            : "Drive conectado e orçamentos carregados."
        );
      } else {
        setStatus("Drive conectado. Ainda não há orçamentos salvos nesta conta.");
      }
      setDriveConnected(true);
    } catch (e) {
      setStatus("Falha ao conectar ou carregar o Drive: " + (e?.message || e));
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
  const modeloPropostaAtivo = clienteAtivo.modeloProposta || "";
  const cadastroClienteOk = useMemo(() => clienteEstaCompleto(clienteAtivo), [clienteAtivo]);
  const etapas = useMemo(() => projetoAtivo?.etapas || [], [projetoAtivo]);
  const bdi = useMemo(() => projetoAtivo?.bdi || BDI_PADRAO, [projetoAtivo]);
  const precos = useMemo(() => projetoAtivo?.precos || [], [projetoAtivo]);
  const cronograma = useMemo(
    () => ({
      dataInicio: projetoAtivo?.cronograma?.dataInicio || "",
      semanas: projetoAtivo?.cronograma?.semanas || 12,
      horasSemana: projetoAtivo?.cronograma?.horasSemana || 44,
      etapas: projetoAtivo?.cronograma?.etapas || {},
    }),
    [projetoAtivo]
  );
  const projetoAtivoDirty = useMemo(() => {
    if (!projetoAtivo) return false;
    return projectHashesRef.current[projetoAtivo.id] !== JSON.stringify(projetoAtivo);
  }, [projetoAtivo]);

  const setPrecos = (nextPrecos) => {
    if (!projetoAtivoId) return;
    setProjetos((prev) =>
      prev.map((projeto) => {
        if (projeto.id !== projetoAtivoId) return projeto;
        const precosAtuais = Array.isArray(projeto.precos) ? projeto.precos : [];
        const precosAtualizados =
          typeof nextPrecos === "function" ? nextPrecos(precosAtuais) : nextPrecos;
        return {
          ...projeto,
          precos: Array.isArray(precosAtualizados) ? precosAtualizados : [],
          bancoPrecosInicializado: true,
        };
      })
    );
  };

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

  const setCronograma = (novoCronograma) => {
    if (!projetoAtivoId) return;
    setProjetos((prev) =>
      prev.map((projeto) => {
        if (projeto.id !== projetoAtivoId) return projeto;
        const cronogramaAtual = {
          dataInicio: projeto.cronograma?.dataInicio || "",
          semanas: projeto.cronograma?.semanas || 12,
          horasSemana: projeto.cronograma?.horasSemana || 44,
          etapas: projeto.cronograma?.etapas || {},
        };
        return {
          ...projeto,
          cronograma:
            typeof novoCronograma === "function"
              ? novoCronograma(cronogramaAtual)
              : novoCronograma,
        };
      })
    );
  };

  const executarMudancaAba = (proximaTab, validarCadastroProjeto = false) => {
    if (validarCadastroProjeto && proximaTab !== "cliente" && !cadastroClienteOk) {
      setTab("cliente");
      setStatus("Preencha Nome do cliente e Local da obra para continuar.");
      setTimeout(() => setStatus(""), 5000);
      return;
    }
    setTab(proximaTab);
  };

  const solicitarMudancaAba = (proximaTab, validarCadastroProjeto = false) => {
    if (proximaTab === tab) return;
    if (tab === "cpus" && cpusDirty) {
      setAbaPendenteAposSalvarCpus({
        tab: proximaTab,
        validarCadastroProjeto,
      });
      return;
    }
    executarMudancaAba(proximaTab, validarCadastroProjeto);
  };

  const abrirAbaProjeto = (proximaTab) => {
    solicitarMudancaAba(proximaTab, true);
  };

  const salvarBaseEContinuar = async () => {
    const destino = abaPendenteAposSalvarCpus;
    if (!destino) return;
    const salvou = await salvarBaseGeral();
    if (!salvou) return;
    setAbaPendenteAposSalvarCpus(null);
    executarMudancaAba(destino.tab, destino.validarCadastroProjeto);
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

  const solicitarExclusaoProjeto = (project) => {
    if (projetos.length <= 1) {
      alert("Não é possível apagar todos os orçamentos.");
      return;
    }
    setProjetoParaExcluir(project);
    setTextoConfirmacaoExclusao("");
  };

  const fecharConfirmacaoExclusao = () => {
    if (busy) return;
    setProjetoParaExcluir(null);
    setTextoConfirmacaoExclusao("");
  };

  const removerProjeto = async (project) => {
    if (!project || projetos.length <= 1) return;
    setBusy(true);
    setStatus(`Excluindo o orçamento "${project.nome}"...`);
    try {
      await deleteGoogleDriveProject(project.id);
      const nextProjects = projetos.filter((item) => item.id !== project.id);
      const nextActiveId = project.id === projetoAtivoId
        ? nextProjects[0]?.id || ""
        : projetoAtivoId;
      setProjetos(nextProjects);
      setProjetoAtivoId(nextActiveId);
      delete projectHashesRef.current[project.id];
      await saveLocalSnapshot({
        cpus,
        projetos: nextProjects,
        clientes,
        precos: legacyPrecosRef.current,
        projetoAtivoId: nextActiveId,
      });
      setStatus("Orçamento excluído do Google Drive.");
      setProjetoParaExcluir(null);
      setTextoConfirmacaoExclusao("");
    } catch (error) {
      setStatus("Falha ao excluir o orçamento: " + (error?.message || error));
    } finally {
      setBusy(false);
      setTimeout(() => setStatus(""), 8000);
    }
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
      itensAtivosDaEtapa(etapa).forEach((item) => {
        const qtdItem = num(item.quantidade);
        (item.insumos || []).forEach((insumo) => {
          if (insumoEhMaterial(insumo.tipo)) {
            const nomeMat = (insumo.descricao || "").toUpperCase().trim();
            if (!nomeMat) return;

            // Usa a mesma resolução do orçamento: preço direto para insumos simples
            // e recálculo completo para subcomposições, inclusive em múltiplos níveis.
            const precoUnit = insumoValorUnitario(insumo, cpus, catalogMap);

            const qtdTotal = num(insumo.coeficiente) * qtdItem;
            const custoTotal = qtdTotal * precoUnit;

            if (!resumoMAT[nomeMat]) {
              resumoMAT[nomeMat] = {
                material: insumo.descricao,
                chave: precoKey(insumo.descricao),
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
  }, [etapas, cpus, catalogMap]);

  const chavesMateriaisOrcamento = useMemo(
    () => processarMateriais.map((material) => material.chave),
    [processarMateriais]
  );

  const definirMateriaisFaturamentoDireto = (chaves) => {
    setBdi((prev) => ({
      ...prev,
      materiaisFaturamentoDireto: Array.from(new Set(chaves)),
    }));
  };

  const alternarMaterialFaturamentoDireto = (chave, marcado) => {
    setBdi((prev) => {
      const atuais = Array.isArray(prev.materiaisFaturamentoDireto)
        ? prev.materiaisFaturamentoDireto
        : chavesMateriaisOrcamento;
      const proximos = new Set(atuais);
      if (marcado) proximos.add(chave);
      else proximos.delete(chave);
      return { ...prev, materiaisFaturamentoDireto: Array.from(proximos) };
    });
  };

  const grandTotal = useMemo(() => {
    return etapas.reduce(
      (s, e) =>
        s +
        itensAtivosDaEtapa(e).reduce(
          (s2, it) =>
            s2 + num(it.quantidade) * cpuValorUnit(it.insumos, cpus, catalogMap),
          0
        ),
      0
    );
  }, [etapas, cpus, catalogMap]);

  const bdiCalc = useMemo(() => {
    const calculo = calcularPrecoVendaProjeto(etapas, bdi, cpus, catalogMap);
    const descontoNegociacao = Math.min(
      calculo.valorVenda,
      descontoNegociacaoProjeto(projetoAtivo)
    );
    const valorVenda = aplicarDescontoNegociacao(
      calculo.valorVenda,
      descontoNegociacao
    );
    const totalDiValor = valorVenda - calculo.custoDireto;
    return {
      ...calculo,
      valorVendaBruto: calculo.valorVenda,
      descontoNegociacao,
      valorVenda,
      totalDiValor,
      totalDiRate:
        calculo.custoDireto > 0 ? totalDiValor / calculo.custoDireto : 0,
    };
  }, [bdi, etapas, cpus, catalogMap, projetoAtivo]);

  const comparativosAlternativas = useMemo(
    () =>
      calcularComparativosAlternativas(etapas, bdi, cpus, catalogMap).map(
        (grupo) => ({
          ...grupo,
          opcoes: grupo.opcoes.map((opcao) => ({
            ...opcao,
            valorVenda: aplicarDescontoNegociacao(
              opcao.valorVenda,
              bdiCalc.descontoNegociacao
            ),
          })),
        })
      ),
    [etapas, bdi, cpus, catalogMap, bdiCalc.descontoNegociacao]
  );

  const projetosComResumo = useMemo(
    () =>
      projetos.map((projeto, indiceOriginal) => {
        const cliente = clienteDoProjeto(projeto);
        const precosProjetoMap = new Map(
          (projeto.precos || []).map((preco) => [
            precoKey(preco.descricao),
            preco,
          ])
        );
        const custoDireto = (projeto.etapas || []).reduce(
          (totalEtapa, etapa) =>
            totalEtapa +
            itensAtivosDaEtapa(etapa).reduce(
              (totalItem, item) =>
                totalItem +
                num(item.quantidade) *
                  cpuValorUnit(item.insumos, cpus, precosProjetoMap),
              0
            ),
          0
        );
        const valorVendaBruto = calcularPrecoVendaProjeto(
          projeto.etapas || [],
          projeto.bdi || BDI_PADRAO,
          cpus,
          precosProjetoMap
        ).valorVenda;
        const descontoNegociacao = Math.min(
          valorVendaBruto,
          descontoNegociacaoProjeto(projeto)
        );
        const valorVenda = aplicarDescontoNegociacao(
          valorVendaBruto,
          descontoNegociacao
        );
        const statusProjeto = obterStatusProjeto(projeto);

        return {
          projeto,
          cliente,
          custoDireto,
          valorVendaBruto,
          descontoNegociacao,
          valorVenda,
          statusProjeto,
          indiceOriginal,
          dataOrdenacao:
            Date.parse(projeto.atualizadoEm || projeto.criadoEm || "") || 0,
          textoBusca: normalizarBusca(
            `${projeto.nome || ""} ${cliente.numeroProposta || ""} ${cliente.nome || ""} ${cliente.local || ""} ${cliente.endereco || ""}`
          ),
        };
      }),
    [projetos, cpus]
  );

  const resumoValoresPorStatus = useMemo(() => {
    const totais = {
      enviado_cliente: 0,
      em_elaboracao: 0,
      aprovado: 0,
      reprovado: 0,
    };
    projetosComResumo.forEach((resumo) => {
      if (Object.prototype.hasOwnProperty.call(totais, resumo.statusProjeto.id)) {
        totais[resumo.statusProjeto.id] += resumo.valorVenda;
      }
    });
    return totais;
  }, [projetosComResumo]);

  const projetosFiltrados = useMemo(() => {
    const termos = normalizarBusca(buscaProjetos)
      .split(/\s+/)
      .filter(Boolean);
    const filtrados = projetosComResumo.filter((resumo) => {
      const correspondeBusca = termos.every((termo) =>
        resumo.textoBusca.includes(termo)
      );
      const correspondeStatus =
        filtroStatusProjetos === "todos" ||
        resumo.statusProjeto.id === filtroStatusProjetos;
      return correspondeBusca && correspondeStatus;
    });

    return [...filtrados].sort((a, b) => {
      if (ordenacaoProjetos === "nome") {
        return String(a.projeto.nome || "").localeCompare(
          String(b.projeto.nome || ""),
          "pt-BR",
          { sensitivity: "base" }
        );
      }
      if (ordenacaoProjetos === "maior_valor") {
        return b.valorVenda - a.valorVenda;
      }
      if (a.dataOrdenacao !== b.dataOrdenacao) {
        return b.dataOrdenacao - a.dataOrdenacao;
      }
      return a.indiceOriginal - b.indiceOriginal;
    });
  }, [
    projetosComResumo,
    buscaProjetos,
    filtroStatusProjetos,
    ordenacaoProjetos,
  ]);

  const projetosPorPagina = 10;
  const totalPaginasProjetos = Math.max(
    1,
    Math.ceil(projetosFiltrados.length / projetosPorPagina)
  );
  const paginaProjetosAtual = Math.min(
    paginaProjetos,
    totalPaginasProjetos
  );
  const projetosPaginados = projetosFiltrados.slice(
    (paginaProjetosAtual - 1) * projetosPorPagina,
    paginaProjetosAtual * projetosPorPagina
  );

  useEffect(() => {
    setPaginaProjetos(1);
  }, [buscaProjetos, filtroStatusProjetos, ordenacaoProjetos]);

  useEffect(() => {
    if (paginaProjetos > totalPaginasProjetos) {
      setPaginaProjetos(totalPaginasProjetos);
    }
  }, [paginaProjetos, totalPaginasProjetos]);

  const abrirProjetoDaLista = (projeto, cliente) => {
    setProjetoAtivoId(projeto.id);
    setTab(clienteEstaCompleto(cliente) ? "custo" : "cliente");
  };

  const atualizarStatusProjeto = async (projectId, novoStatus) => {
    const statusPermitidos = [
      "rascunho",
      "em_elaboracao",
      "enviado_cliente",
      "aprovado",
      "reprovado",
      "cancelado",
      "concluido",
    ];
    if (!statusPermitidos.includes(novoStatus)) return;

    const dadosAtuais = dadosAtuaisRef.current;
    const projetosAtualizados = dadosAtuais.projetos.map((projeto) =>
      projeto.id === projectId
        ? { ...projeto, status: novoStatus }
        : projeto
    );
    dadosAtuaisRef.current = {
      ...dadosAtuais,
      projetos: projetosAtualizados,
    };
    setProjetos(projetosAtualizados);
    setStatus("Status alterado. Salvando automaticamente no Google Drive...");
    await salvarProjeto(projectId);
  };

  const atualizarDescontoProjeto = (projectId, valor) => {
    const dadosAtuais = dadosAtuaisRef.current;
    const projetosAtualizados = dadosAtuais.projetos.map((projeto) =>
      projeto.id === projectId
        ? { ...projeto, descontoNegociacao: valor }
        : projeto
    );
    dadosAtuaisRef.current = {
      ...dadosAtuais,
      projetos: projetosAtualizados,
    };
    setProjetos(projetosAtualizados);
  };

  // Abas disponíveis apenas dentro de um projeto ativo
  const abasProjeto = ["cliente", "custo", "planilha", "bdi", "precovenda", "cronograma", "histograma", "maoobra", "materiais", "precos"];
  const tabEhDeProjeto = abasProjeto.includes(tab);

  return (
    <div className="min-h-screen bg-stone-50 text-stone-900">
      {!driveConnected && (
        <div
          className="fixed inset-0 z-[100] bg-stone-950/55 backdrop-blur-sm flex items-center justify-center p-4"
          role="dialog"
          aria-modal="true"
          aria-labelledby="drive-required-title"
        >
          <div className="w-full max-w-sm bg-white border border-stone-200 rounded-lg shadow-2xl p-6 text-center">
            <img
              src={alphaLogo}
              alt="Alpha Engenharia"
              className="w-20 h-20 object-cover rounded-sm mx-auto mb-4"
            />
            <h2 id="drive-required-title" className="text-lg font-semibold text-stone-900">
              Conexão com o Google Drive
            </h2>
            <p className="mt-2 text-sm text-stone-500">
              Conecte a conta Google da equipe para carregar os orçamentos e acessar o aplicativo.
            </p>
            <button
              type="button"
              onClick={conectarGoogleDrive}
              disabled={busy}
              className="mt-5 w-full flex items-center justify-center gap-2 px-4 py-2.5 text-sm font-medium bg-stone-900 text-white rounded-lg hover:bg-stone-800 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {busy ? <RefreshCw size={16} className="animate-spin" /> : <LogIn size={16} />}
              {busy ? "Conectando e carregando..." : "Conectar ao Google Drive"}
            </button>
            {status && (
              <p className={`mt-3 text-xs ${status.startsWith("Falha") ? "text-red-600" : "text-stone-500"}`}>
                {status}
              </p>
            )}
          </div>
        </div>
      )}

      {projetoParaExcluir && (
        <div
          className="fixed inset-0 z-[95] bg-stone-950/60 backdrop-blur-sm flex items-center justify-center p-4"
          role="dialog"
          aria-modal="true"
          aria-labelledby="confirmar-exclusao-title"
        >
          <div className="w-full max-w-md bg-white border border-stone-200 rounded-lg shadow-2xl p-6">
            <div className="flex items-start gap-3">
              <div className="w-10 h-10 shrink-0 rounded-md bg-red-50 text-red-600 flex items-center justify-center">
                <Trash2 size={19} />
              </div>
              <div className="min-w-0">
                <h2
                  id="confirmar-exclusao-title"
                  className="text-base font-semibold text-stone-900"
                >
                  Excluir orçamento?
                </h2>
                <p className="mt-1.5 text-sm text-stone-500">
                  Esta ação excluirá permanentemente o orçamento do Google Drive.
                </p>
              </div>
            </div>

            <div className="mt-5 p-3 bg-stone-50 border border-stone-200 rounded-md">
              <p className="text-[10px] font-semibold uppercase text-stone-400">
                Orçamento selecionado
              </p>
              <p className="mt-1 text-sm font-semibold text-stone-800 uppercase break-words">
                {projetoParaExcluir.nome || "Orçamento sem nome"}
              </p>
            </div>

            <label className="block mt-5 text-sm text-stone-600">
              Digite <strong className="text-stone-900">EXCLUIR</strong> para
              confirmar:
              <input
                type="text"
                value={textoConfirmacaoExclusao}
                onChange={(e) => setTextoConfirmacaoExclusao(e.target.value)}
                autoFocus
                autoComplete="off"
                spellCheck={false}
                className="mt-2 w-full h-10 px-3 text-sm uppercase border border-stone-300 rounded-md outline-none focus:border-red-500 focus:ring-1 focus:ring-red-500"
                placeholder="EXCLUIR"
              />
            </label>

            <div className="mt-6 flex flex-col-reverse sm:flex-row sm:justify-end gap-2">
              <button
                type="button"
                onClick={fecharConfirmacaoExclusao}
                disabled={busy}
                className="px-4 py-2 text-sm border border-stone-300 rounded-md bg-white text-stone-700 hover:bg-stone-50 disabled:opacity-50"
              >
                Cancelar
              </button>
              <button
                type="button"
                onClick={() => removerProjeto(projetoParaExcluir)}
                disabled={
                  busy ||
                  normalizarBusca(textoConfirmacaoExclusao) !== "excluir"
                }
                className="flex items-center justify-center gap-2 px-4 py-2 text-sm font-medium bg-red-600 text-white rounded-md hover:bg-red-700 disabled:opacity-40 disabled:cursor-not-allowed"
              >
                {busy ? (
                  <RefreshCw size={15} className="animate-spin" />
                ) : (
                  <Trash2 size={15} />
                )}
                {busy ? "Excluindo..." : "Excluir definitivamente"}
              </button>
            </div>
          </div>
        </div>
      )}

      {abaPendenteAposSalvarCpus && (
        <div
          className="fixed inset-0 z-[90] bg-stone-950/55 backdrop-blur-sm flex items-center justify-center p-4"
          role="dialog"
          aria-modal="true"
          aria-labelledby="salvar-base-cpus-title"
        >
          <div className="w-full max-w-md bg-white border border-stone-200 rounded-lg shadow-2xl p-6">
            <div className="flex items-start gap-3">
              <div className="w-9 h-9 shrink-0 rounded-md bg-amber-100 text-amber-800 flex items-center justify-center">
                <AlertTriangle size={18} />
              </div>
              <div>
                <h2 id="salvar-base-cpus-title" className="text-base font-semibold text-stone-900">
                  Salve a Base de CPUs antes de sair
                </h2>
                <p className="mt-1.5 text-sm text-stone-500">
                  Existem alterações na base compartilhada. O acesso à próxima aba será liberado depois que elas forem salvas no Google Drive.
                </p>
              </div>
            </div>
            <div className="mt-5 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setAbaPendenteAposSalvarCpus(null)}
                disabled={busy}
                className="px-3 py-2 text-sm border border-stone-300 rounded-md bg-white text-stone-700 hover:bg-stone-50 disabled:opacity-50"
              >
                Continuar editando
              </button>
              <button
                type="button"
                onClick={salvarBaseEContinuar}
                disabled={busy}
                className="flex items-center gap-2 px-3 py-2 text-sm font-medium bg-stone-900 text-white rounded-md hover:bg-stone-800 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {busy ? <RefreshCw size={15} className="animate-spin" /> : <Save size={15} />}
                {busy ? "Salvando..." : "Salvar e continuar"}
              </button>
            </div>
          </div>
        </div>
      )}

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
              <SideTabBtn active={tab === "projetos"} onClick={() => solicitarMudancaAba("projetos")} icon={<FolderKanban size={15} />}>
                Orçamentos ({projetos.length})
              </SideTabBtn>
              <SideTabBtn active={tab === "cpus"} onClick={() => solicitarMudancaAba("cpus")} icon={<Database size={15} />}>
                Base de CPUs ({cpus.length})
              </SideTabBtn>

              {projetoAtivo && (
                <>
                  <div className="max-lg:hidden border-t border-stone-200 my-2 pt-2">
                    <p className="px-2 text-[10px] font-semibold text-stone-400 uppercase truncate">
                      {projetoAtivo.nome}
                    </p>
                  </div>
                  <SideTabBtn active={tab === "cliente"} onClick={() => abrirAbaProjeto("cliente")} icon={<User size={15} />}>
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
                    BDI
                  </SideTabBtn>
                  <SideTabBtn active={tab === "precovenda"} onClick={() => abrirAbaProjeto("precovenda")} icon={<TrendingUp size={15} />}>
                    Venda
                  </SideTabBtn>
                  <SideTabBtn active={tab === "cronograma"} onClick={() => abrirAbaProjeto("cronograma")} icon={<CalendarDays size={15} />}>
                    Cronograma
                  </SideTabBtn>
                  <SideTabBtn active={tab === "histograma"} onClick={() => abrirAbaProjeto("histograma")} icon={<BarChart3 size={15} />}>
                    Histograma
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
            <div className="max-lg:hidden border-t border-stone-200 px-3 py-3">
              <p className="text-[10px] font-semibold uppercase text-stone-400 mb-2">
                Resumo dos orçamentos
              </p>
              <div className="space-y-1.5 text-[10px]">
                <div className="flex items-center justify-between gap-2">
                  <span className="text-amber-700">Enviados ao cliente</span>
                  <span className="font-mono font-semibold text-stone-700 whitespace-nowrap">
                    R$ {fmt(resumoValoresPorStatus.enviado_cliente)}
                  </span>
                </div>
                <div className="flex items-center justify-between gap-2">
                  <span className="text-sky-700">Em elaboração</span>
                  <span className="font-mono font-semibold text-stone-700 whitespace-nowrap">
                    R$ {fmt(resumoValoresPorStatus.em_elaboracao)}
                  </span>
                </div>
                <div className="flex items-center justify-between gap-2">
                  <span className="text-emerald-700">Aprovados</span>
                  <span className="font-mono font-semibold text-stone-700 whitespace-nowrap">
                    R$ {fmt(resumoValoresPorStatus.aprovado)}
                  </span>
                </div>
                <div className="flex items-center justify-between gap-2">
                  <span className="text-red-700">Reprovados</span>
                  <span className="font-mono font-semibold text-stone-700 whitespace-nowrap">
                    R$ {fmt(resumoValoresPorStatus.reprovado)}
                  </span>
                </div>
              </div>
            </div>
            {projetoAtivo && tabEhDeProjeto && (
              <div className="p-2 border-t border-stone-200">
                <button
                  type="button"
                  onClick={() => salvarProjeto(projetoAtivo.id)}
                  disabled={busy || !loaded}
                  className="w-full flex items-center justify-center gap-2 px-3 py-2 text-xs font-medium bg-stone-900 text-white rounded-md hover:bg-stone-800 disabled:opacity-50 disabled:cursor-not-allowed"
                  title={`Salvar o orçamento ${projetoAtivo.nome} e seu Banco de Preços`}
                >
                  <Save size={14} /> {busy ? "Salvando..." : "Salvar orçamento"}
                </button>
              </div>
            )}
          </div>
        </aside>

        <main className="min-w-0 flex-1">
        {/* HEADER */}
        <header className="flex items-center justify-between mb-6 gap-4">
          <div className="min-w-0 flex items-center gap-3">
            <img
              src={alphaLogo}
              alt="Alpha Engenharia"
              className="w-14 h-14 rounded-sm object-cover shrink-0"
            />
            <div className="min-w-0">
              <h1 className="text-2xl font-semibold tracking-tight">Orçamentador por CPU</h1>
              <p className="text-sm text-stone-500 truncate">
                {projetoAtivo
                  ? `Orçamento: ${projetoAtivo.nome}  -  ${clienteAtivo.nome || "Cliente não cadastrado"}`
                  : "Crie ou selecione um orçamento para começar"}
              </p>
            </div>
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
                  setProjetos((prev) => {
                    const numeroProposta = proximoNumeroProposta(prev);
                    return [
                      ...prev,
                      {
                      id: pId,
                      nome: `Novo Orçamento - ${prev.length + 1}`,
                      numeracaoPropostaVersao: VERSAO_NUMERACAO_PROPOSTAS,
                      criadoEm: new Date().toISOString(),
                      atualizadoEm: "",
                      cliente: "",
                      clienteCadastro: { ...CLIENTE_PADRAO, numeroProposta },
                      etapas: [{ id: uid(), nome: "Etapa Inicial", itens: [] }],
                      precos: [],
                      bancoPrecosInicializado: true,
                      cronograma: {
                        dataInicio: "",
                        semanas: 12,
                        horasSemana: 44,
                        etapas: {},
                      },
                      bdi: {
                        custoInicial: 0,
                        admCentral: 0,
                        contabilidade: 0,
                        contingenciamento: 0,
                        custoFinanceiro: 0,
                        dasAnexoIV: 0,
                        art: 0,
                        retencaoInss: 0,
                        lucro: 0,
                        faturamentoDireto: false,
                        collemAtivo: false,
                        collemX: 1,
                        collemY: 1,
                        materiais: { admCentral: 0, contabilidade: 0, contingenciamento: 0, custoFinanceiro: 0, lucro: 0, dasAnexoIV: 0, art: 0 }
                      }
                      },
                    ];
                  });
                  setProjetoAtivoId(pId);
                  setTab("cliente");
                }}
                className="flex items-center gap-1.5 px-3 py-1.5 bg-stone-900 text-white rounded-lg text-xs font-medium hover:bg-stone-800"
              >
                <Plus size={14} /> Novo Orçamento
              </button>
            </div>

            <div className="bg-white border border-stone-200 rounded-lg">
              <div className="p-4 border-b border-stone-200 flex flex-col lg:flex-row lg:items-center gap-3">
                <div className="relative flex-1 min-w-0">
                  <Search
                    size={15}
                    className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none"
                  />
                  <input
                    type="search"
                    value={buscaProjetos}
                    onChange={(e) => setBuscaProjetos(e.target.value)}
                    placeholder="Buscar orçamento, cliente ou local..."
                    className="w-full h-10 pl-9 pr-3 text-sm bg-white border border-stone-300 rounded-md outline-none focus:border-stone-600"
                  />
                </div>
                <select
                  value={filtroStatusProjetos}
                  onChange={(e) => setFiltroStatusProjetos(e.target.value)}
                  aria-label="Filtrar orçamentos por status"
                  className="h-10 px-3 text-sm bg-white border border-stone-300 rounded-md outline-none focus:border-stone-600 lg:w-48"
                >
                  <option value="todos">Todos os status</option>
                  <option value="em_elaboracao">Em elaboração</option>
                  <option value="cadastro_pendente">Cadastro pendente</option>
                  <option value="rascunho">Rascunho</option>
                  <option value="concluido">Concluído</option>
                  <option value="enviado_cliente">Enviado p/ cliente</option>
                  <option value="cancelado">Cancelado</option>
                  <option value="aprovado">Aprovado</option>
                  <option value="reprovado">Reprovado</option>
                </select>
                <div className="relative lg:w-44">
                  <ArrowUpDown
                    size={14}
                    className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none"
                  />
                  <select
                    value={ordenacaoProjetos}
                    onChange={(e) => setOrdenacaoProjetos(e.target.value)}
                    aria-label="Ordenar orçamentos"
                    className="w-full h-10 pl-9 pr-3 text-sm bg-white border border-stone-300 rounded-md outline-none focus:border-stone-600 appearance-none"
                  >
                    <option value="recentes">Mais recentes</option>
                    <option value="nome">Ordem alfabética</option>
                    <option value="maior_valor">Maior valor</option>
                  </select>
                  <ChevronDown
                    size={14}
                    className="absolute right-3 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none"
                  />
                </div>
              </div>

              {projetosPaginados.length === 0 ? (
                <div className="py-16 px-4 text-center">
                  <Search size={30} className="mx-auto text-stone-300 mb-3" />
                  <p className="text-sm font-medium text-stone-600">
                    Nenhum orçamento encontrado
                  </p>
                  <p className="text-xs text-stone-400 mt-1">
                    Ajuste a busca ou o filtro para visualizar outros projetos.
                  </p>
                </div>
              ) : (
                <>
                  <div className="hidden xl:block overflow-x-auto">
                    <div className="min-w-[1060px]">
                      <div className="grid grid-cols-[88px_minmax(145px,1.2fr)_minmax(165px,1.3fr)_90px_105px_115px_80px_105px_108px] gap-3 px-5 py-3 bg-stone-50 border-b border-stone-200 text-[10px] font-semibold uppercase text-stone-500">
                        <div>Número</div>
                        <div>Orçamento</div>
                        <div>Cliente / Local</div>
                        <div className="text-right">Desconto</div>
                        <div className="text-right">Custo direto</div>
                        <div className="text-right">Preço de venda</div>
                        <div>Atualizado</div>
                        <div>Status</div>
                        <div className="text-right">Ações</div>
                      </div>

                      {projetosPaginados.map((resumo) => {
                        const {
                          projeto,
                          cliente,
                          custoDireto,
                          descontoNegociacao,
                          valorVenda,
                          statusProjeto,
                        } = resumo;
                        const isActive = projeto.id === projetoAtivoId;
                        const isSentToClient = statusProjeto.id === "enviado_cliente";

                        return (
                          <div
                            key={projeto.id}
                            role="button"
                            tabIndex={0}
                            onClick={() => abrirProjetoDaLista(projeto, cliente)}
                            onKeyDown={(e) => {
                              if (e.key === "Enter" || e.key === " ") {
                                e.preventDefault();
                                abrirProjetoDaLista(projeto, cliente);
                              }
                            }}
                            className={`grid grid-cols-[88px_minmax(145px,1.2fr)_minmax(165px,1.3fr)_90px_105px_115px_80px_105px_108px] gap-3 px-5 py-4 border-b border-stone-200 last:border-b-0 items-center cursor-pointer outline-none transition-colors border-l-4 ${
                              isSentToClient
                                ? "border-l-amber-600 bg-yellow-300 hover:bg-yellow-200 focus:bg-yellow-200"
                                : isActive
                                ? "border-l-[#6f9255] bg-[#f3f7ef]"
                                : "border-l-transparent hover:bg-stone-50 focus:bg-stone-50"
                            }`}
                          >
                            <div className="text-xs font-mono font-semibold text-stone-700 whitespace-nowrap">
                              {cliente.numeroProposta || "Pendente"}
                            </div>
                            <div className="min-w-0">
                              <p className="text-xs font-semibold text-stone-900 uppercase truncate">
                                {projeto.nome || "Orçamento sem nome"}
                              </p>
                              <p className="text-[11px] text-stone-400 mt-1 truncate">
                                {cliente.nome || "Cliente não cadastrado"}
                              </p>
                            </div>
                            <div className="min-w-0">
                              <p className="text-xs text-stone-700 truncate">
                                {cliente.nome || "Cliente não cadastrado"}
                              </p>
                              <p className="text-[11px] text-stone-400 mt-1 flex items-center gap-1 min-w-0">
                                <MapPin size={11} className="shrink-0" />
                                <span className="truncate">
                                  {cliente.local || "Local da obra pendente"}
                                </span>
                              </p>
                            </div>
                            <div>
                              <input
                                type="text"
                                inputMode="decimal"
                                value={projeto.descontoNegociacao ?? ""}
                                onClick={(e) => e.stopPropagation()}
                                onKeyDown={(e) => e.stopPropagation()}
                                onChange={(e) => atualizarDescontoProjeto(projeto.id, e.target.value)}
                                onBlur={() => salvarProjeto(projeto.id)}
                                placeholder="0,00"
                                aria-label={`Desconto do orçamento ${projeto.nome}`}
                                className="w-full h-8 px-2 text-right text-xs font-mono bg-white border border-stone-300 rounded-md outline-none focus:border-stone-600"
                              />
                              {descontoNegociacao > 0 && (
                                <p className="mt-1 text-[9px] text-right text-red-600 font-mono">
                                  - R$ {fmt(descontoNegociacao)}
                                </p>
                              )}
                            </div>
                            <div className="text-xs text-right font-mono text-stone-600 whitespace-nowrap">
                              R$ {fmt(custoDireto)}
                            </div>
                            <div className="text-sm text-right font-mono font-bold text-stone-900 whitespace-nowrap">
                              R$ {fmt(valorVenda)}
                            </div>
                            <div className="text-[11px] text-stone-500 whitespace-nowrap">
                              {formatarAtualizacaoProjeto(projeto.atualizadoEm)}
                            </div>
                            <div>
                              <select
                                value={statusProjeto.id}
                                onClick={(e) => e.stopPropagation()}
                                onKeyDown={(e) => e.stopPropagation()}
                                onChange={(e) => {
                                  e.stopPropagation();
                                  atualizarStatusProjeto(projeto.id, e.target.value);
                                }}
                                disabled={busy || statusProjeto.id === "cadastro_pendente"}
                                aria-label={`Alterar status do orçamento ${projeto.nome}`}
                                className={`w-full h-8 px-2 rounded border text-[10px] font-medium outline-none disabled:cursor-not-allowed ${statusProjeto.className}`}
                              >
                                {statusProjeto.id === "cadastro_pendente" ? (
                                  <option value="cadastro_pendente">Cadastro pendente</option>
                                ) : (
                                  <>
                                    <option value="rascunho">Rascunho</option>
                                    <option value="em_elaboracao">Em elaboração</option>
                                    <option value="enviado_cliente">Enviado p/ cliente</option>
                                    <option value="aprovado">Aprovado</option>
                                    <option value="reprovado">Reprovado</option>
                                    <option value="cancelado">Cancelado</option>
                                    <option value="concluido">Concluído</option>
                                  </>
                                )}
                              </select>
                            </div>
                            <div className="flex justify-end gap-1.5">
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  abrirProjetoDaLista(projeto, cliente);
                                }}
                                className="w-8 h-8 inline-flex items-center justify-center border border-stone-300 rounded-md bg-white text-stone-600 hover:bg-stone-100"
                                title={`Abrir o orçamento ${projeto.nome}`}
                              >
                                <ChevronRight size={15} />
                              </button>
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  salvarProjeto(projeto.id);
                                }}
                                disabled={busy || !loaded}
                                className="w-8 h-8 inline-flex items-center justify-center border border-stone-300 rounded-md bg-white text-stone-600 hover:bg-stone-100 disabled:opacity-50"
                                title={`Salvar o orçamento ${projeto.nome}`}
                              >
                                <Save size={14} />
                              </button>
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  solicitarExclusaoProjeto(projeto);
                                }}
                                disabled={busy}
                                className="w-8 h-8 inline-flex items-center justify-center border border-stone-300 rounded-md bg-white text-stone-500 hover:border-red-200 hover:bg-red-50 hover:text-red-600 disabled:opacity-50"
                                title={`Excluir o orçamento ${projeto.nome}`}
                              >
                                <Trash2 size={14} />
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  <div className="xl:hidden">
                    {projetosPaginados.map((resumo) => {
                      const {
                        projeto,
                        cliente,
                        custoDireto,
                        descontoNegociacao,
                        valorVenda,
                        statusProjeto,
                      } = resumo;
                      const isActive = projeto.id === projetoAtivoId;
                      const isSentToClient = statusProjeto.id === "enviado_cliente";

                      return (
                        <div
                          key={projeto.id}
                          className={`p-4 border-b border-stone-200 last:border-b-0 border-l-4 ${
                            isSentToClient
                              ? "border-l-amber-600 bg-yellow-300"
                              : isActive
                              ? "border-l-[#6f9255] bg-[#f3f7ef]"
                              : "border-l-transparent"
                          }`}
                        >
                          <div className="flex items-start justify-between gap-3">
                            <button
                              type="button"
                              onClick={() => abrirProjetoDaLista(projeto, cliente)}
                              className="min-w-0 flex-1 text-left"
                            >
                              <div className="min-w-0">
                                <p className="text-sm font-semibold text-stone-900 uppercase truncate">
                                  {projeto.nome || "Orçamento sem nome"}
                                </p>
                                <p className="text-xs text-stone-500 mt-1 truncate">
                                  {cliente.numeroProposta
                                    ? `${cliente.numeroProposta} · ${cliente.nome || "Cliente não cadastrado"}`
                                    : cliente.nome || "Cliente não cadastrado"}
                                </p>
                                <p className="text-[11px] text-stone-400 mt-1 flex items-center gap-1">
                                  <MapPin size={11} className="shrink-0" />
                                  <span className="truncate">
                                    {cliente.local || "Local da obra pendente"}
                                  </span>
                                </p>
                              </div>
                            </button>
                            <select
                              value={statusProjeto.id}
                              onChange={(e) =>
                                atualizarStatusProjeto(projeto.id, e.target.value)
                              }
                              disabled={busy || statusProjeto.id === "cadastro_pendente"}
                              aria-label={`Alterar status do orçamento ${projeto.nome}`}
                              className={`shrink-0 max-w-36 h-8 px-2 rounded border text-[10px] font-medium outline-none disabled:cursor-not-allowed ${statusProjeto.className}`}
                            >
                              {statusProjeto.id === "cadastro_pendente" ? (
                                <option value="cadastro_pendente">Cadastro pendente</option>
                              ) : (
                                <>
                                  <option value="rascunho">Rascunho</option>
                                  <option value="em_elaboracao">Em elaboração</option>
                                  <option value="enviado_cliente">Enviado p/ cliente</option>
                                  <option value="aprovado">Aprovado</option>
                                  <option value="reprovado">Reprovado</option>
                                  <option value="cancelado">Cancelado</option>
                                  <option value="concluido">Concluído</option>
                                </>
                              )}
                            </select>
                          </div>

                          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-4 pt-3 border-t border-stone-200/70">
                            <div>
                              <p className="text-[9px] uppercase text-stone-400">Desconto</p>
                              <input
                                type="text"
                                inputMode="decimal"
                                value={projeto.descontoNegociacao ?? ""}
                                onChange={(e) => atualizarDescontoProjeto(projeto.id, e.target.value)}
                                onBlur={() => salvarProjeto(projeto.id)}
                                placeholder="0,00"
                                aria-label={`Desconto do orçamento ${projeto.nome}`}
                                className="mt-1 w-full h-8 px-2 text-right text-xs font-mono bg-white border border-stone-300 rounded-md outline-none focus:border-stone-600"
                              />
                              {descontoNegociacao > 0 && (
                                <p className="mt-1 text-[9px] text-red-600 font-mono whitespace-nowrap">
                                  - R$ {fmt(descontoNegociacao)}
                                </p>
                              )}
                            </div>
                            <div>
                              <p className="text-[9px] uppercase text-stone-400">Custo direto</p>
                              <p className="text-xs font-mono text-stone-700 mt-0.5 whitespace-nowrap">
                                R$ {fmt(custoDireto)}
                              </p>
                            </div>
                            <div>
                              <p className="text-[9px] uppercase text-stone-400">Preço de venda</p>
                              <p className="text-xs font-mono font-bold text-stone-900 mt-0.5 whitespace-nowrap">
                                R$ {fmt(valorVenda)}
                              </p>
                            </div>
                            <div>
                              <p className="text-[9px] uppercase text-stone-400">Atualizado</p>
                              <p className="text-xs text-stone-600 mt-0.5 whitespace-nowrap">
                                {formatarAtualizacaoProjeto(projeto.atualizadoEm)}
                              </p>
                            </div>
                          </div>

                          <div className="flex justify-between items-center mt-4">
                            {isActive ? (
                              <span className="text-[10px] font-semibold uppercase text-[#56713f]">
                                Orçamento ativo
                              </span>
                            ) : (
                              <span />
                            )}
                            <div className="flex gap-1.5">
                              <button
                                type="button"
                                onClick={() => abrirProjetoDaLista(projeto, cliente)}
                                className="w-9 h-9 inline-flex items-center justify-center border border-stone-300 rounded-md bg-white text-stone-600"
                                title={`Abrir o orçamento ${projeto.nome}`}
                              >
                                <ChevronRight size={16} />
                              </button>
                              <button
                                type="button"
                                onClick={() => salvarProjeto(projeto.id)}
                                disabled={busy || !loaded}
                                className="w-9 h-9 inline-flex items-center justify-center border border-stone-300 rounded-md bg-white text-stone-600 disabled:opacity-50"
                                title={`Salvar o orçamento ${projeto.nome}`}
                              >
                                <Save size={15} />
                              </button>
                              <button
                                type="button"
                                onClick={() => solicitarExclusaoProjeto(projeto)}
                                disabled={busy}
                                className="w-9 h-9 inline-flex items-center justify-center border border-stone-300 rounded-md bg-white text-stone-500 disabled:opacity-50"
                                title={`Excluir o orçamento ${projeto.nome}`}
                              >
                                <Trash2 size={15} />
                              </button>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </>
              )}

              <div className="px-4 py-3 border-t border-stone-200 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
                <p className="text-xs text-stone-500">
                  {projetosFiltrados.length === 0
                    ? "0 orçamentos"
                    : `Mostrando ${(paginaProjetosAtual - 1) * projetosPorPagina + 1}-${Math.min(
                        paginaProjetosAtual * projetosPorPagina,
                        projetosFiltrados.length
                      )} de ${projetosFiltrados.length} orçamento(s)`}
                </p>
                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={() =>
                      setPaginaProjetos((pagina) => Math.max(1, pagina - 1))
                    }
                    disabled={paginaProjetosAtual <= 1}
                    className="w-8 h-8 inline-flex items-center justify-center border border-stone-300 rounded-md bg-white text-stone-600 hover:bg-stone-50 disabled:opacity-35 disabled:cursor-not-allowed"
                    title="Página anterior"
                  >
                    <ChevronLeft size={15} />
                  </button>
                  <span className="min-w-8 h-8 px-2 inline-flex items-center justify-center border border-[#8eaa78] rounded-md text-xs font-semibold text-[#56713f] bg-white">
                    {paginaProjetosAtual}
                  </span>
                  <button
                    type="button"
                    onClick={() =>
                      setPaginaProjetos((pagina) =>
                        Math.min(totalPaginasProjetos, pagina + 1)
                      )
                    }
                    disabled={paginaProjetosAtual >= totalPaginasProjetos}
                    className="w-8 h-8 inline-flex items-center justify-center border border-stone-300 rounded-md bg-white text-stone-600 hover:bg-stone-50 disabled:opacity-35 disabled:cursor-not-allowed"
                    title="Próxima página"
                  >
                    <ChevronRight size={15} />
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
        {tab === "cpus" && (
          <CpuLibrary
            cpus={cpus}
            setCpus={setCpus}
            fileInputRef={fileInputRef}
            catalogMap={catalogMap}
            onSaveBase={salvarBaseGeral}
            saving={busy}
            baseDirty={cpusDirty}
          />
        )}

        {/* Abas de projeto - só renderizam se houver projeto ativo */}
        {tabEhDeProjeto && !projetoAtivo && (
          <div className="text-center py-20 text-stone-400">
            <FolderKanban size={40} className="mx-auto mb-3 opacity-30" />
            <p className="text-sm font-medium">Nenhum orçamento selecionado.</p>
            <button onClick={() => solicitarMudancaAba("projetos")} className="mt-3 text-xs underline">
              Criar ou selecionar um orçamento
            </button>
          </div>
        )}
        {tab === "cliente" && projetoAtivo && (
          <CadastroCliente
            projeto={projetoAtivo}
            cliente={clienteAtivo}
            clientes={clientes}
            setProjetos={setProjetos}
            setClientes={setClientes}
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
                      data.push([`${idxE + 1}`, etapa.nome, "", "", "", itensAtivosDaEtapa(etapa).reduce((acc, it) => acc + (num(it.quantidade) * cpuValorUnit(it.insumos, cpus, catalogMap)), 0)]);
                      itensAtivosDaEtapa(etapa).forEach((item, idxI) => {
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
                            R$ {fmt(itensAtivosDaEtapa(etapa).reduce((acc, it) => acc + (num(it.quantidade) * cpuValorUnit(it.insumos, cpus, catalogMap)), 0))}
                          </span>
                        </div>

                        {isEtapaAberta && itensAtivosDaEtapa(etapa).map((item, idxItem) => {
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
            <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3 border border-stone-200 bg-stone-50 px-4 py-3 rounded-lg">
              <div>
                <p className="text-xs font-semibold text-stone-800">Modelo da proposta</p>
                <p className="text-[11px] text-stone-500 mt-0.5">
                  Selecione a empresa antes de gerar os documentos comerciais.
                </p>
              </div>
              <div className="inline-flex w-full sm:w-auto border border-stone-300 rounded-md overflow-hidden bg-white" role="radiogroup" aria-label="Modelo da proposta">
                <button
                  type="button"
                  role="radio"
                  aria-checked={modeloPropostaAtivo === "collem"}
                  onClick={() => setClienteAtivo((prev) => ({ ...prev, modeloProposta: "collem" }))}
                  className={`flex-1 sm:flex-none px-4 py-2 text-xs font-semibold border-r border-stone-300 ${
                    modeloPropostaAtivo === "collem"
                      ? "bg-[#126594] text-white"
                      : "bg-white text-stone-600 hover:bg-stone-50"
                  }`}
                >
                  COLLEM
                </button>
                <button
                  type="button"
                  role="radio"
                  aria-checked={modeloPropostaAtivo === "alpha"}
                  onClick={() => setClienteAtivo((prev) => ({
                    ...prev,
                    modeloProposta: "alpha",
                    numeroProposta: prev.numeroProposta || proximoNumeroProposta(projetos),
                  }))}
                  className={`flex-1 sm:flex-none px-4 py-2 text-xs font-semibold ${
                    modeloPropostaAtivo === "alpha"
                      ? "bg-[#789654] text-white"
                      : "bg-white text-stone-600 hover:bg-stone-50"
                  }`}
                >
                  ALPHA ENGENHARIA
                </button>
              </div>
            </div>
            {!modeloPropostaAtivo && (
              <div className="flex items-center gap-2 px-3 py-2 border border-amber-300 bg-amber-50 text-amber-800 rounded-md text-xs">
                <AlertTriangle size={14} /> Escolha COLLEM ou ALPHA ENGENHARIA para liberar a geração da proposta.
              </div>
            )}
            <div className="flex justify-between items-center">
              <div>
                <h2 className="text-base font-semibold text-stone-800">Planilha de Preço de Venda</h2>
                <p className="text-xs text-stone-500">
                  Apresentação comercial por etapas e serviços. Valor total: <span className="font-bold text-stone-800 font-mono">R$ {fmt(bdiCalc.valorVenda)}</span>
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

                <button
                  onClick={() =>
                    exportarPropostaXlsx({
                      projeto: projetoAtivo,
                      cliente: clienteAtivo,
                      etapas,
                      bdiCalc,
                      cpus,
                      catalogMap,
                      modelo: modeloPropostaAtivo,
                    })
                  }
                  disabled={!modeloPropostaAtivo}
                  className="px-2 py-1 text-[11px] font-medium border border-stone-900 text-white bg-stone-900 rounded hover:bg-stone-800 flex items-center gap-1 disabled:opacity-40 disabled:cursor-not-allowed"
                >
                  <FileText size={12} /> {modeloPropostaAtivo === "collem" ? "Anexo I .xlsx" : "Proposta .xlsx"}
                </button>

                {modeloPropostaAtivo && (
                  <button
                    type="button"
                    onClick={async () => {
                      try {
                        if (modeloPropostaAtivo === "collem") {
                          const { gerarPropostaCollemDocx } = await import(
                            "./propostas/collemProposal"
                          );
                          await gerarPropostaCollemDocx({
                            projeto: projetoAtivo,
                            cliente: clienteAtivo,
                            totalGeral: bdiCalc.valorVenda,
                          });
                          return;
                        }

                        const numeroProposta = clienteAtivo.numeroProposta || proximoNumeroProposta(projetos);
                        if (!clienteAtivo.numeroProposta) {
                          setClienteAtivo((prev) => ({ ...prev, numeroProposta }));
                        }
                        const grupos = montarItensProposta(etapas, bdiCalc, cpus, catalogMap, clienteAtivo);
                        const totalGeral = bdiCalc.valorVenda;
                        const descricaoRegimeMateriais = materialPorContaCliente(clienteAtivo)
                          ? "Material por conta do cliente. A proposta considera somente os serviços, mão de obra, equipamentos e demais custos não classificados como material."
                          : materialFaturamentoDireto(clienteAtivo) || bdiCalc.faturamentoDireto
                            ? Array.isArray(bdiCalc.materiaisFaturamentoDireto)
                              ? "Somente os materiais selecionados no orçamento são considerados com faturamento direto para o cliente, aplicando o BDI específico configurado."
                              : "Materiais considerados com faturamento direto para o cliente, aplicando BDI específico de materiais quando configurado."
                            : "Materiais inclusos no fornecimento da ALPHA ENGENHARIA conforme composição do orçamento.";
                        const { gerarPropostaAlphaDocx } = await import(
                          "./propostas/alphaProposal"
                        );
                        await gerarPropostaAlphaDocx({
                          nomeProjeto: projetoAtivo?.nome || "Orçamento",
                          nomeCliente: clienteAtivo?.nome || "Cliente",
                          localObra: clienteAtivo?.local || clienteAtivo?.endereco || "",
                          contato: clienteAtivo?.contato || "",
                          numeroProposta,
                          grupos,
                          comparativos: montarComparativosProposta(etapas, bdiCalc, cpus, catalogMap, clienteAtivo),
                          totalGeral,
                          descontoNegociacao: bdiCalc.descontoNegociacao,
                          descricaoRegimeMateriais,
                          responsabilidadesAlpha: listaTextoOuPadrao(clienteAtivo?.responsabilidadesAlpha, RESPONSABILIDADES_ALPHA_PADRAO),
                          responsabilidadesCliente: listaTextoOuPadrao(clienteAtivo?.responsabilidadesCliente, RESPONSABILIDADES_CLIENTE_PADRAO),
                          condicoesPagamento: clienteAtivo?.condicoesPagamento || `Entrada de 40% (R$ ${fmt(totalGeral * 0.4)}) e o restante (R$ ${fmt(totalGeral * 0.6)}) conforme avanço dos serviços em medições.`,
                          prazoExecucao: clienteAtivo?.prazoExecucao || "A definir conforme cronograma aprovado entre as partes.",
                          observacoes: clienteAtivo?.observacoes || "",
                        });
                      } catch (erro) {
                        alert(`Falha ao gerar a proposta: ${erro?.message || erro}`);
                      }
                    }}
                    className={`px-2 py-1 text-[11px] font-medium border text-white rounded flex items-center gap-1 ${modeloPropostaAtivo === "collem" ? "border-[#126594] bg-[#126594] hover:bg-[#0d527a]" : "border-[#7B9A56] bg-[#7B9A56] hover:bg-[#698448]"}`}
                  >
                    <FileText size={12} /> Proposta .docx
                  </button>
                )}

                <button
                  onClick={async () => {
                    if (modeloPropostaAtivo === "collem") {
                      const { gerarPropostaCollemPdf } = await import(
                        "./propostas/collemProposal"
                      );
                      gerarPropostaCollemPdf({
                        projeto: projetoAtivo,
                        cliente: clienteAtivo,
                        totalGeral: bdiCalc.valorVenda,
                      });
                      return;
                    }
                    const numeroProposta = clienteAtivo.numeroProposta || proximoNumeroProposta(projetos);
                    if (!clienteAtivo.numeroProposta) {
                      setClienteAtivo((prev) => ({ ...prev, numeroProposta }));
                    }
                    gerarPropostaPdf({
                      projeto: projetoAtivo,
                      cliente: { ...clienteAtivo, numeroProposta },
                      etapas,
                      bdiCalc,
                      cpus,
                      catalogMap,
                      numeroPropostaAutomatico: numeroProposta,
                    });
                  }}
                  disabled={!modeloPropostaAtivo}
                  className="px-2 py-1 text-[11px] font-medium border border-red-200 text-red-700 bg-red-50/50 rounded hover:bg-red-50 flex items-center gap-1 disabled:opacity-40 disabled:cursor-not-allowed"
                >
                  <Download size={12} /> Proposta PDF
                </button>

              </div>
            </div>

            {comparativosAlternativas.length > 0 && (
              <div className="border border-amber-200 rounded-lg overflow-hidden">
                <div className="px-4 py-2.5 bg-amber-50 border-b border-amber-200">
                  <h3 className="text-xs font-semibold text-amber-900">
                    Comparativo de alternativas
                  </h3>
                </div>
                <div className="divide-y divide-amber-100">
                  {comparativosAlternativas.map((grupo) => {
                    const valorSelecionado =
                      grupo.opcoes.find((opcao) => opcao.selecionada)
                        ?.valorVenda || 0;
                    return (
                      <div
                        key={`${grupo.etapaId}-${grupo.grupoId}`}
                        className="p-3 bg-white"
                      >
                        <div className="mb-2">
                          <p className="text-[10px] uppercase text-stone-400">
                            {grupo.etapaNome}
                          </p>
                          <p className="text-xs font-semibold text-stone-800">
                            {grupo.grupoNome}
                          </p>
                        </div>
                        <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-2">
                          {grupo.opcoes.map((opcao) => {
                            const diferenca =
                              opcao.valorVenda - valorSelecionado;
                            return (
                              <button
                              key={opcao.id}
                              type="button"
                              onClick={() =>
                                setEtapas((atuais) =>
                                  atuais.map((etapa) =>
                                    etapa.id === grupo.etapaId
                                      ? etapaComOpcaoAtiva(
                                          etapa,
                                          grupo.grupoId,
                                          opcao.id
                                        )
                                      : etapa
                                  )
                                )
                              }
                              className={`text-left border rounded-md px-3 py-2 ${
                                opcao.selecionada
                                  ? "border-emerald-400 bg-emerald-50"
                                  : "border-stone-200 bg-white hover:bg-stone-50"
                              }`}
                            >
                              <span className="flex items-center justify-between gap-2">
                                <span className="text-[11px] font-semibold text-stone-800 truncate">
                                  {opcao.nome}
                                </span>
                                {opcao.selecionada && (
                                  <Check size={13} className="text-emerald-600 shrink-0" />
                                )}
                              </span>
                              <span className="mt-1 block font-mono text-sm font-semibold text-stone-900">
                                R$ {fmt(opcao.valorVenda)}
                              </span>
                              {!opcao.selecionada && (
                                <span className="block text-[10px] text-stone-400">
                                  {diferenca >= 0 ? "Acréscimo" : "Economia"}: R$ {fmt(Math.abs(diferenca))}
                                </span>
                              )}
                              </button>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

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
                    itensAtivosDaEtapa(etapa).forEach(it => {
                      const qCpu = num(it.quantidade);
                      (it.insumos || []).forEach(ins => {
                        const cIn = num(ins.coeficiente) * qCpu * insumoValorUnitario(ins, cpus, catalogMap);
                        totalEtapaComBdi += cIn * fatorVendaInsumo(ins, bdiCalc);
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

                        {isEtapaAberta && itensAtivosDaEtapa(etapa).map((item, idxItem) => {
                          const numCpu = `${numEtapa}.${idxItem + 1}`;
                          const itemId = item.id || `item-${numCpu}`;
                          const isCpuAberta = !!cpusExpandidas[itemId];
                          const qtdItem = num(item.quantidade);

                          let totalCpuComBdi = 0;
                          (item.insumos || []).forEach(ins => {
                            const cIn = num(ins.coeficiente) * insumoValorUnitario(ins, cpus, catalogMap);
                            totalCpuComBdi += cIn * fatorVendaInsumo(ins, bdiCalc);
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
                                    const custoUnit = insumoValorUnitario(insumo, cpus, catalogMap);
                                    
                                    const precoVendaInsumo = custoUnit * fatorVendaInsumo(insumo, bdiCalc);
                                    
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
                {bdiCalc.descontoNegociacao > 0 && (
                  <div className="grid grid-cols-12 gap-2 px-4 py-2.5 bg-amber-50 border-t border-amber-200 text-xs font-semibold uppercase text-amber-900">
                    <span className="col-span-6">Desconto da negociação</span>
                    <span className="col-span-4 text-right font-mono text-stone-500">
                      Venda bruta: R$ {fmt(bdiCalc.valorVendaBruto)}
                    </span>
                    <span className="col-span-2 text-right font-mono text-red-700">
                      - R$ {fmt(bdiCalc.descontoNegociacao)}
                    </span>
                  </div>
                )}
                <div className="grid grid-cols-12 gap-2 px-4 py-3 bg-stone-900 text-white text-sm font-semibold uppercase tracking-wider">
                  <span className="col-span-6">VALOR FINAL DE VENDA</span>
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

        {tab === "cronograma" && projetoAtivo && (
          <CronogramaSemanal
            projeto={projetoAtivo}
            etapas={etapas}
            cronograma={cronograma}
            setCronograma={setCronograma}
            bdiCalc={bdiCalc}
            cpus={cpus}
            catalogMap={catalogMap}
            cliente={clienteAtivo}
          />
        )}

        {tab === "histograma" && projetoAtivo && (
          <HistogramaMaoObra
            projeto={projetoAtivo}
            etapas={etapas}
            cronograma={cronograma}
            setCronograma={setCronograma}
            cpus={cpus}
          />
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
                  const listaMo = consolidarMaoDeObra(
                    etapas,
                    cpus,
                    catalogMap
                  );
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
            <div className="flex items-start justify-between gap-4 flex-wrap">
              <div>
                <h2 className="text-base font-semibold text-stone-800">Quantitativo de Materiais</h2>
                <p className="text-xs text-stone-500">Escolha individualmente quais materiais terão o BDI de faturamento direto.</p>
              </div>
              <label className="flex items-center gap-2 border border-stone-200 rounded-md px-3 py-2 text-xs font-medium text-stone-700 cursor-pointer bg-stone-50">
                <input
                  type="checkbox"
                  checked={!!bdi.faturamentoDireto}
                  onChange={(e) => setBdi((prev) => ({ ...prev, faturamentoDireto: e.target.checked }))}
                  className="w-4 h-4 accent-emerald-600"
                />
                Usar faturamento direto por material
              </label>
            </div>

            <div className={`flex items-center justify-between gap-3 flex-wrap rounded-md border px-3 py-2 ${bdi.faturamentoDireto ? "border-emerald-200 bg-emerald-50" : "border-stone-200 bg-stone-50"}`}>
              <span className="text-xs text-stone-600">
                {bdi.faturamentoDireto
                  ? `${processarMateriais.filter((material) => !Array.isArray(bdi.materiaisFaturamentoDireto) || bdi.materiaisFaturamentoDireto.includes(material.chave)).length} de ${processarMateriais.length} material(is) no faturamento direto`
                  : "Ative o faturamento direto para liberar a seleção individual."}
              </span>
              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={() => definirMateriaisFaturamentoDireto(chavesMateriaisOrcamento)}
                  disabled={!bdi.faturamentoDireto || processarMateriais.length === 0}
                  className="px-2.5 py-1.5 text-xs border border-stone-200 rounded-md bg-white text-stone-700 hover:bg-stone-100 disabled:opacity-40 disabled:cursor-not-allowed"
                >
                  Marcar todos
                </button>
                <button
                  type="button"
                  onClick={() => definirMateriaisFaturamentoDireto([])}
                  disabled={!bdi.faturamentoDireto || processarMateriais.length === 0}
                  className="px-2.5 py-1.5 text-xs border border-stone-200 rounded-md bg-white text-stone-700 hover:bg-stone-100 disabled:opacity-40 disabled:cursor-not-allowed"
                >
                  Desmarcar todos
                </button>
              </div>
            </div>

            <div className="border border-stone-200 rounded-lg overflow-x-auto">
              <div className="min-w-[820px]">
              <div className="grid grid-cols-12 gap-2 px-4 py-2.5 bg-stone-100 border-b border-stone-200 text-stone-500 font-semibold text-[11px] uppercase tracking-wider">
                <span className="col-span-5">Material</span>
                <span className="col-span-1 text-center">Und</span>
                <span className="col-span-1.5 text-right">Qtd Total</span>
                <span className="col-span-1.5 text-right">Preço Unit.</span>
                <span className="col-span-2 text-right">Total Bruto</span>
                <span className="col-span-1 text-center">Fat. direto</span>
              </div>

              <div className="divide-y divide-stone-200 max-h-[500px] overflow-y-auto">
                {processarMateriais.length === 0 ? (
                  <div className="p-8 text-center text-stone-400 italic text-xs">Nenhum material localizado neste orçamento.</div>
                ) : (
                  <>
                    {processarMateriais.map((material) => {
                      const marcado = !!bdi.faturamentoDireto && (
                        !Array.isArray(bdi.materiaisFaturamentoDireto) ||
                        bdi.materiaisFaturamentoDireto.includes(material.chave)
                      );
                      return (
                        <div key={material.chave} className={`grid grid-cols-12 gap-2 px-4 py-2 text-xs items-center uppercase ${marcado ? "bg-emerald-50/60" : "hover:bg-stone-50/60"}`}>
                          <span className="col-span-5 font-medium text-stone-800 truncate" title={material.material}>{material.material}</span>
                          <span className="col-span-1 text-center font-mono text-stone-400">{material.unidade}</span>
                          <span className="col-span-1.5 text-right font-mono text-stone-900">{material.quantidade.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 3 })}</span>
                          <span className="col-span-1.5 text-right font-mono text-stone-400">R$ {fmt(material.valorUnitario)}</span>
                          <span className="col-span-2 text-right font-mono font-semibold text-emerald-700">R$ {fmt(material.valorTotal)}</span>
                          <label className="col-span-1 flex justify-center cursor-pointer" title="Aplicar faturamento direto neste material">
                            <input
                              type="checkbox"
                              checked={marcado}
                              disabled={!bdi.faturamentoDireto}
                              onChange={(e) => alternarMaterialFaturamentoDireto(material.chave, e.target.checked)}
                              aria-label={`Faturamento direto: ${material.material}`}
                              className="w-4 h-4 accent-emerald-600 disabled:opacity-40"
                            />
                          </label>
                        </div>
                      );
                    })}

                    <div className="grid grid-cols-12 gap-2 px-4 py-3 bg-stone-900 text-white text-sm font-semibold">
                      <span className="col-span-5">TOTAL GERAL EM MATERIAIS</span>
                      <span className="col-span-1"></span>
                      <span className="col-span-1.5"></span>
                      <span className="col-span-1.5"></span>
                      <span className="col-span-2 text-right font-mono text-amber-400">
                        R$ {fmt(processarMateriais.reduce((acc, material) => acc + material.valorTotal, 0))}
                      </span>
                      <span className="col-span-1"></span>
                    </div>
                  </>
                )}
              </div>
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
            nomeProjeto={projetoAtivo.nome}
            onSaveProject={() => salvarProjeto(projetoAtivo.id)}
            saving={busy}
            projectDirty={projetoAtivoDirty}
          />
        )}
        </main>
      </div>
    </div>
  );
}

function CadastroCliente({ projeto, cliente, clientes, setProjetos, setClientes, setCliente, completo, onContinuar }) {
  const clientesOrdenados = useMemo(
    () => [...(clientes || [])].sort((a, b) => String(a.nome || "").localeCompare(String(b.nome || ""), "pt-BR")),
    [clientes]
  );

  const atualizarCampo = (campo, valor) => {
    if (!CAMPOS_CLIENTE_COMPARTILHADOS.includes(campo)) {
      setCliente((prev) => ({ ...prev, [campo]: valor }));
      return;
    }

    const clienteId = cliente.clienteId || uid();
    const atualizadoEm = new Date().toISOString();
    setClientes((atuais) => {
      const existente = atuais.find((item) => item.id === clienteId);
      const atualizado = {
        id: clienteId,
        ...dadosClienteCompartilhado(existente || cliente),
        [campo]: valor,
        atualizadoEm,
      };
      return existente
        ? atuais.map((item) => (item.id === clienteId ? atualizado : item))
        : [...atuais, atualizado];
    });
    setProjetos((atuais) =>
      atuais.map((item) => {
        const cadastroAtual = clienteDoProjeto(item);
        const projetoAtivo = item.id === projeto.id;
        if (!projetoAtivo && cadastroAtual.clienteId !== clienteId) return item;
        const clienteCadastro = {
          ...cadastroAtual,
          clienteId,
          [campo]: valor,
        };
        return {
          ...item,
          cliente: clienteCadastro.nome || "",
          clienteCadastro,
        };
      })
    );
  };

  const selecionarCliente = (clienteId) => {
    if (!clienteId) {
      setCliente((atual) => ({
        ...atual,
        clienteId: "",
        ...dadosClienteCompartilhado({}),
      }));
      return;
    }
    const selecionado = clientes.find((item) => item.id === clienteId);
    if (!selecionado) return;
    setCliente((atual) => ({
      ...atual,
      ...dadosClienteCompartilhado(selecionado),
      clienteId: selecionado.id,
    }));
  };

  const atualizarNomeProjeto = (valor) => {
    setProjetos((prev) => prev.map((p) => (p.id === projeto.id ? { ...p, nome: valor } : p)));
  };

  const atualizarRegimeMateriais = (valor) => {
    atualizarCampo("regimeMateriais", valor);
    setProjetos((prev) =>
      prev.map((p) => {
        if (p.id !== projeto.id) return p;
        const bdiAtual = p.bdi || BDI_PADRAO;
        return {
          ...p,
          bdi: {
            ...bdiAtual,
            faturamentoDireto: valor === "faturamentoDireto",
            materiais: bdiAtual.materiais || {
              admCentral: 0,
              contabilidade: 0,
              contingenciamento: 0,
              custoFinanceiro: 0,
              lucro: 0,
              dasAnexoIV: 0,
              art: 0,
            },
          },
        };
      })
    );
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
        <div className="rounded-lg border border-stone-200 bg-stone-50 p-4">
          <div className="grid grid-cols-1 md:grid-cols-[minmax(0,1fr)_auto] gap-3 items-end">
            <div>
              <label className="text-xs font-medium text-stone-600 mb-1 flex items-center gap-1.5">
                <Building2 size={14} /> Cliente compartilhado
              </label>
              <select
                value={cliente.clienteId || ""}
                onChange={(e) => selecionarCliente(e.target.value)}
                className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm outline-none bg-white focus:border-stone-700 focus:ring-1 focus:ring-stone-700"
              >
                <option value="">+ Cadastrar novo cliente</option>
                {clientesOrdenados.map((item) => (
                  <option key={item.id} value={item.id}>
                    {item.nome || "Cliente sem nome"}{item.documento ? ` — ${item.documento}` : ""}
                  </option>
                ))}
              </select>
            </div>
            <span className="text-[11px] text-stone-500 pb-2">
              {clientesOrdenados.length} cliente(s) disponível(is)
            </span>
          </div>
          <p className="text-[11px] text-stone-500 mt-2">
            Nome, contato, telefone, e-mail, CPF/CNPJ, CEP e endereço são compartilhados entre os orçamentos vinculados.
          </p>
        </div>

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
            onChange={(valor) => atualizarCampo("telefone", formatarTelefone(valor))}
            icon={<Phone size={14} />}
            inputMode="tel"
            maxLength={15}
            autoComplete="tel"
          />
          <CampoCliente
            label="E-mail"
            value={cliente.email || ""}
            onChange={(valor) => atualizarCampo("email", valor)}
            icon={<Mail size={14} />}
            type="email"
            autoComplete="email"
          />
          <CampoCliente
            label="CPF/CNPJ"
            value={cliente.documento || ""}
            onChange={(valor) => atualizarCampo("documento", formatarCpfCnpj(valor))}
            icon={<FileText size={14} />}
            inputMode="numeric"
            maxLength={18}
          />
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <CampoCliente
            label="CEP"
            value={cliente.cep || ""}
            onChange={(valor) => atualizarCampo("cep", formatarCep(valor))}
            icon={<MapPin size={14} />}
            inputMode="numeric"
            maxLength={9}
            autoComplete="postal-code"
          />
          <CampoCliente
            label="Endereço"
            value={cliente.endereco || ""}
            onChange={(valor) => atualizarCampo("endereco", valor)}
            icon={<MapPin size={14} />}
            autoComplete="street-address"
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

        <div className="border border-stone-200 rounded-lg p-4 space-y-4">
          <h3 className="text-sm font-semibold text-stone-800 flex items-center gap-2">
            <FileText size={15} /> Dados da proposta
          </h3>
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
            <CampoCliente
              label="Número da proposta (automático)"
              value={cliente.numeroProposta || ""}
              onChange={() => {}}
              icon={<FileText size={14} />}
              placeholder="Número gerado automaticamente"
              readOnly
            />
            <div>
              <label className="text-xs font-medium text-stone-500 mb-1 flex items-center gap-1.5">
                <FileText size={14} /> Materiais na proposta
              </label>
              <select
                value={cliente.regimeMateriais || "alpha"}
                onChange={(e) => atualizarRegimeMateriais(e.target.value)}
                className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm outline-none bg-white focus:border-stone-700 focus:ring-1 focus:ring-stone-700"
              >
                <option value="alpha">Material incluso pela Alpha</option>
                <option value="cliente">Material por conta do cliente</option>
                <option value="faturamentoDireto">Faturamento direto ao cliente</option>
              </select>
            </div>
            <CampoCliente
              label="Prazo para execução"
              value={cliente.prazoExecucao || ""}
              onChange={(valor) => atualizarCampo("prazoExecucao", valor)}
              icon={<FileText size={14} />}
              placeholder="Ex.: 60 dias úteis"
            />
            <CampoCliente
              label="Condições de pagamento"
              value={cliente.condicoesPagamento || ""}
              onChange={(valor) => atualizarCampo("condicoesPagamento", valor)}
              icon={<FileText size={14} />}
              placeholder="Ex.: Entrada de 40% e restante por medição"
            />
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <CampoTextoCliente
              label="Responsabilidade da Alpha"
              value={cliente.responsabilidadesAlpha || ""}
              onChange={(valor) => atualizarCampo("responsabilidadesAlpha", valor)}
              placeholder={RESPONSABILIDADES_ALPHA_PADRAO.join("\n")}
            />
            <CampoTextoCliente
              label="Responsabilidade do cliente"
              value={cliente.responsabilidadesCliente || ""}
              onChange={(valor) => atualizarCampo("responsabilidadesCliente", valor)}
              placeholder={RESPONSABILIDADES_CLIENTE_PADRAO.join("\n")}
            />
          </div>
          {cliente.modeloProposta === "collem" && (
            <div className="border-t border-stone-200 pt-4 space-y-4">
              <div>
                <h4 className="text-xs font-semibold text-[#126594]">Dados específicos da proposta COLLEM</h4>
                <p className="text-[11px] text-stone-500 mt-0.5">
                  Campos vazios usarão os textos padrão do modelo oficial.
                </p>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <CampoCliente
                  label="Prazo de execução COLLEM"
                  value={cliente.prazoExecucaoCollem || ""}
                  onChange={(valor) => atualizarCampo("prazoExecucaoCollem", valor)}
                  icon={<FileText size={14} />}
                  placeholder="Ex.: 60 dias úteis"
                />
                <CampoCliente
                  label="Sinal de negócio (%)"
                  value={cliente.percentualSinalCollem ?? 20}
                  onChange={(valor) => atualizarCampo("percentualSinalCollem", valor)}
                  icon={<Percent size={14} />}
                  type="number"
                  placeholder="20"
                />
                <CampoCliente
                  label="Responsável pela proposta"
                  value={cliente.responsavelCollem || ""}
                  onChange={(valor) => atualizarCampo("responsavelCollem", valor)}
                  icon={<User size={14} />}
                  placeholder="Geraldo Belloni Perez"
                />
              </div>
              <CampoTextoCliente
                label="Texto de apresentação COLLEM"
                value={cliente.textoApresentacaoCollem || ""}
                onChange={(valor) => atualizarCampo("textoApresentacaoCollem", valor)}
                placeholder="Atendendo à vossa solicitação, estamos encaminhando nossa proposta..."
              />
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <CampoTextoCliente
                  label="Não inclusos"
                  value={cliente.naoInclusosCollem || ""}
                  onChange={(valor) => atualizarCampo("naoInclusosCollem", valor)}
                  placeholder="Itens e fornecimentos que não fazem parte desta proposta"
                />
                <CampoTextoCliente
                  label="Condições especiais"
                  value={cliente.condicoesEspeciaisCollem || ""}
                  onChange={(valor) => atualizarCampo("condicoesEspeciaisCollem", valor)}
                  placeholder="Alojamento, alimentação, acessos e demais condições acordadas"
                />
              </div>
            </div>
          )}
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

function CampoCliente({ label, value, onChange, icon, required, type = "text", inputClassName, placeholder, readOnly = false, inputMode, maxLength, autoComplete }) {
  return (
    <div>
      <label className="text-xs font-medium text-stone-500 mb-1 flex items-center gap-1.5">
        {icon} {label} {required && <span className="text-amber-600">*</span>}
      </label>
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        readOnly={readOnly}
        inputMode={inputMode}
        maxLength={maxLength}
        autoComplete={autoComplete}
        className={inputClassName || `w-full border border-stone-300 rounded-lg px-3 py-2 text-sm outline-none ${
          readOnly
            ? "bg-stone-100 text-stone-600 cursor-not-allowed"
            : "bg-white focus:border-stone-700 focus:ring-1 focus:ring-stone-700"
        }`}
        placeholder={placeholder}
      />
    </div>
  );
}

function CampoTextoCliente({ label, value, onChange, placeholder }) {
  return (
    <div>
      <label className="text-xs font-medium text-stone-500 mb-1 flex items-center gap-1.5">
        <FileText size={14} /> {label}
      </label>
      <textarea
        value={value || ""}
        onChange={(e) => onChange(e.target.value)}
        rows={5}
        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm outline-none bg-white focus:border-stone-700 focus:ring-1 focus:ring-stone-700 resize-none"
        placeholder={placeholder}
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

function CpuLibrary({ cpus, setCpus, fileInputRef, catalogMap, onSaveBase, saving, baseDirty }) {
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
    () => criarIndiceBuscaCpus(cpus),
    [cpus]
  );

  const queryTokens = useMemo(() => {
    const tokens = [];
    const re = /"([^"]+)"|(\S+)/g;
    let m;
    while ((m = re.exec(query)) !== null) {
      const t = normalizarBusca((m[1] || m[2] || "").trim());
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

  const insumosFiltrados = useMemo(
    () => buscarInsumosCatalogo(catalogMap, query),
    [catalogMap, query]
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
          `Importação concluída: ${resultado.adicionadas} nova(s), ${resultado.atualizadas} atualizada(s), ${resultado.semMudanca} sem mudança. Clique em Salvar Base de CPUs para gravar na nuvem.`
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
        <button onClick={() => setEditing("new")} className="flex items-center gap-1.5 px-3 py-2 text-sm border border-stone-300 bg-white text-stone-700 rounded-lg hover:bg-stone-100">
          <Plus size={15} /> Nova CPU Base
        </button>
        <button
          type="button"
          onClick={onSaveBase}
          disabled={saving}
          className="flex items-center gap-1.5 px-3 py-2 text-sm bg-stone-900 text-white rounded-lg hover:bg-stone-700 disabled:opacity-50 disabled:cursor-not-allowed"
          title="Salvar a Base de CPUs compartilhada"
        >
          <Save size={15} /> {saving ? "Salvando..." : "Salvar Base de CPUs"}
          {baseDirty && <span className="w-1.5 h-1.5 rounded-full bg-amber-400" title="Alterações pendentes" />}
        </button>
      </div>

      {importMsg && <div className="mb-4 text-xs px-3 py-2 rounded-lg bg-amber-50 text-amber-800 border border-amber-200">{importMsg}</div>}

      {query.trim() && insumosFiltrados.length > 0 && (
        <div className="mb-4 border border-stone-200 rounded-lg overflow-hidden bg-white">
          <div className="px-3 py-2 bg-stone-100 border-b border-stone-200 flex items-center justify-between gap-3">
            <span className="text-xs font-semibold text-stone-700">Insumos encontrados</span>
            <span className="text-[10px] text-stone-400">Pesquisa pelo nome do insumo</span>
          </div>
          <div className="divide-y divide-stone-100 max-h-56 overflow-y-auto">
            {insumosFiltrados.map((insumo) => (
              <div key={insumo.key || insumo.id} className="grid grid-cols-[auto_1fr_auto] gap-3 items-center px-3 py-2 text-xs">
                <span className="px-1.5 py-0.5 bg-stone-100 text-stone-600 rounded font-medium text-[10px]">{insumo.tipo || "INS"}</span>
                <span className="text-stone-800 truncate" title={insumo.descricao}>{insumo.descricao}</span>
                <span className="font-mono text-stone-400">{insumo.unidade || "un"}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="space-y-2">
        {query.trim() && (
          <div className="flex items-center justify-between px-1 pb-1">
            <span className="text-xs font-semibold text-stone-700">CPUs encontradas pelo nome ou código</span>
            <span className="text-[10px] text-stone-400">{filtered.length} resultado(s)</span>
          </div>
        )}
        {query.trim() && filtered.length === 0 && (
          <div className="border border-stone-200 rounded-lg bg-white px-4 py-5 text-center text-xs text-stone-400">
            Nenhuma CPU possui essas palavras no nome ou código.
          </div>
        )}
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

  const renderSubCpuTree = (
    cpu,
    insumosCpu,
    onChangeInsumosCpu,
    nivel = 0,
    visited = new Set()
  ) => {
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
        {(insumosCpu || []).map((subInsumo, subIndex) => {
          const rowKey = `${cpu.id}-${subInsumo.id || subInsumo.codigo || subInsumo.descricao}-${nivel}`;
          const subSubCpu = findSubCpu(subInsumo, cpus);
          const estaAberta = !!subCpusExpandidas[rowKey];
          const subValor = insumoValorUnitario(subInsumo, cpus, catalogMap, nextVisited);
          const subTotal = num(subInsumo.coeficiente) * subValor;
          const podeEditarValor = !readOnly && !subSubCpu && onUpsertPreco;
          const podeEditarCoeficiente = !readOnly && !!onChangeInsumosCpu;
          const atualizarSubInsumo = (patch) => {
            if (!onChangeInsumosCpu) return;
            onChangeInsumosCpu(
              insumosCpu.map((item, index) =>
                index === subIndex ? { ...item, ...patch } : item
              )
            );
          };

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
                <span className="col-span-2 text-right font-mono text-stone-600">
                  {podeEditarCoeficiente ? (
                    <input
                      type="number"
                      step="any"
                      value={subInsumo.coeficiente ?? ""}
                      onChange={(e) => atualizarSubInsumo({ coeficiente: e.target.value })}
                      className="w-24 border border-amber-200 rounded px-1 py-0.5 text-right font-mono bg-white outline-none focus:ring-1 focus:ring-amber-500"
                      title="Coeficiente personalizado somente para este orçamento"
                    />
                  ) : (
                    fmt(subInsumo.coeficiente)
                  )}
                </span>
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
                    renderSubCpuTree(
                      subSubCpu,
                      insumosResolvidosSubCpu(subInsumo, subSubCpu),
                      (novosInsumos) => atualizarSubInsumo({ subCpuInsumos: novosInsumos }),
                      nivel + 1,
                      nextVisited
                    )
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
          const insumosSubCpu = subCpu ? insumosResolvidosSubCpu(i, subCpu) : [];
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
                {renderSubCpuTree(
                  subCpu,
                  insumosSubCpu,
                  (novosInsumos) => setMany(i.id, { subCpuInsumos: novosInsumos })
                )}
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
function PrecosTab({
  catalog,
  onUpsert,
  onRemove,
  onApplyToCpus,
  onApplyAllToCpus,
  nomeProjeto,
  onSaveProject,
  saving,
  projectDirty,
}) {
  const [editing, setEditing] = useState(null);
  const [query, setQuery] = useState("");
  const [sortConfig, setSortConfig] = useState({ key: "descricao", direction: "asc" });

  const hasPrice = (value) => value !== "" && value !== null && value !== undefined && Number.isFinite(num(value));

  const requestSort = (key) => {
    setSortConfig((prev) => {
      if (prev.key === key) {
        return { key, direction: prev.direction === "asc" ? "desc" : "asc" };
      }
      return { key, direction: "asc" };
    });
  };

  const SortIcon = ({ sortKey }) => {
    if (sortConfig.key !== sortKey) return <ArrowUpDown size={11} className="text-stone-300" />;
    return sortConfig.direction === "asc"
      ? <ArrowUp size={11} className="text-stone-700" />
      : <ArrowDown size={11} className="text-stone-700" />;
  };

  const SortableHeader = ({ sortKey, className = "", align = "left", children }) => {
    const alignClass = align === "right" ? "justify-end text-right" : align === "center" ? "justify-center text-center" : "justify-start text-left";

    return (
      <th className={className}>
        <button
          type="button"
          onClick={() => requestSort(sortKey)}
          className={`inline-flex w-full items-center gap-1 ${alignClass} rounded px-1 py-1 font-medium text-stone-500 hover:bg-stone-100 hover:text-stone-900`}
          title="Clique para ordenar"
        >
          <span>{children}</span>
          <SortIcon sortKey={sortKey} />
        </button>
      </th>
    );
  };

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

  const filtered = useMemo(() => {
    // Divide o texto digitado por espaços e remove itens vazios
    const searchTerms = normalizarBusca(query).split(/\s+/).filter(Boolean);
    const base = catalog.filter((c) => {
      const targetText = normalizarBusca(c.descricao);

      // Verifica se TODAS as palavras buscadas estão presentes na descrição do insumo
      return searchTerms.every((term) => targetText.includes(term));
    });

    const direction = sortConfig.direction === "asc" ? 1 : -1;
    const compareText = (a, b, field) => String(a[field] || "").localeCompare(String(b[field] || ""), "pt-BR", { sensitivity: "base" });
    const compareNumber = (a, b, field) => (num(a[field]) - num(b[field]));

    return [...base].sort((a, b) => {
      if (sortConfig.key === "valorUnitario") {
        const aHasPrice = hasPrice(a.valorUnitario);
        const bHasPrice = hasPrice(b.valorUnitario);
        if (aHasPrice !== bHasPrice) return aHasPrice ? -1 : 1;
        if (!aHasPrice && !bHasPrice) return compareText(a, b, "descricao");
        return compareNumber(a, b, "valorUnitario") * direction;
      }

      if (sortConfig.key === "ocorrencias") {
        const diff = compareNumber(a, b, "ocorrencias");
        if (diff !== 0) return diff * direction;
        return compareText(a, b, "descricao");
      }

      const diff = compareText(a, b, sortConfig.key);
      if (diff !== 0) return diff * direction;
      return compareText(a, b, "descricao");
    });
  }, [catalog, query, sortConfig]);

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4">
      <div className="mb-4">
        <h2 className="text-base font-semibold text-stone-800">Banco de Preços do Orçamento</h2>
        <p className="text-xs text-stone-500 mt-0.5">
          Valores exclusivos de {nomeProjeto}. Alterações não afetam os demais orçamentos.
        </p>
      </div>
      <div className="flex flex-wrap items-center justify-between gap-4 mb-4">
        <div className="relative w-72">
          <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-stone-400" />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Filtrar catálogo de preços..." className="w-full pl-8 pr-3 py-2 text-sm border border-stone-300 rounded-lg" />
        </div>
        <div className="flex flex-wrap justify-end gap-2">
          <button onClick={onApplyAllToCpus} className="flex items-center gap-1.5 px-3 py-1.5 text-xs border border-stone-300 rounded-lg font-medium bg-stone-50 hover:bg-stone-100 text-stone-700">
            <RefreshCw size={13} /> Sincronizar Tudo na Planilha de Custos
          </button>
          <button onClick={exportarXls} className="flex items-center gap-1.5 px-3 py-1.5 text-xs border border-stone-300 rounded-lg font-medium bg-white hover:bg-stone-50 text-stone-700">
            <Download size={13} /> Exportar .xlsx
          </button>
          <button onClick={() => setEditing({ id: null, descricao: "", tipo: "MAT", unidade: "un", valorUnitario: "" })} className="flex items-center gap-1.5 px-3 py-1.5 text-xs border border-stone-300 bg-white text-stone-700 rounded-lg font-medium hover:bg-stone-100">
            <Plus size={13} /> Novo Insumo Manual
          </button>
          <button
            type="button"
            onClick={onSaveProject}
            disabled={saving}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs bg-stone-900 text-white rounded-lg font-medium hover:bg-stone-700 disabled:opacity-50 disabled:cursor-not-allowed"
            title={`Salvar o Banco de Preços do orçamento ${nomeProjeto}`}
          >
            <Save size={13} /> {saving ? "Salvando..." : "Salvar orçamento"}
            {projectDirty && <span className="w-1.5 h-1.5 rounded-full bg-amber-400" title="Alterações pendentes" />}
          </button>
        </div>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-left text-xs border-collapse">
          <thead>
            <tr className="border-b border-stone-200 text-stone-400 font-normal">
              <SortableHeader sortKey="tipo" className="py-2 pr-3 w-28">Tipo</SortableHeader>
              <SortableHeader sortKey="descricao" className="py-2 pr-3">Descrição Única do Insumo</SortableHeader>
              <SortableHeader sortKey="unidade" className="py-2 pr-3 w-20">Un.</SortableHeader>
              <SortableHeader sortKey="valorUnitario" className="py-2 pr-3 w-32 text-right" align="right">Preço Padrão (R$)</SortableHeader>
              <SortableHeader sortKey="ocorrencias" className="py-2 pr-3 w-28 text-center" align="center">Na Planilha Ativa</SortableHeader>
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
function QuantidadeFormulaInput({ valor, onConfirmar }) {
  const textoDoValor = (valorAtual) => {
    if (valorAtual === "" || valorAtual === null || valorAtual === undefined) return "";
    const numero = Number(valorAtual);
    if (!Number.isFinite(numero)) return String(valorAtual);
    return numero.toLocaleString("pt-BR", {
      useGrouping: false,
      maximumFractionDigits: 15,
    });
  };
  const [texto, setTexto] = useState(() => textoDoValor(valor));
  const [emEdicao, setEmEdicao] = useState(false);
  const [invalida, setInvalida] = useState(false);
  const ignorarProximoBlurRef = useRef(false);

  useEffect(() => {
    if (!emEdicao && !invalida) setTexto(textoDoValor(valor));
  }, [valor, emEdicao, invalida]);

  const confirmarFormula = () => {
    const textoLimpo = texto.trim();
    if (!textoLimpo) {
      onConfirmar("");
      setInvalida(false);
      return;
    }

    const resultado = avaliarExpressaoNumerica(textoLimpo);
    if (resultado === null) {
      setInvalida(true);
      return;
    }

    onConfirmar(resultado);
    setTexto(textoDoValor(resultado));
    setInvalida(false);
  };

  return (
    <input
      type="text"
      inputMode="text"
      value={texto}
      onFocus={(evt) => {
        setEmEdicao(true);
        evt.currentTarget.select();
      }}
      onChange={(evt) => {
        setTexto(evt.target.value);
        setInvalida(false);
      }}
      onBlur={() => {
        setEmEdicao(false);
        if (ignorarProximoBlurRef.current) {
          ignorarProximoBlurRef.current = false;
          return;
        }
        confirmarFormula();
      }}
      onKeyDown={(evt) => {
        if (evt.key === "Enter") {
          evt.preventDefault();
          evt.currentTarget.blur();
        } else if (evt.key === "Escape") {
          ignorarProximoBlurRef.current = true;
          setTexto(textoDoValor(valor));
          setInvalida(false);
          evt.currentTarget.blur();
        }
      }}
      aria-invalid={invalida}
      title={
        invalida
          ? "Fórmula inválida"
          : "Aceita fórmulas como =2+3, 10-1,5, 4*2 e (3+2)*4"
      }
      className={`w-24 border rounded px-1.5 py-0.5 text-right font-mono bg-white outline-none ${
        invalida
          ? "border-red-400 ring-1 ring-red-200"
          : "border-stone-200 focus:border-stone-500"
      }`}
    />
  );
}

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
    () => criarIndiceBuscaCpus(cpus),
    [cpus]
  );

  const adicionarEtapa = () => {
    setEtapas([...etapas, { id: uid(), nome: `Nova Etapa ${etapas.length + 1}`, itens: [] }]);
  };

  const removerEtapa = (id) => {
    if (etapas.length <= 1) return;
    setEtapas(etapas.filter((e) => e.id !== id));
  };

  const adicionarGrupoAlternativas = (etapaId) => {
    const opcaoAId = uid();
    const opcaoBId = uid();
    setEtapas((atuais) =>
      atuais.map((etapa) => {
        if (etapa.id !== etapaId) return etapa;
        const grupos = gruposAlternativasDaEtapa(etapa);
        return {
          ...etapa,
          gruposAlternativas: [
            ...grupos,
            {
              id: uid(),
              nome: `Alternativa técnica ${grupos.length + 1}`,
              opcaoAtivaId: opcaoAId,
              opcoes: [
                { id: opcaoAId, nome: "Opção A" },
                { id: opcaoBId, nome: "Opção B" },
              ],
            },
          ],
        };
      })
    );
  };

  const atualizarGrupoAlternativas = (etapaId, grupoId, atualizador) => {
    setEtapas((atuais) =>
      atuais.map((etapa) =>
        etapa.id !== etapaId
          ? etapa
          : {
              ...etapa,
              gruposAlternativas: gruposAlternativasDaEtapa(etapa).map(
                (grupo) =>
                  grupo.id === grupoId ? atualizador(grupo) : grupo
              ),
            }
      )
    );
  };

  const adicionarOpcaoAlternativa = (etapaId, grupoId) => {
    atualizarGrupoAlternativas(etapaId, grupoId, (grupo) => ({
      ...grupo,
      opcoes: [
        ...(grupo.opcoes || []),
        {
          id: uid(),
          nome: `Opção ${String.fromCharCode(65 + (grupo.opcoes || []).length)}`,
        },
      ],
    }));
  };

  const removerOpcaoAlternativa = (etapaId, grupoId, opcaoId) => {
    setEtapas((atuais) =>
      atuais.map((etapa) => {
        if (etapa.id !== etapaId) return etapa;
        const grupoAtual = gruposAlternativasDaEtapa(etapa).find(
          (grupo) => grupo.id === grupoId
        );
        if (!grupoAtual || (grupoAtual.opcoes || []).length <= 2) return etapa;
        const novasOpcoes = grupoAtual.opcoes.filter(
          (opcao) => opcao.id !== opcaoId
        );
        return {
          ...etapa,
          gruposAlternativas: gruposAlternativasDaEtapa(etapa).map((grupo) =>
            grupo.id === grupoId
              ? {
                  ...grupo,
                  opcoes: novasOpcoes,
                  opcaoAtivaId:
                    grupo.opcaoAtivaId === opcaoId
                      ? novasOpcoes[0]?.id || ""
                      : grupo.opcaoAtivaId,
                }
              : grupo
          ),
          itens: (etapa.itens || []).map((item) =>
            item.alternativaGrupoId === grupoId &&
            item.alternativaOpcaoId === opcaoId
              ? {
                  ...item,
                  alternativaGrupoId: "",
                  alternativaOpcaoId: "",
                }
              : item
          ),
        };
      })
    );
  };

  const removerGrupoAlternativas = (etapaId, grupoId) => {
    setEtapas((atuais) =>
      atuais.map((etapa) =>
        etapa.id !== etapaId
          ? etapa
          : {
              ...etapa,
              gruposAlternativas: gruposAlternativasDaEtapa(etapa).filter(
                (grupo) => grupo.id !== grupoId
              ),
              itens: (etapa.itens || []).map((item) =>
                item.alternativaGrupoId === grupoId
                  ? {
                      ...item,
                      alternativaGrupoId: "",
                      alternativaOpcaoId: "",
                    }
                  : item
              ),
            }
      )
    );
  };

  const vincularItemAlternativa = (etapaId, itemId, valor) => {
    const [grupoId = "", opcaoId = ""] = String(valor || "").split("|");
    setEtapas((atuais) =>
      atuais.map((etapa) =>
        etapa.id !== etapaId
          ? etapa
          : {
              ...etapa,
              itens: (etapa.itens || []).map((item) =>
                item.id === itemId
                  ? {
                      ...item,
                      alternativaGrupoId: grupoId,
                      alternativaOpcaoId: opcaoId,
                    }
                  : item
              ),
            }
      )
    );
  };

  const moverEtapa = (index, direction) => {
    setEtapas((currentEtapas) => {
      const targetIndex = index + direction;
      if (targetIndex < 0 || targetIndex >= currentEtapas.length) return currentEtapas;
      const next = [...currentEtapas];
      [next[index], next[targetIndex]] = [next[targetIndex], next[index]];
      return next;
    });
  };

  const moverCpuNaEtapa = (etapaId, index, direction) => {
    setEtapas((currentEtapas) =>
      currentEtapas.map((etapa) => {
        if (etapa.id !== etapaId) return etapa;
        const targetIndex = index + direction;
        if (targetIndex < 0 || targetIndex >= etapa.itens.length) return etapa;
        const nextItens = [...etapa.itens];
        [nextItens[index], nextItens[targetIndex]] = [nextItens[targetIndex], nextItens[index]];
        return { ...etapa, itens: nextItens };
      })
    );
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

  const expandirTudo = () => {
    const itensAbertos = {};
    etapas.forEach((etapa) => {
      itensAtivosDaEtapa(etapa).forEach((item) => {
        itensAbertos[item.id] = true;
      });
    });
    setEtapasRecolhidas({});
    setItensExpandidos(itensAbertos);
  };

  const recolherTudo = () => {
    const etapasFechadas = {};
    etapas.forEach((etapa) => {
      etapasFechadas[etapa.id] = true;
    });
    setEtapasRecolhidas(etapasFechadas);
    setItensExpandidos({});
  };

  const obterCpusFiltradas = (textoBusca) => {
    if (!textoBusca || !textoBusca.trim()) return [];
    const searchTerms = normalizarBusca(textoBusca).split(/\s+/).filter(Boolean);
    const result = [];
    for (const item of cpuSearchIndex) {
      if (searchTerms.every((term) => item.haystack.includes(term))) {
        result.push(item.cpu);
        if (result.length >= 10) break;
      }
    }
    return result;
  };

  const obterInsumosFiltrados = (textoBusca) =>
    buscarInsumosCatalogo(catalogMap, textoBusca, 8);

  const handleKeyDown = (evt, etapaId, listaCpus) => {
    const currentIndex = activeIndices[etapaId] !== undefined ? activeIndices[etapaId] : -1;

    if (evt.key === "ArrowDown") {
      if (listaCpus.length === 0) return;
      evt.preventDefault();
      const nextIndex = (currentIndex + 1) % listaCpus.length;
      setActiveIndices({ ...activeIndices, [etapaId]: nextIndex });
    } else if (evt.key === "ArrowUp") {
      if (listaCpus.length === 0) return;
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
      <div className="flex flex-wrap justify-between items-center gap-2 bg-white border border-stone-200 rounded-lg p-3">
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={expandirTudo}
            className="px-2 py-1.5 text-[11px] font-medium border border-stone-200 rounded hover:bg-stone-50 text-stone-600 flex items-center gap-1"
          >
            <ChevronDown size={13} /> Expandir Tudo
          </button>
          <button
            type="button"
            onClick={recolherTudo}
            className="px-2 py-1.5 text-[11px] font-medium border border-stone-200 rounded hover:bg-stone-50 text-stone-600 flex items-center gap-1"
          >
            <ChevronRight size={13} /> Recolher Tudo
          </button>
        </div>
        <button onClick={adicionarEtapa} className="flex items-center gap-1 px-3 py-1.5 text-xs bg-stone-900 text-white rounded-lg hover:bg-stone-700">
          <Plus size={14} /> Adicionar Nova Etapa
        </button>
      </div>

      {/* Listagem das Etapas */}
      <div className="space-y-4">
        {etapas.map((e, etapaIndex) => {
          const termoBuscaEtapa = buscasPorEtapa[e.id] || "";
          const filtradasParaEstaEtapa = obterCpusFiltradas(termoBuscaEtapa);
          const insumosParaEstaEtapa = obterInsumosFiltrados(termoBuscaEtapa);
          const activeIndex = activeIndices[e.id] !== undefined ? activeIndices[e.id] : -1;
          const etapaRecolhida = !!etapasRecolhidas[e.id];
          const totalEtapa = itensAtivosDaEtapa(e).reduce(
            (s, it) => s + num(it.quantidade) * cpuValorUnit(it.insumos, cpus, catalogMap),
            0
          );
          const gruposAlternativas = gruposAlternativasDaEtapa(e);

          return (
            <div key={e.id} className="bg-white border border-stone-200 rounded-lg overflow-visible">
              <div
                onClick={() => toggleRecolherEtapa(e.id)}
                className="bg-stone-200 px-4 py-2.5 flex flex-wrap justify-between items-center gap-3 border-b border-stone-300 cursor-pointer select-none hover:bg-stone-300"
              >
                <div className="min-w-0 flex-1 flex items-center gap-2">
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
                  {etapas.length > 1 && (
                    <div
                      className="flex items-center rounded-md border border-stone-300 bg-white overflow-hidden"
                      onClick={(evt) => evt.stopPropagation()}
                    >
                      <button
                        type="button"
                        onClick={() => moverEtapa(etapaIndex, -1)}
                        disabled={etapaIndex === 0}
                        className="w-7 h-7 flex items-center justify-center text-stone-500 hover:bg-stone-100 hover:text-stone-900 disabled:opacity-25 disabled:cursor-not-allowed"
                        title="Mover etapa para cima"
                        aria-label="Mover etapa para cima"
                      >
                        <ArrowUp size={13} />
                      </button>
                      <button
                        type="button"
                        onClick={() => moverEtapa(etapaIndex, 1)}
                        disabled={etapaIndex === etapas.length - 1}
                        className="w-7 h-7 flex items-center justify-center border-l border-stone-200 text-stone-500 hover:bg-stone-100 hover:text-stone-900 disabled:opacity-25 disabled:cursor-not-allowed"
                        title="Mover etapa para baixo"
                        aria-label="Mover etapa para baixo"
                      >
                        <ArrowDown size={13} />
                      </button>
                    </div>
                  )}
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
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <p className="text-xs font-semibold text-stone-700">
                      Alternativas técnicas
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() => adicionarGrupoAlternativas(e.id)}
                    className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-[11px] font-medium border border-amber-300 bg-amber-50 text-amber-800 rounded-md hover:bg-amber-100"
                  >
                    <Plus size={13} /> Grupo de alternativas
                  </button>
                </div>

                {gruposAlternativas.length > 0 && (
                  <div className="space-y-2">
                    {gruposAlternativas.map((grupo) => (
                      <div
                        key={grupo.id}
                        className="border border-amber-200 bg-amber-50/40 rounded-md p-3"
                      >
                        <div className="flex items-center gap-2">
                          <input
                            value={grupo.nome || ""}
                            onChange={(evt) =>
                              atualizarGrupoAlternativas(
                                e.id,
                                grupo.id,
                                (atual) => ({
                                  ...atual,
                                  nome: evt.target.value,
                                })
                              )
                            }
                            className="min-w-0 flex-1 bg-transparent border-b border-amber-300 px-1 py-1 text-xs font-semibold text-amber-900 outline-none focus:border-amber-600"
                            aria-label="Nome do grupo de alternativas"
                          />
                          <button
                            type="button"
                            onClick={() =>
                              removerGrupoAlternativas(e.id, grupo.id)
                            }
                            className="w-7 h-7 inline-flex items-center justify-center text-amber-600 hover:bg-red-50 hover:text-red-600 rounded"
                            title="Remover grupo de alternativas"
                            aria-label="Remover grupo de alternativas"
                          >
                            <Trash2 size={13} />
                          </button>
                        </div>

                        <div className="mt-2 flex flex-wrap items-center gap-2">
                          {(grupo.opcoes || []).map((opcao) => {
                            const selecionada =
                              grupo.opcaoAtivaId === opcao.id;
                            return (
                              <label
                                key={opcao.id}
                                className={`flex items-center gap-1.5 border rounded-md px-2 py-1 ${
                                  selecionada
                                    ? "border-emerald-400 bg-emerald-50"
                                    : "border-stone-200 bg-white"
                                }`}
                              >
                                <input
                                  type="radio"
                                  name={`alternativa-${grupo.id}`}
                                  checked={selecionada}
                                  onChange={() =>
                                    atualizarGrupoAlternativas(
                                      e.id,
                                      grupo.id,
                                      (atual) => ({
                                        ...atual,
                                        opcaoAtivaId: opcao.id,
                                      })
                                    )
                                  }
                                  className="accent-emerald-600"
                                />
                                <input
                                  value={opcao.nome || ""}
                                  onChange={(evt) =>
                                    atualizarGrupoAlternativas(
                                      e.id,
                                      grupo.id,
                                      (atual) => ({
                                        ...atual,
                                        opcoes: (atual.opcoes || []).map(
                                          (item) =>
                                            item.id === opcao.id
                                              ? {
                                                  ...item,
                                                  nome: evt.target.value,
                                                }
                                              : item
                                        ),
                                      })
                                    )
                                  }
                                  className="w-32 bg-transparent text-[11px] font-medium outline-none"
                                  aria-label="Nome da alternativa"
                                />
                                {(grupo.opcoes || []).length > 2 && (
                                  <button
                                    type="button"
                                    onClick={() =>
                                      removerOpcaoAlternativa(
                                        e.id,
                                        grupo.id,
                                        opcao.id
                                      )
                                    }
                                    className="text-stone-300 hover:text-red-500"
                                    title="Remover alternativa"
                                  >
                                    <X size={12} />
                                  </button>
                                )}
                                {selecionada && (
                                  <span className="text-[9px] font-semibold uppercase text-emerald-700">
                                    No total
                                  </span>
                                )}
                              </label>
                            );
                          })}
                          <button
                            type="button"
                            onClick={() =>
                              adicionarOpcaoAlternativa(e.id, grupo.id)
                            }
                            className="inline-flex items-center gap-1 px-2 py-1 text-[10px] text-stone-500 hover:text-stone-900"
                          >
                            <Plus size={11} /> Opção
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}

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
                      {filtradasParaEstaEtapa.length === 0 && insumosParaEstaEtapa.length === 0 && (
                        <p className="p-3 text-stone-400">Nenhuma CPU ou insumo encontrado.</p>
                      )}
                      {filtradasParaEstaEtapa.length > 0 && (
                        <div className="px-2.5 py-1.5 bg-stone-100 border-b border-stone-200 text-[10px] font-semibold uppercase text-stone-500">
                          CPUs pelo nome ou código
                        </div>
                      )}
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
                      {insumosParaEstaEtapa.length > 0 && (
                        <>
                          <div className="px-2.5 py-1.5 bg-stone-100 border-y border-stone-200 text-[10px] font-semibold uppercase text-stone-500">
                            Insumos encontrados
                          </div>
                          {insumosParaEstaEtapa.map((insumo) => (
                            <div key={insumo.key || insumo.id} className="p-2 border-b border-stone-100 last:border-0 flex items-center gap-2 bg-stone-50/50">
                              <span className="text-[9px] px-1.5 py-0.5 rounded bg-white border border-stone-200 text-stone-500 shrink-0">
                                {insumo.tipo || "INS"}
                              </span>
                              <span className="flex-1 min-w-0 truncate text-stone-700" title={insumo.descricao}>{insumo.descricao}</span>
                              <span className="font-mono text-[10px] text-stone-400 shrink-0">{insumo.unidade || "un"}</span>
                            </div>
                          ))}
                        </>
                      )}
                    </div>
                  )}
                </div>

                {/* Exibição dos itens da Etapa */}
                {e.itens.length === 0 && !termoBuscaEtapa.trim() && (
                  <p className="text-xs text-stone-400 italic pt-1">Nenhuma CPU lançada nesta etapa.</p>
                )}
                {e.itens.map((it, itemIndex) => {
                  const estaExpandido = !!itensExpandidos[it.id]; // Por padrão, undefined avalia como falso (recolhido)
                  const grupoDoItem = grupoAlternativaDoItem(e, it);
                  const opcaoDoItem = (grupoDoItem?.opcoes || []).find(
                    (opcao) => opcao.id === it.alternativaOpcaoId
                  );
                  const itemAtivo = itemIncluidoNoCalculo(e, it);

                  return (
                    <div
                      key={it.id}
                      className={`border rounded-lg p-3 ${
                        grupoDoItem
                          ? itemAtivo
                            ? "border-emerald-200 bg-emerald-50/20"
                            : "border-stone-200 bg-stone-100/60 opacity-65"
                          : "border-stone-100 bg-stone-50/30"
                      }`}
                    >
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
                          <span
                            className="shrink-0 rounded bg-stone-200 px-1.5 py-0.5 font-mono text-[10px] font-semibold text-stone-600"
                            title="Numeração automática da CPU"
                          >
                            {etapaIndex + 1}.{itemIndex + 1}
                          </span>
                          <div className="min-w-0">
                            <span className="font-mono text-[10px] text-stone-400">{it.codigo}</span>
                            <h4 className="text-xs font-semibold text-stone-800 truncate">{it.servico}</h4>
                          </div>
                        </div>
                        
                        <div className="flex flex-wrap items-center justify-end gap-3 text-xs">
                          {gruposAlternativas.length > 0 && (
                            <select
                              value={
                                grupoDoItem && opcaoDoItem
                                  ? `${grupoDoItem.id}|${opcaoDoItem.id}`
                                  : ""
                              }
                              onChange={(evt) =>
                                vincularItemAlternativa(
                                  e.id,
                                  it.id,
                                  evt.target.value
                                )
                              }
                              className={`max-w-52 h-7 px-2 border rounded-md bg-white text-[10px] outline-none ${
                                grupoDoItem
                                  ? itemAtivo
                                    ? "border-emerald-300 text-emerald-800"
                                    : "border-stone-300 text-stone-500"
                                  : "border-stone-200 text-stone-500"
                              }`}
                              title="Classificar CPU como serviço comum ou alternativa"
                              aria-label={`Classificação de ${it.servico}`}
                            >
                              <option value="">Serviço comum</option>
                              {gruposAlternativas.map((grupo) => (
                                <optgroup key={grupo.id} label={grupo.nome}>
                                  {(grupo.opcoes || []).map((opcao) => (
                                    <option
                                      key={opcao.id}
                                      value={`${grupo.id}|${opcao.id}`}
                                    >
                                      {opcao.nome}
                                    </option>
                                  ))}
                                </optgroup>
                              ))}
                            </select>
                          )}
                          {grupoDoItem && (
                            <span
                              className={`px-1.5 py-0.5 rounded text-[9px] font-semibold uppercase ${
                                itemAtivo
                                  ? "bg-emerald-100 text-emerald-700"
                                  : "bg-stone-200 text-stone-500"
                              }`}
                            >
                              {itemAtivo ? "Incluída" : "Fora do total"}
                            </span>
                          )}
                          {e.itens.length > 1 && (
                            <div className="flex items-center rounded-md border border-stone-200 bg-white overflow-hidden shrink-0">
                              <button
                                type="button"
                                onClick={() => moverCpuNaEtapa(e.id, itemIndex, -1)}
                                disabled={itemIndex === 0}
                                className="w-7 h-7 flex items-center justify-center text-stone-400 hover:bg-stone-100 hover:text-stone-900 disabled:opacity-25 disabled:cursor-not-allowed"
                                title="Mover CPU para cima"
                                aria-label="Mover CPU para cima"
                              >
                                <ArrowUp size={12} />
                              </button>
                              <button
                                type="button"
                                onClick={() => moverCpuNaEtapa(e.id, itemIndex, 1)}
                                disabled={itemIndex === e.itens.length - 1}
                                className="w-7 h-7 flex items-center justify-center border-l border-stone-100 text-stone-400 hover:bg-stone-100 hover:text-stone-900 disabled:opacity-25 disabled:cursor-not-allowed"
                                title="Mover CPU para baixo"
                                aria-label="Mover CPU para baixo"
                              >
                                <ArrowDown size={12} />
                              </button>
                            </div>
                          )}
                          <div className="flex items-center gap-1.5">
                            <span className="text-stone-400">Qtd:</span>
                            <QuantidadeFormulaInput
                              valor={it.quantidade}
                              onConfirmar={(quantidade) => mudarQuantidadeItem(e.id, it.id, quantidade)}
                            />
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
  const collemAtivo = !!bdi.collemAtivo;
  const collemX = bdi.collemX === "" ? "" : (bdi.collemX ?? 1);
  const collemY = bdi.collemY === "" ? "" : (bdi.collemY ?? 1);

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

  const handleCollemChange = (campo, valor) => {
    setBdi(prev => ({ ...prev, [campo]: valor }));
  };

  const normalizarDivisorCollem = (campo, valor) => {
    if (num(valor) > 0) return;
    setBdi(prev => ({ ...prev, [campo]: 1 }));
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
  const calcularValoresIndices = (taxas, baseCalculo) => {
    const valores = {};
    let acumulado = num(baseCalculo);

    ["admCentral", "contabilidade", "contingenciamento", "custoFinanceiro", "lucro"].forEach((campo) => {
      valores[campo] = acumulado * num(taxas?.[campo]);
      acumulado += valores[campo];
    });

    const das = num(taxas?.dasAnexoIV);
    const art = num(taxas?.art);
    const denominador = 1 - das - art;
    const venda = denominador > 0 ? acumulado / denominador : acumulado;
    valores.dasAnexoIV = denominador > 0 ? venda * das : 0;
    valores.art = denominador > 0 ? venda * art : 0;
    return valores;
  };
  const valoresIndicesGerais = calcularValoresIndices(bdi, bdiCalc.custoBaseBdiGeral);
  const valoresIndicesMateriais = calcularValoresIndices(bdiMats, bdiCalc.custoMateriaisFaturamentoDireto);

  return (
    <div className="space-y-6">
      {/* Barra de controle superior */}
      <div className="bg-white border border-stone-200 rounded-lg p-4 flex justify-between items-center flex-wrap gap-3">
        <div>
          <h3 className="font-semibold text-sm text-stone-800">Opções do Regime de Faturamento</h3>
          <p className="text-xs text-stone-400">Ative o BDI diferenciado e escolha os materiais correspondentes na aba Materiais.</p>
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

      <div className={`border rounded-lg p-4 transition-colors ${collemAtivo ? "border-amber-300 bg-amber-50/60" : "border-stone-200 bg-white"}`}>
        <div className="flex justify-between items-start flex-wrap gap-3">
          <div>
            <h3 className="font-semibold text-sm text-stone-800">Condição comercial COLLEM</h3>
            <p className="text-xs text-stone-500">Aplica os divisores X e Y sobre o preço de venda calculado com BDI.</p>
          </div>
          <label className="flex items-center gap-2 bg-white border border-stone-200 px-3 py-1.5 rounded-md cursor-pointer select-none hover:bg-stone-50 text-xs font-semibold text-stone-700">
            <input
              type="checkbox"
              checked={collemAtivo}
              onChange={(e) => setBdi(prev => ({
                ...prev,
                collemAtivo: e.target.checked,
                collemX: num(prev.collemX) > 0 ? prev.collemX : 1,
                collemY: num(prev.collemY) > 0 ? prev.collemY : 1,
              }))}
              className="w-4 h-4 accent-amber-600 rounded"
            />
            Ativar COLLEM
          </label>
        </div>

        {collemAtivo && (
          <div className="mt-4 grid grid-cols-1 sm:grid-cols-[minmax(120px,180px)_minmax(120px,180px)_1fr] gap-3 items-end">
            <label className="text-xs text-stone-600">
              <span className="block mb-1 font-medium">Divisor X</span>
              <input
                type="number"
                min="0.000001"
                step="any"
                value={collemX}
                onChange={(e) => handleCollemChange("collemX", e.target.value)}
                onBlur={(e) => normalizarDivisorCollem("collemX", e.target.value)}
                className="w-full h-9 border border-amber-300 rounded-md px-2.5 bg-white text-sm font-mono outline-none focus:ring-1 focus:ring-amber-500"
              />
            </label>
            <label className="text-xs text-stone-600">
              <span className="block mb-1 font-medium">Divisor Y</span>
              <input
                type="number"
                min="0.000001"
                step="any"
                value={collemY}
                onChange={(e) => handleCollemChange("collemY", e.target.value)}
                onBlur={(e) => normalizarDivisorCollem("collemY", e.target.value)}
                className="w-full h-9 border border-amber-300 rounded-md px-2.5 bg-white text-sm font-mono outline-none focus:ring-1 focus:ring-amber-500"
              />
            </label>
            <div className="min-h-9 rounded-md border border-amber-200 bg-white px-3 py-2 text-xs text-stone-600">
              <span className="font-mono">R$ {fmt(bdiCalc.valorVendaBase)} ÷ {fmt(bdiCalc.collemX)} ÷ {fmt(bdiCalc.collemY)}</span>
              <span className="font-semibold text-stone-900 ml-2">= R$ {fmt(bdiCalc.valorVendaBruto)}</span>
            </div>
          </div>
        )}
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
              <div className="grid grid-cols-[minmax(0,1fr)_5rem_7.5rem] gap-2 text-[9px] uppercase text-stone-400 font-medium">
                <span>Índice</span><span className="text-right">%</span><span className="text-right">Valor</span>
              </div>
              
              <div className="space-y-3">
                <h5 className="font-medium text-stone-400 uppercase text-[9px]">Administração e Riscos</h5>
                <BdiInput label="Administração Central" value={bdi.admCentral} amount={valoresIndicesGerais.admCentral} onChange={(v) => handleGeralChange("admCentral", v)} />
                <BdiInput label="Contabilidade / Seguros" value={bdi.contabilidade} amount={valoresIndicesGerais.contabilidade} onChange={(v) => handleGeralChange("contabilidade", v)} />
                <BdiInput label="Contingenciamento" value={bdi.contingenciamento} amount={valoresIndicesGerais.contingenciamento} onChange={(v) => handleGeralChange("contingenciamento", v)} />
                <BdiInput label="Custo Financeiro" value={bdi.custoFinanceiro} amount={valoresIndicesGerais.custoFinanceiro} onChange={(v) => handleGeralChange("custoFinanceiro", v)} />
                
                <h5 className="font-medium text-stone-400 uppercase text-[9px] pt-1">Margem e Impostos</h5>
                <BdiInput label="Lucro Real de Venda" value={bdi.lucro} amount={valoresIndicesGerais.lucro} onChange={(v) => handleGeralChange("lucro", v)} />
                <BdiInput label="DAS / Tributos (Anexo IV)" value={bdi.dasAnexoIV} amount={valoresIndicesGerais.dasAnexoIV} onChange={(v) => handleGeralChange("dasAnexoIV", v)} />
                <BdiInput label="ART / Encargos Contrato" value={bdi.art} amount={valoresIndicesGerais.art} onChange={(v) => handleGeralChange("art", v)} />
                <BdiInput
                  label="Retenção de INSS (somente MO)"
                  value={bdi.retencaoInss || 0}
                  amount={bdiCalc.retencaoInssValor}
                  onChange={(v) => handleGeralChange("retencaoInss", v)}
                  maxPercent={99}
                />
                <p className="text-[10px] leading-4 text-stone-400">
                  Recompõe o preço da mão de obra para compensar o desconto da retenção.
                </p>
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
                <div className="grid grid-cols-[minmax(0,1fr)_5rem_7.5rem] gap-2 text-[9px] uppercase text-emerald-600/70 font-medium">
                  <span>Índice</span><span className="text-right">%</span><span className="text-right">Valor</span>
                </div>

                <div className="space-y-3">
                  <h5 className="font-medium text-emerald-600/70 uppercase text-[9px]">Administração e Riscos</h5>
                  <BdiInput label="Administração Central" value={bdiMats.admCentral} amount={valoresIndicesMateriais.admCentral} onChange={(v) => handleMatChange("admCentral", v)} />
                  <BdiInput label="Contabilidade / Seguros" value={bdiMats.contabilidade} amount={valoresIndicesMateriais.contabilidade} onChange={(v) => handleMatChange("contabilidade", v)} />
                  <BdiInput label="Contingenciamento" value={bdiMats.contingenciamento} amount={valoresIndicesMateriais.contingenciamento} onChange={(v) => handleMatChange("contingenciamento", v)} />
                  <BdiInput label="Custo Financeiro" value={bdiMats.custoFinanceiro} amount={valoresIndicesMateriais.custoFinanceiro} onChange={(v) => handleMatChange("custoFinanceiro", v)} />
                  
                  <h5 className="font-medium text-emerald-600/70 uppercase text-[9px] pt-1">Margem e Impostos</h5>
                  <BdiInput label="Lucro Real de Venda" value={bdiMats.lucro} amount={valoresIndicesMateriais.lucro} onChange={(v) => handleMatChange("lucro", v)} />
                  <BdiInput label="DAS / Tributos (Anexo IV)" value={bdiMats.dasAnexoIV} amount={valoresIndicesMateriais.dasAnexoIV} onChange={(v) => handleMatChange("dasAnexoIV", v)} />
                  <BdiInput label="ART / Encargos Contrato" value={bdiMats.art} amount={valoresIndicesMateriais.art} onChange={(v) => handleMatChange("art", v)} />
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
              {bdiCalc.retencaoInss > 0 && (
                <>
                  <div className="flex justify-between border-t border-stone-800 pt-2">
                    <span className="text-stone-400">Base de mão de obra:</span>
                    <span className="font-mono">R$ {fmt(bdiCalc.custoMaoObra)}</span>
                  </div>
                  <div className="flex justify-between text-amber-300">
                    <span>Compensação INSS ({fmt(bdiCalc.retencaoInss * 100)}%):</span>
                    <span className="font-mono">R$ {fmt(bdiCalc.retencaoInssValor)}</span>
                  </div>
                </>
              )}
              {collemAtivo && (
                <>
                  <div className="flex justify-between border-t border-stone-800 pt-2">
                    <span className="text-stone-400">Venda antes do COLLEM:</span>
                    <span className="font-mono">R$ {fmt(bdiCalc.valorVendaBase)}</span>
                  </div>
                  <div className="flex justify-between text-amber-300">
                    <span>Divisores COLLEM:</span>
                    <span className="font-mono">÷ {fmt(bdiCalc.collemX)} ÷ {fmt(bdiCalc.collemY)}</span>
                  </div>
                </>
              )}
              {bdiCalc.descontoNegociacao > 0 && (
                <>
                  <div className="flex justify-between border-t border-stone-800 pt-2">
                    <span className="text-stone-400">Venda bruta:</span>
                    <span className="font-mono">R$ {fmt(bdiCalc.valorVendaBruto)}</span>
                  </div>
                  <div className="flex justify-between text-red-300">
                    <span>Desconto da negociação:</span>
                    <span className="font-mono">- R$ {fmt(bdiCalc.descontoNegociacao)}</span>
                  </div>
                </>
              )}
              <div className="flex justify-between border-t border-stone-800 pt-2">
                <span className="text-stone-400">{collemAtivo ? "Resultado sobre o custo:" : "Total BDI (Rateio):"}</span>
                <span className="font-mono">R$ {fmt(bdiCalc.totalDiValor)}</span>
              </div>
            </div>
          </div>
          <div className="mt-6 pt-4 border-t border-stone-800 text-right">
            <span className="text-[10px] text-stone-400 block uppercase font-medium">{collemAtivo ? "Preço Final de Venda com COLLEM" : "Preço Final de Venda"}</span>
            <span className="text-2xl font-bold font-mono text-white">R$ {fmt(bdiCalc.valorVenda)}</span>
          </div>
        </div>
      </div>
    </div>
  );
}

const limitarInteiroCronograma = (valor, minimo, maximo, padrao = minimo) => {
  const numero = Math.round(Number(valor));
  if (!Number.isFinite(numero)) return padrao;
  return Math.min(maximo, Math.max(minimo, numero));
};

const distribuirPercentuaisCronograma = (inicio, duracao) => {
  const percentuais = {};
  const percentualBase = Number((100 / duracao).toFixed(6));

  for (let indice = 0; indice < duracao; indice += 1) {
    const semana = inicio + indice;
    percentuais[String(semana)] =
      indice === duracao - 1
        ? Number((100 - percentualBase * (duracao - 1)).toFixed(6))
        : percentualBase;
  }

  return percentuais;
};

const dataLocalCronograma = (valor) => {
  const partes = String(valor || "").split("-").map(Number);
  if (partes.length !== 3 || partes.some((parte) => !Number.isFinite(parte))) return null;
  return new Date(partes[0], partes[1] - 1, partes[2]);
};

const adicionarDiasCronograma = (data, dias) => {
  const resultado = new Date(data);
  resultado.setDate(resultado.getDate() + dias);
  return resultado;
};

const formatarDataCronograma = (data) =>
  data.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" });

function CronogramaSemanal({
  projeto,
  etapas,
  cronograma,
  setCronograma,
  bdiCalc,
  cpus,
  catalogMap,
  cliente,
}) {
  const semanas = limitarInteiroCronograma(cronograma.semanas, 1, 104, 12);
  const dataInicio = dataLocalCronograma(cronograma.dataInicio);
  const tabelaScrollRef = useRef(null);
  const barraFixaScrollRef = useRef(null);
  const [barraFixa, setBarraFixa] = useState({
    visivel: false,
    esquerda: 0,
    largura: 0,
    larguraConteudo: 0,
  });

  const linhas = useMemo(
    () =>
      (etapas || []).map((etapa, indice) => {
        const id = etapa.id || `etapa-${indice}`;
        const totalVenda = itensAtivosDaEtapa(etapa).reduce(
          (soma, item) =>
            soma + itemVendaResumo(item, bdiCalc, cpus, catalogMap, cliente).total,
          0
        );
        const configuracao = cronograma.etapas?.[id] || {};
        const inicioPadrao = Math.min(indice + 1, semanas);
        const inicio = limitarInteiroCronograma(
          configuracao.inicio,
          1,
          semanas,
          inicioPadrao
        );
        const duracao = limitarInteiroCronograma(
          configuracao.duracao,
          1,
          semanas - inicio + 1,
          1
        );
        const distribuicaoPadrao = distribuirPercentuaisCronograma(inicio, duracao);
        const possuiDistribuicao =
          configuracao.percentuais &&
          Object.keys(configuracao.percentuais).length > 0;
        const percentuais = possuiDistribuicao
          ? configuracao.percentuais
          : distribuicaoPadrao;

        return {
          id,
          nome: etapa.nome || `Etapa ${indice + 1}`,
          numero: indice + 1,
          totalVenda,
          inicio,
          duracao,
          percentuais,
        };
      }),
    [etapas, cronograma.etapas, semanas, bdiCalc, cpus, catalogMap, cliente]
  );

  const totalProjeto = useMemo(
    () => linhas.reduce((soma, linha) => soma + linha.totalVenda, 0),
    [linhas]
  );

  const linhasCalculadas = useMemo(
    () =>
      linhas.map((linha) => {
        const percentuais = Array.from({ length: semanas }, (_, indice) => {
          const semana = indice + 1;
          const ativa =
            semana >= linha.inicio &&
            semana < linha.inicio + linha.duracao;
          return ativa ? Math.max(0, num(linha.percentuais[String(semana)])) : 0;
        });
        const somaPercentuais = percentuais.reduce(
          (soma, percentual) => soma + percentual,
          0
        );
        const pesoFisico =
          totalProjeto > 0 ? (linha.totalVenda / totalProjeto) * 100 : 0;

        return {
          ...linha,
          percentuais,
          somaPercentuais,
          pesoFisico,
          valoresSemanais: percentuais.map(
            (percentual) => linha.totalVenda * (percentual / 100)
          ),
          fisicoSemanal: percentuais.map(
            (percentual) => pesoFisico * (percentual / 100)
          ),
        };
      }),
    [linhas, semanas, totalProjeto]
  );

  const totaisSemanais = useMemo(
    () =>
      Array.from({ length: semanas }, (_, indice) =>
        linhasCalculadas.reduce(
          (soma, linha) => soma + linha.valoresSemanais[indice],
          0
        )
      ),
    [linhasCalculadas, semanas]
  );

  const fisicoSemanal = useMemo(
    () =>
      Array.from({ length: semanas }, (_, indice) =>
        linhasCalculadas.reduce(
          (soma, linha) => soma + linha.fisicoSemanal[indice],
          0
        )
      ),
    [linhasCalculadas, semanas]
  );

  const acumuladoFinanceiro = useMemo(() => {
    let acumulado = 0;
    return totaisSemanais.map((valor) => {
      acumulado += valor;
      return acumulado;
    });
  }, [totaisSemanais]);

  const acumuladoFisico = useMemo(() => {
    let acumulado = 0;
    return fisicoSemanal.map((valor) => {
      acumulado += valor;
      return acumulado;
    });
  }, [fisicoSemanal]);

  const totalPlanejado = totaisSemanais.reduce((soma, valor) => soma + valor, 0);
  const fisicoPlanejado = fisicoSemanal.reduce((soma, valor) => soma + valor, 0);
  const saldoPlanejado =
    Math.abs(totalProjeto - totalPlanejado) < 0.005
      ? 0
      : totalProjeto - totalPlanejado;

  useEffect(() => {
    const tabelaScroll = tabelaScrollRef.current;
    if (!tabelaScroll) return undefined;

    const atualizarBarraFixa = () => {
      const retangulo = tabelaScroll.getBoundingClientRect();
      const margem = 8;
      const esquerda = Math.max(margem, retangulo.left);
      const direita = Math.min(window.innerWidth - margem, retangulo.right);
      const largura = Math.max(0, direita - esquerda);
      const visivel =
        largura > 100 &&
        tabelaScroll.scrollWidth > tabelaScroll.clientWidth + 2;

      setBarraFixa({
        visivel,
        esquerda,
        largura,
        larguraConteudo: tabelaScroll.scrollWidth,
      });

      requestAnimationFrame(() => {
        if (barraFixaScrollRef.current) {
          barraFixaScrollRef.current.scrollLeft = tabelaScroll.scrollLeft;
        }
      });
    };

    atualizarBarraFixa();
    const observador = new ResizeObserver(atualizarBarraFixa);
    observador.observe(tabelaScroll);
    window.addEventListener("resize", atualizarBarraFixa);

    return () => {
      observador.disconnect();
      window.removeEventListener("resize", atualizarBarraFixa);
    };
  }, [semanas, linhasCalculadas.length]);

  const sincronizarComBarraFixa = (evento) => {
    const barraFixaScroll = barraFixaScrollRef.current;
    if (
      barraFixaScroll &&
      Math.abs(barraFixaScroll.scrollLeft - evento.currentTarget.scrollLeft) > 1
    ) {
      barraFixaScroll.scrollLeft = evento.currentTarget.scrollLeft;
    }
  };

  const sincronizarComTabela = (evento) => {
    const tabelaScroll = tabelaScrollRef.current;
    if (
      tabelaScroll &&
      Math.abs(tabelaScroll.scrollLeft - evento.currentTarget.scrollLeft) > 1
    ) {
      tabelaScroll.scrollLeft = evento.currentTarget.scrollLeft;
    }
  };

  const rotuloSemana = (indice) => {
    if (!dataInicio) return "";
    const inicioSemana = adicionarDiasCronograma(dataInicio, indice * 7);
    const fimSemana = adicionarDiasCronograma(inicioSemana, 6);
    return `${formatarDataCronograma(inicioSemana)} a ${formatarDataCronograma(fimSemana)}`;
  };

  const salvarConfiguracaoEtapa = (id, configuracao) => {
    setCronograma((atual) => ({
      ...atual,
      etapas: {
        ...(atual.etapas || {}),
        [id]: configuracao,
      },
    }));
  };

  const alterarPeriodo = (linha, campo, valor) => {
    let inicio = linha.inicio;
    let duracao = linha.duracao;

    if (campo === "inicio") {
      inicio = limitarInteiroCronograma(valor, 1, semanas, linha.inicio);
      duracao = Math.min(duracao, semanas - inicio + 1);
    } else {
      duracao = limitarInteiroCronograma(
        valor,
        1,
        semanas - inicio + 1,
        linha.duracao
      );
    }

    salvarConfiguracaoEtapa(linha.id, {
      inicio,
      duracao,
      percentuais: distribuirPercentuaisCronograma(inicio, duracao),
    });
  };

  const alterarPercentual = (linha, semana, valor) => {
    const percentual =
      valor === "" ? 0 : Math.min(100, Math.max(0, num(valor)));
    const percentuais = {};

    for (let numeroSemana = linha.inicio; numeroSemana < linha.inicio + linha.duracao; numeroSemana += 1) {
      percentuais[String(numeroSemana)] =
        numeroSemana === semana
          ? percentual
          : num(linha.percentuais[String(numeroSemana)]);
    }

    salvarConfiguracaoEtapa(linha.id, {
      inicio: linha.inicio,
      duracao: linha.duracao,
      percentuais,
    });
  };

  const distribuirTodasAsEtapas = () => {
    const configuracoes = {};
    linhas.forEach((linha) => {
      configuracoes[linha.id] = {
        inicio: linha.inicio,
        duracao: linha.duracao,
        percentuais: distribuirPercentuaisCronograma(
          linha.inicio,
          linha.duracao
        ),
      };
    });
    setCronograma((atual) => ({ ...atual, etapas: configuracoes }));
  };

  const exportarCronograma = () => {
    const cabecalho = [
      "Etapa",
      "Valor de Venda (R$)",
      "Peso Físico (%)",
      "Semana Inicial",
      "Duração (semanas)",
    ];
    for (let semana = 1; semana <= semanas; semana += 1) {
      const periodo = rotuloSemana(semana - 1);
      cabecalho.push(
        `Semana ${semana}${periodo ? ` - ${periodo}` : ""} (%)`,
        `Semana ${semana}${periodo ? ` - ${periodo}` : ""} (R$)`
      );
    }

    const dados = linhasCalculadas.map((linha) => {
      const registro = [
        `${linha.numero}. ${linha.nome}`,
        linha.totalVenda,
        linha.pesoFisico,
        linha.inicio,
        linha.duracao,
      ];
      linha.percentuais.forEach((percentual, indice) => {
        registro.push(percentual, linha.valoresSemanais[indice]);
      });
      return registro;
    });

    const linhaTotal = ["TOTAL SEMANAL", totalProjeto, 100, "", ""];
    const linhaAcumulada = ["ACUMULADO", "", "", "", ""];
    for (let indice = 0; indice < semanas; indice += 1) {
      linhaTotal.push(fisicoSemanal[indice], totaisSemanais[indice]);
      linhaAcumulada.push(
        acumuladoFisico[indice],
        acumuladoFinanceiro[indice]
      );
    }

    const linhasPlanilha = [
      [`CRONOGRAMA FÍSICO-FINANCEIRO SEMANAL - ${projeto.nome || "ORÇAMENTO"}`],
      [
        cronograma.dataInicio
          ? `Data de início: ${dataInicio?.toLocaleDateString("pt-BR") || cronograma.dataInicio}`
          : "Data de início não definida",
      ],
      [],
      cabecalho,
      ...dados,
      linhaTotal,
      linhaAcumulada,
    ];
    const planilha = XLSX.utils.aoa_to_sheet(linhasPlanilha);
    const ultimaColuna = cabecalho.length - 1;
    planilha["!merges"] = [
      XLSX.utils.decode_range(`A1:${XLSX.utils.encode_col(ultimaColuna)}1`),
      XLSX.utils.decode_range(`A2:${XLSX.utils.encode_col(ultimaColuna)}2`),
    ];
    planilha["!cols"] = [
      { wch: 42 },
      { wch: 18 },
      { wch: 15 },
      { wch: 14 },
      { wch: 18 },
      ...Array.from({ length: semanas * 2 }, (_, indice) => ({
        wch: indice % 2 === 0 ? 16 : 18,
      })),
    ];
    planilha["!freeze"] = { xSplit: 1, ySplit: 4 };

    const faixa = XLSX.utils.decode_range(planilha["!ref"]);
    for (let coluna = faixa.s.c; coluna <= faixa.e.c; coluna += 1) {
      const celulaCabecalho = planilha[XLSX.utils.encode_cell({ r: 3, c: coluna })];
      if (celulaCabecalho) {
        celulaCabecalho.s = {
          font: { bold: true, color: { rgb: "FFFFFF" } },
          fill: { fgColor: { rgb: "55753A" } },
          alignment: { horizontal: "center", vertical: "center", wrapText: true },
          border: {
            top: { style: "thin", color: { rgb: "D6D3D1" } },
            bottom: { style: "thin", color: { rgb: "D6D3D1" } },
            left: { style: "thin", color: { rgb: "D6D3D1" } },
            right: { style: "thin", color: { rgb: "D6D3D1" } },
          },
        };
      }
    }

    const primeiraLinhaDados = 4;
    const ultimaLinhaDados = primeiraLinhaDados + linhasCalculadas.length - 1;
    for (let linha = primeiraLinhaDados; linha <= ultimaLinhaDados + 2; linha += 1) {
      for (let coluna = 0; coluna <= ultimaColuna; coluna += 1) {
        const endereco = XLSX.utils.encode_cell({ r: linha, c: coluna });
        const celula = planilha[endereco];
        if (!celula) continue;
        celula.s = {
          ...(celula.s || {}),
          font: {
            name: "Aptos",
            sz: 10,
            bold: linha > ultimaLinhaDados,
          },
          fill:
            linha > ultimaLinhaDados
              ? { fgColor: { rgb: linha === ultimaLinhaDados + 1 ? "E7E5E4" : "F5F5F4" } }
              : undefined,
          alignment: {
            vertical: "center",
            wrapText: coluna === 0,
          },
          border: {
            bottom: { style: "thin", color: { rgb: "E7E5E4" } },
          },
        };
        if (
          coluna === 1 ||
          (coluna >= 5 && (coluna - 5) % 2 === 1)
        ) {
          celula.z = XLSX_MOEDA;
        } else if (
          coluna === 2 ||
          (coluna >= 5 && (coluna - 5) % 2 === 0)
        ) {
          celula.z = "0.00%";
          celula.v = num(celula.v) / 100;
        }
      }
    }

    if (planilha.A1) {
      planilha.A1.s = {
        font: { name: "Aptos", sz: 15, bold: true, color: { rgb: "385723" } },
        alignment: { horizontal: "left", vertical: "center" },
      };
    }
    if (planilha.A2) {
      planilha.A2.s = {
        font: { name: "Aptos", sz: 10, color: { rgb: "57534E" } },
      };
    }
    planilha["!rows"] = [{ hpt: 24 }, { hpt: 18 }, { hpt: 8 }, { hpt: 38 }];

    const arquivo = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(arquivo, planilha, "Cronograma Semanal");
    XLSX.writeFile(
      arquivo,
      `${nomeArquivoSeguro(projeto.nome)}_Cronograma_Semanal.xlsx`
    );
  };

  return (
    <div className="bg-white border border-stone-200 shadow-sm rounded-lg overflow-hidden">
      <div className="p-5 border-b border-stone-200 flex flex-wrap items-end justify-between gap-4">
        <div>
          <h2 className="text-base font-semibold text-stone-800">
            Cronograma Físico-Financeiro Semanal
          </h2>
          <p className="text-xs text-stone-500">
            {linhasCalculadas.length} etapa(s) planejada(s)
          </p>
        </div>

        <div className="flex flex-wrap items-end gap-2">
          <label className="text-xs text-stone-600">
            <span className="block mb-1">Início da obra</span>
            <input
              type="date"
              value={cronograma.dataInicio}
              onChange={(evento) =>
                setCronograma((atual) => ({
                  ...atual,
                  dataInicio: evento.target.value,
                }))
              }
              className="h-9 border border-stone-300 rounded px-2 text-sm bg-white"
            />
          </label>
          <label className="text-xs text-stone-600">
            <span className="block mb-1">Total de semanas</span>
            <input
              type="number"
              min="1"
              max="104"
              value={semanas}
              onChange={(evento) =>
                setCronograma((atual) => ({
                  ...atual,
                  semanas: limitarInteiroCronograma(
                    evento.target.value,
                    1,
                    104,
                    semanas
                  ),
                }))
              }
              className="h-9 w-24 border border-stone-300 rounded px-2 text-right font-mono text-sm bg-white"
            />
          </label>
          <button
            type="button"
            onClick={distribuirTodasAsEtapas}
            className="h-9 px-3 inline-flex items-center gap-1.5 border border-stone-300 rounded text-xs font-medium text-stone-700 hover:bg-stone-50"
          >
            <RefreshCw size={14} /> Distribuir tudo
          </button>
          <button
            type="button"
            onClick={exportarCronograma}
            className="h-9 px-3 inline-flex items-center gap-1.5 border border-emerald-700 bg-emerald-700 rounded text-xs font-medium text-white hover:bg-emerald-800"
          >
            <Download size={14} /> Excel
          </button>
        </div>
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-4 border-b border-stone-200 bg-stone-50">
        <div className="px-5 py-3 border-r border-b lg:border-b-0 border-stone-200">
          <span className="block text-[10px] uppercase text-stone-500">Venda do orçamento</span>
          <strong className="font-mono text-sm text-stone-800">R$ {fmt(totalProjeto)}</strong>
        </div>
        <div className="px-5 py-3 border-b lg:border-b-0 lg:border-r border-stone-200">
          <span className="block text-[10px] uppercase text-stone-500">Financeiro planejado</span>
          <strong className="font-mono text-sm text-stone-800">R$ {fmt(totalPlanejado)}</strong>
        </div>
        <div className="px-5 py-3 border-r border-stone-200">
          <span className="block text-[10px] uppercase text-stone-500">Físico planejado</span>
          <strong className={`font-mono text-sm ${Math.abs(fisicoPlanejado - 100) < 0.01 ? "text-emerald-700" : "text-amber-700"}`}>
            {fmt(fisicoPlanejado)}%
          </strong>
        </div>
        <div className="px-5 py-3">
          <span className="block text-[10px] uppercase text-stone-500">Saldo a distribuir</span>
          <strong className={`font-mono text-sm ${Math.abs(saldoPlanejado) < 0.01 ? "text-emerald-700" : "text-amber-700"}`}>
            R$ {fmt(saldoPlanejado)}
          </strong>
        </div>
      </div>

      {linhasCalculadas.length === 0 ? (
        <div className="px-5 py-12 text-center text-sm text-stone-500">
          Nenhuma etapa cadastrada neste orçamento.
        </div>
      ) : (
        <div
          ref={tabelaScrollRef}
          onScroll={sincronizarComBarraFixa}
          className="overflow-x-auto"
        >
          <table
            className="w-full border-separate border-spacing-0 text-xs"
            style={{ minWidth: `${650 + semanas * 118}px` }}
          >
            <thead>
              <tr className="bg-stone-100 text-stone-600">
                <th className="sticky left-0 z-20 bg-stone-100 min-w-[270px] px-3 py-3 text-left border-b border-r border-stone-200 font-semibold">
                  Etapa
                </th>
                <th className="min-w-[120px] px-3 py-3 text-right border-b border-r border-stone-200 font-semibold">
                  Venda
                </th>
                <th className="min-w-[88px] px-3 py-3 text-right border-b border-r border-stone-200 font-semibold">
                  Peso físico
                </th>
                <th className="min-w-[72px] px-2 py-3 text-center border-b border-r border-stone-200 font-semibold">
                  Início
                </th>
                <th className="min-w-[72px] px-2 py-3 text-center border-b border-r border-stone-200 font-semibold">
                  Duração
                </th>
                {Array.from({ length: semanas }, (_, indice) => (
                  <th
                    key={indice}
                    className="min-w-[118px] px-2 py-2 text-center border-b border-r border-stone-200 font-semibold"
                  >
                    <span className="block">Semana {indice + 1}</span>
                    {rotuloSemana(indice) && (
                      <span className="block mt-0.5 text-[9px] font-normal text-stone-500">
                        {rotuloSemana(indice)}
                      </span>
                    )}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {linhasCalculadas.map((linha) => (
                <tr key={linha.id} className="group">
                  <td className="sticky left-0 z-10 bg-white group-hover:bg-stone-50 px-3 py-3 border-b border-r border-stone-200">
                    <span className="block font-semibold text-stone-800">
                      {linha.numero}. {linha.nome}
                    </span>
                    <span className={`block mt-1 font-mono text-[10px] ${Math.abs(linha.somaPercentuais - 100) < 0.01 ? "text-emerald-700" : "text-amber-700"}`}>
                      Distribuído: {fmt(linha.somaPercentuais)}%
                    </span>
                  </td>
                  <td className="px-3 py-3 text-right font-mono text-stone-700 border-b border-r border-stone-200 group-hover:bg-stone-50">
                    R$ {fmt(linha.totalVenda)}
                  </td>
                  <td className="px-3 py-3 text-right font-mono text-stone-700 border-b border-r border-stone-200 group-hover:bg-stone-50">
                    {fmt(linha.pesoFisico)}%
                  </td>
                  <td className="px-2 py-3 border-b border-r border-stone-200 group-hover:bg-stone-50">
                    <input
                      type="number"
                      min="1"
                      max={semanas}
                      value={linha.inicio}
                      onChange={(evento) =>
                        alterarPeriodo(linha, "inicio", evento.target.value)
                      }
                      className="w-full h-8 border border-stone-300 rounded px-1 text-center font-mono bg-white"
                      aria-label={`Semana inicial de ${linha.nome}`}
                    />
                  </td>
                  <td className="px-2 py-3 border-b border-r border-stone-200 group-hover:bg-stone-50">
                    <input
                      type="number"
                      min="1"
                      max={semanas - linha.inicio + 1}
                      value={linha.duracao}
                      onChange={(evento) =>
                        alterarPeriodo(linha, "duracao", evento.target.value)
                      }
                      className="w-full h-8 border border-stone-300 rounded px-1 text-center font-mono bg-white"
                      aria-label={`Duração de ${linha.nome}`}
                    />
                  </td>
                  {linha.percentuais.map((percentual, indice) => {
                    const semana = indice + 1;
                    const ativa =
                      semana >= linha.inicio &&
                      semana < linha.inicio + linha.duracao;
                    return (
                      <td
                        key={semana}
                        className={`px-2 py-2 border-b border-r border-stone-200 text-right ${
                          ativa ? "bg-emerald-50/50" : "bg-stone-50/70"
                        }`}
                      >
                        {ativa ? (
                          <>
                            <div className="flex items-center justify-end gap-1">
                              <input
                                type="number"
                                min="0"
                                max="100"
                                step="0.01"
                                value={Number(percentual.toFixed(6))}
                                onChange={(evento) =>
                                  alterarPercentual(
                                    linha,
                                    semana,
                                    evento.target.value
                                  )
                                }
                                className="w-16 h-7 border border-emerald-200 rounded px-1 text-right font-mono bg-white"
                                aria-label={`Percentual de ${linha.nome} na semana ${semana}`}
                              />
                              <span className="text-stone-400">%</span>
                            </div>
                            <span className="block mt-1 font-mono text-[10px] text-stone-600">
                              R$ {fmt(linha.valoresSemanais[indice])}
                            </span>
                          </>
                        ) : (
                          <span className="text-stone-300">-</span>
                        )}
                      </td>
                    );
                  })}
                </tr>
              ))}

              <tr className="bg-stone-100 font-semibold text-stone-800">
                <td className="sticky left-0 z-10 bg-stone-100 px-3 py-3 border-b border-r border-stone-300">
                  Total semanal
                </td>
                <td className="px-3 py-3 text-right font-mono border-b border-r border-stone-300">
                  R$ {fmt(totalPlanejado)}
                </td>
                <td className="px-3 py-3 text-right font-mono border-b border-r border-stone-300">
                  {fmt(fisicoPlanejado)}%
                </td>
                <td className="border-b border-r border-stone-300" />
                <td className="border-b border-r border-stone-300" />
                {totaisSemanais.map((valor, indice) => (
                  <td
                    key={indice}
                    className="px-2 py-2 text-right border-b border-r border-stone-300"
                  >
                    <span className="block font-mono text-[10px]">
                      {fmt(fisicoSemanal[indice])}%
                    </span>
                    <span className="block mt-1 font-mono">
                      R$ {fmt(valor)}
                    </span>
                  </td>
                ))}
              </tr>

              <tr className="bg-stone-900 font-semibold text-white">
                <td className="sticky left-0 z-10 bg-stone-900 px-3 py-3 border-r border-stone-700">
                  Acumulado
                </td>
                <td className="px-3 py-3 text-right font-mono border-r border-stone-700">
                  R$ {fmt(totalPlanejado)}
                </td>
                <td className="px-3 py-3 text-right font-mono border-r border-stone-700">
                  {fmt(fisicoPlanejado)}%
                </td>
                <td className="border-r border-stone-700" />
                <td className="border-r border-stone-700" />
                {acumuladoFinanceiro.map((valor, indice) => (
                  <td
                    key={indice}
                    className="px-2 py-2 text-right border-r border-stone-700"
                  >
                    <span className="block font-mono text-[10px] text-emerald-300">
                      {fmt(acumuladoFisico[indice])}%
                    </span>
                    <span className="block mt-1 font-mono text-amber-300">
                      R$ {fmt(valor)}
                    </span>
                  </td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>
      )}

      {barraFixa.visivel && (
        <div
          className="fixed bottom-2 z-50 pointer-events-none"
          style={{
            left: `${barraFixa.esquerda}px`,
            width: `${barraFixa.largura}px`,
          }}
        >
          <div className="rounded-md border border-stone-300 bg-white/95 px-1 pt-1 shadow-lg backdrop-blur-sm pointer-events-auto">
            <div
              ref={barraFixaScrollRef}
              onScroll={sincronizarComTabela}
              className="h-5 overflow-x-scroll overflow-y-hidden"
              aria-label="Rolagem horizontal fixa do cronograma"
            >
              <div
                className="h-px"
                style={{ width: `${barraFixa.larguraConteudo}px` }}
              />
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function BdiInput({ label, value, amount, onChange, maxPercent }) {
  return (
    <div className="grid grid-cols-[minmax(0,1fr)_5rem_7.5rem] items-center gap-2">
      <span className="text-stone-600 min-w-0">{label}</span>
      <input type="number" min="0" max={maxPercent} step="any" value={value === 0 ? "" : num(value) * 100} onChange={(e) => onChange(e.target.value === "" ? 0 : num(e.target.value) / 100)} className="w-20 border border-stone-300 rounded px-2 py-1 text-right font-mono" placeholder="0.00" />
      <span className="text-right font-mono text-[10px] text-stone-500 whitespace-nowrap" title="Valor correspondente ao índice">
        R$ {fmt(amount)}
      </span>
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
      itensAtivosDaEtapa(e).forEach((it) => {
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
          const itensAtivos = itensAtivosDaEtapa(e);
          const custoEtapa = itensAtivos.reduce((s, it) => s + num(it.quantidade) * cpuValorUnit(it.insumos, cpus, catalogMap), 0);
          return (
            <div key={e.id} className="border border-stone-100 rounded-lg overflow-hidden">
              <div className="bg-stone-50/50 px-4 py-2 flex justify-between text-xs font-semibold text-stone-700">
                <span>{e.nome}</span>
                <span className="font-mono">R$ {fmt(custoEtapa * FatorBdi)}</span>
              </div>
              <div className="divide-y divide-stone-50">
                {itensAtivos.map((it) => {
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



