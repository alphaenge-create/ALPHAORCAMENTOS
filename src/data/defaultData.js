import { uid } from "../utils/format";

export const TIPOS = [
  { v: "MO", label: "Mao de obra" },
  { v: "MAT", label: "Material" },
  { v: "EQUIP", label: "Equipamento" },
  { v: "OUTROS", label: "Outros" },
];

export const FONTES_PADRAO = ["SUDECAP", "DER-MG", "SEINFRA", "SINAPI", "Propria"];

export const seedCpus = () => [
  {
    id: uid(),
    codigo: "SUDECAP 04.20.0010",
    fonte: "SUDECAP",
    descricao: "Alvenaria de vedacao em bloco ceramico 9x19x19cm, assentado com argamassa",
    unidade: "m2",
    insumos: [
      { id: uid(), tipo: "MO", descricao: "Pedreiro", unidade: "h", coeficiente: 0.7, valorUnitario: "" },
      { id: uid(), tipo: "MO", descricao: "Servente", unidade: "h", coeficiente: 0.45, valorUnitario: "" },
      { id: uid(), tipo: "MAT", descricao: "Bloco ceramico 9x19x19", unidade: "un", coeficiente: 25, valorUnitario: "" },
      { id: uid(), tipo: "MAT", descricao: "Argamassa de assentamento", unidade: "m3", coeficiente: 0.012, valorUnitario: "" },
    ],
  },
];

export const createDefaultProject = () => {
  const pId = uid();
  return {
    id: pId,
    nome: "Orcamento Padrao Inicial",
    criadoEm: new Date().toISOString(),
    atualizadoEm: "",
    cliente: "",
    clienteCadastro: {
      nome: "",
      local: "",
      contato: "",
      telefone: "",
      email: "",
      documento: "",
      endereco: "",
      numeroProposta: "",
      modeloProposta: "",
      regimeMateriais: "alpha",
      prazoExecucao: "",
      condicoesPagamento: "",
      percentualSinalCollem: 20,
      textoApresentacaoCollem: "",
      naoInclusosCollem: "",
      condicoesEspeciaisCollem: "",
      responsavelCollem: "Geraldo Belloni Perez",
      responsabilidadesAlpha: "",
      responsabilidadesCliente: "",
      observacoes: "",
    },
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
    },
  };
};

