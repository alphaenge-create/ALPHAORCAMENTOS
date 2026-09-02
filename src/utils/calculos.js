import { norm, normalizarBusca, num } from "./format.js";

export const precoKey = (descricao) => norm(descricao);

const subCpuMatchText = (value) =>
  norm(value)
    .replace(/\bref\s*\d+\b/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const subCpuLookupCache = new WeakMap();

const addLookupEntry = (map, prefix, value, cpu) => {
  if (!value) return;
  const key = `${prefix}:${value}`;
  if (!map.has(key)) map.set(key, cpu);
};

const getSubCpuLookup = (cpusArray = []) => {
  if (!Array.isArray(cpusArray)) return new Map();
  const cached = subCpuLookupCache.get(cpusArray);
  if (cached) return cached;

  const map = new Map();
  cpusArray.forEach((cpu) => {
    addLookupEntry(map, "desc", subCpuMatchText(cpu.descricao), cpu);
    addLookupEntry(map, "codigo", norm(cpu.codigo || ""), cpu);
  });
  subCpuLookupCache.set(cpusArray, map);
  return map;
};

export const findSubCpu = (insumo, cpusArray = []) => {
  const lookup = getSubCpuLookup(cpusArray);
  const descricao = subCpuMatchText(insumo.descricao);
  const descricaoComoCodigo = norm(insumo.descricao || "");
  const codigo = norm(insumo.codigo || "");

  return (
    (descricao ? lookup.get(`desc:${descricao}`) : null) ||
    (descricaoComoCodigo ? lookup.get(`codigo:${descricaoComoCodigo}`) : null) ||
    (codigo ? lookup.get(`codigo:${codigo}`) : null) ||
    null
  );
};

export const insumosResolvidosSubCpu = (insumo, subCpu) =>
  Array.isArray(insumo?.subCpuInsumos)
    ? insumo.subCpuInsumos
    : (subCpu?.insumos || []);

export const criarIndiceBuscaCpus = (cpusArray = []) => {
  return (cpusArray || []).map((cpu) => ({
    cpu,
    // A busca de CPU considera apenas sua própria identificação. Insumos são
    // pesquisados e apresentados separadamente para evitar resultados ambíguos.
    haystack: normalizarBusca(
      [cpu.codigo, cpu.descricao, cpu.fonte].filter(Boolean).join(" ")
    ),
  }));
};

export const insumoValorUnitario = (insumo, cpusArray = [], catMap = null, visited = new Set()) => {
  const subCpu = findSubCpu(insumo, cpusArray);

  if (subCpu || Array.isArray(insumo?.subCpuInsumos)) {
    const chave = subCpu?.id || subCpu || insumo;
    if (visited.has(chave)) return 0;
    visited.add(chave);
    const val = cpuValorUnit(
      insumosResolvidosSubCpu(insumo, subCpu),
      cpusArray,
      catMap,
      visited
    );
    visited.delete(chave);
    return val;
  }

  if (catMap) {
    const entry = catMap.get(precoKey(insumo.descricao));
    if (entry && entry.valorUnitario !== "" && entry.valorUnitario !== null && entry.valorUnitario !== undefined) {
      return num(entry.valorUnitario);
    }
  }

  return num(insumo.valorUnitario);
};

// Expande somente a composição efetivamente usada, preservando o snapshot do
// orçamento e multiplicando os coeficientes em todos os níveis. Nunca grava dados.
export const decomporInsumo = (insumo, cpus = [], catalogMap = null, multiplicador = 1, visitadas = new Set(), problemas = []) => {
  const coeficiente = num(insumo?.coeficiente) * multiplicador;
  if (!Number.isFinite(coeficiente) || coeficiente < 0) {
    problemas.push(`Coeficiente inválido: ${insumo?.descricao || "insumo sem descrição"}`);
    return [];
  }
  if (coeficiente === 0) return [];
  const subCpu = findSubCpu(insumo, cpus);
  if (subCpu || Array.isArray(insumo?.subCpuInsumos)) {
    const chave = subCpu?.id || subCpu || insumo;
    if (visitadas.has(chave)) {
      problemas.push(`Referência circular na composição: ${insumo.descricao}`);
      return [];
    }
    const caminho = new Set(visitadas).add(chave);
    const internos = insumosResolvidosSubCpu(insumo, subCpu);
    if (internos.length === 0) problemas.push(`Subcomposição sem insumos: ${insumo.descricao}`);
    return internos.flatMap((interno) => decomporInsumo(interno, cpus, catalogMap, coeficiente, caminho, problemas))
      .map((folha) => ({ ...folha, chavesComposicoes: [precoKey(insumo.descricao), ...(folha.chavesComposicoes || [])] }));
  }
  const entrada = catalogMap?.get(precoKey(insumo.descricao));
  const preco = entrada?.valorUnitario;
  const valor = preco !== "" && preco !== null && preco !== undefined ? preco : insumo.valorUnitario;
  return [{ ...insumo, coeficiente, valorUnitario: num(valor), precoAusente: valor === "" || valor === null || valor === undefined }];
};

export const cpuValorUnit = (insumos, cpusArray = [], catMap = null, visited = new Set()) => {
  return (insumos || []).reduce((s, i) => {
    return s + num(i.coeficiente) * insumoValorUnitario(i, cpusArray, catMap, visited);
  }, 0);
};

export function buildCatalog(cpus, projetos, projetoAtivoId, precos) {
  const map = new Map();

  const obterEntrada = (insumo) => {
    const key = precoKey(insumo?.descricao);
    if (!key) return null;
    if (!map.has(key)) {
      map.set(key, {
        key,
        id: key,
        tipo: insumo.tipo,
        descricao: insumo.descricao,
        unidade: insumo.unidade,
        ocorrencias: 0,
        valoresEncontrados: new Set(),
        valorUnitario: "",
      });
    }
    return map.get(key);
  };

  (cpus || []).forEach((cpu) => {
    (cpu.insumos || []).forEach(obterEntrada);
  });

  const pAtivo = (projetos || []).find((p) => p.id === projetoAtivoId);
  if (pAtivo && pAtivo.etapas) {
    const registrarInsumoAtivo = (insumo, cpusVisitadas = new Set()) => {
      const entry = obterEntrada(insumo);
      if (!entry) return;

      entry.ocorrencias += 1;
      const valor = insumo.valorUnitario;
      if (
        valor !== "" &&
        valor !== null &&
        valor !== undefined &&
        !Number.isNaN(Number(valor))
      ) {
        entry.valoresEncontrados.add(Number(valor));
      }

      const subCpu = findSubCpu(insumo, cpus);
      if (!subCpu || cpusVisitadas.has(subCpu.id)) return;

      const proximoCaminho = new Set(cpusVisitadas);
      proximoCaminho.add(subCpu.id);
      insumosResolvidosSubCpu(insumo, subCpu).forEach((subInsumo) =>
        registrarInsumoAtivo(subInsumo, proximoCaminho)
      );
    };

    pAtivo.etapas.forEach((e) => {
      (e.itens || []).forEach((it) => {
        (it.insumos || []).forEach((insumo) =>
          registrarInsumoAtivo(insumo, new Set(it.cpuId ? [it.cpuId] : []))
        );
      });
    });
  }

  (precos || []).forEach((p) => {
    const key = precoKey(p.descricao);
    if (!key) return;
    if (!map.has(key)) {
      map.set(key, {
        key,
        id: p.id || key,
        tipo: p.tipo,
        descricao: p.descricao,
        unidade: p.unidade,
        ocorrencias: 0,
        valoresEncontrados: new Set(),
        valorUnitario: "",
      });
    }
    const entry = map.get(key);
    entry.id = p.id || entry.id;
    entry.tipo = p.tipo || entry.tipo;
    entry.descricao = p.descricao || entry.descricao;
    entry.unidade = p.unidade || entry.unidade;
    entry.valorUnitario = p.valorUnitario;
  });

  return Array.from(map.values())
    .map((e) => ({
      ...e,
      divergente:
        e.valoresEncontrados.size > 1 ||
        (e.valoresEncontrados.size === 1 &&
          e.valorUnitario !== "" &&
          e.valorUnitario !== null &&
          e.valorUnitario !== undefined &&
          !e.valoresEncontrados.has(Number(e.valorUnitario))),
    }))
    .sort((a, b) => a.descricao.localeCompare(b.descricao, "pt-BR"));
}

export const applyCatalogToInsumos = (insumos, catalogMap) =>
  (insumos || []).map((i) => {
    const entry = catalogMap.get(precoKey(i.descricao));
    if (entry && entry.valorUnitario !== "" && entry.valorUnitario !== null && entry.valorUnitario !== undefined) {
      return { ...i, valorUnitario: entry.valorUnitario };
    }
    return i;
  });

export const calcBdi = (b, custoInicialOverride) => {
  const custoInicial = num(custoInicialOverride !== undefined ? custoInicialOverride : b.custoInicial);
  const ac = num(b.admCentral);
  const ct = num(b.contabilidade);
  const co = num(b.contingenciamento);
  const cf = num(b.custoFinanceiro);
  const lucro = num(b.lucro);
  const das = num(b.dasAnexoIV);
  const art = num(b.art);
  const pv = das + art;
  const numerador = (1 + ac) * (1 + ct) * (1 + co) * (1 + cf) * (1 + lucro);
  const denominador = 1 - pv;
  const FatorBdi = denominador <= 0 ? 1 : numerador / denominador;
  const bdiRate = FatorBdi - 1;
  const valorVenda = custoInicial * FatorBdi;
  return { bdiRate, FatorBdi, valorVenda };
};
