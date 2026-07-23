import { norm, num } from "./format";

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

export const insumoValorUnitario = (insumo, cpusArray = [], catMap = null, visited = new Set()) => {
  const subCpu = findSubCpu(insumo, cpusArray);

  if (subCpu) {
    if (visited.has(subCpu.id)) return 0;
    visited.add(subCpu.id);
    const val = cpuValorUnit(subCpu.insumos, cpusArray, catMap, visited);
    visited.delete(subCpu.id);
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

export const cpuValorUnit = (insumos, cpusArray = [], catMap = null, visited = new Set()) => {
  return (insumos || []).reduce((s, i) => {
    return s + num(i.coeficiente) * insumoValorUnitario(i, cpusArray, catMap, visited);
  }, 0);
};

export function buildCatalog(cpus, projetos, projetoAtivoId, precos) {
  const map = new Map();

  (cpus || []).forEach((cpu) => {
    (cpu.insumos || []).forEach((i) => {
      const key = precoKey(i.descricao);
      if (!key) return;
      if (!map.has(key)) {
        map.set(key, {
          key,
          id: key,
          tipo: i.tipo,
          descricao: i.descricao,
          unidade: i.unidade,
          ocorrencias: 0,
          valoresEncontrados: new Set(),
          valorUnitario: "",
        });
      }
    });
  });

  const pAtivo = (projetos || []).find((p) => p.id === projetoAtivoId);
  if (pAtivo && pAtivo.etapas) {
    pAtivo.etapas.forEach((e) => {
      (e.itens || []).forEach((it) => {
        (it.insumos || []).forEach((i) => {
          const key = precoKey(i.descricao);
          if (!key) return;
          if (!map.has(key)) {
            map.set(key, {
              key,
              id: key,
              tipo: i.tipo,
              descricao: i.descricao,
              unidade: i.unidade,
              ocorrencias: 0,
              valoresEncontrados: new Set(),
              valorUnitario: "",
            });
          }
          const entry = map.get(key);
          entry.ocorrencias += 1;
          const v = i.valorUnitario;
          if (v !== "" && v !== null && v !== undefined && !Number.isNaN(Number(v))) {
            entry.valoresEncontrados.add(Number(v));
          }
        });
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
