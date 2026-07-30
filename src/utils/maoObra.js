import { findSubCpu, precoKey } from "./calculos";
import { itensAtivosDaEtapa } from "./alternativas";
import { norm, num } from "./format";

const ehMaoDeObra = (insumo) => {
  const tipo = norm(insumo?.tipo).trim();
  return (
    tipo === "mo" ||
    tipo === "mao de obra" ||
    tipo.includes("mao de obra")
  );
};

const chaveProfissional = (descricao) =>
  norm(descricao).replace(/\s+/g, " ").trim();

const valorHora = (insumo, catalogMap) => {
  const entrada = catalogMap?.get(precoKey(insumo?.descricao));
  if (
    entrada &&
    entrada.valorUnitario !== "" &&
    entrada.valorUnitario !== null &&
    entrada.valorUnitario !== undefined
  ) {
    return num(entrada.valorUnitario);
  }
  return num(insumo?.valorUnitario);
};

const registrarMaoDeObra = (mapa, insumo, horas, catalogMap) => {
  const descricao = String(insumo?.descricao || "").trim();
  const chave = chaveProfissional(descricao);
  if (!chave || !Number.isFinite(horas) || horas === 0) return;

  const precoHora = valorHora(insumo, catalogMap);
  const atual = mapa.get(chave) || {
    chave,
    descricao,
    unidade: insumo.unidade || "h",
    horas: 0,
    total: 0,
  };
  atual.horas += horas;
  atual.total += horas * precoHora;
  mapa.set(chave, atual);
};

const coletarMaoDeObra = (
  insumos,
  multiplicador,
  cpus,
  catalogMap,
  resultado,
  visitadas = new Set()
) => {
  (insumos || []).forEach((insumo) => {
    const horasOuQuantidade = multiplicador * num(insumo.coeficiente);
    if (!Number.isFinite(horasOuQuantidade) || horasOuQuantidade === 0) return;

    // O tipo MO tem prioridade para impedir que uma CPU homônima transforme
    // acidentalmente um profissional em subcomposição.
    if (ehMaoDeObra(insumo)) {
      registrarMaoDeObra(resultado, insumo, horasOuQuantidade, catalogMap);
      return;
    }

    const subCpu = findSubCpu(insumo, cpus);
    if (!subCpu || visitadas.has(subCpu.id)) return;

    const proximoCaminho = new Set(visitadas);
    proximoCaminho.add(subCpu.id);
    coletarMaoDeObra(
      subCpu.insumos,
      horasOuQuantidade,
      cpus,
      catalogMap,
      resultado,
      proximoCaminho
    );
  });
};

export const calcularMaoDeObraPorEtapa = (
  etapas,
  cpus,
  catalogMap = null
) =>
  (etapas || []).map((etapa, indice) => {
    const profissionais = new Map();
    itensAtivosDaEtapa(etapa).forEach((item) => {
      coletarMaoDeObra(
        item.insumos,
        num(item.quantidade),
        cpus,
        catalogMap,
        profissionais,
        new Set(item.cpuId ? [item.cpuId] : [])
      );
    });

    return {
      id: etapa.id || `etapa-${indice}`,
      profissionais: Array.from(profissionais.values())
        .map((profissional) => ({
          ...profissional,
          valorUnit:
            profissional.horas > 0
              ? profissional.total / profissional.horas
              : 0,
        }))
        .sort((a, b) => b.horas - a.horas),
    };
  });

export const consolidarMaoDeObra = (etapas, cpus, catalogMap = null) => {
  const consolidado = new Map();

  calcularMaoDeObraPorEtapa(etapas, cpus, catalogMap).forEach((etapa) => {
    etapa.profissionais.forEach((profissional) => {
      const atual = consolidado.get(profissional.chave) || {
        chave: profissional.chave,
        descricao: profissional.descricao,
        unidade: profissional.unidade,
        qtd: 0,
        total: 0,
      };
      atual.qtd += profissional.horas;
      atual.total += profissional.total;
      consolidado.set(profissional.chave, atual);
    });
  });

  return Array.from(consolidado.values())
    .map((profissional) => ({
      ...profissional,
      valorUnit:
        profissional.qtd > 0
          ? profissional.total / profissional.qtd
          : 0,
    }))
    .sort((a, b) => b.total - a.total);
};
