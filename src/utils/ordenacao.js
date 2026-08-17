const valorVazio = (valor) =>
  valor === null || valor === undefined || String(valor).trim() === "";

export const compararValores = (valorA, valorB) => {
  const vazioA = valorVazio(valorA);
  const vazioB = valorVazio(valorB);
  if (vazioA || vazioB) {
    if (vazioA && vazioB) return 0;
    return vazioA ? 1 : -1;
  }

  if (typeof valorA === "number" || typeof valorB === "number") {
    const numeroA = Number(valorA);
    const numeroB = Number(valorB);
    if (Number.isFinite(numeroA) && Number.isFinite(numeroB)) return numeroA - numeroB;
  }

  if (typeof valorA === "boolean" || typeof valorB === "boolean") {
    return Number(Boolean(valorA)) - Number(Boolean(valorB));
  }

  return String(valorA).localeCompare(String(valorB), "pt-BR", {
    sensitivity: "base",
    numeric: true,
  });
};

export const ordenarLista = (lista = [], ordenacao, obterValor) => {
  if (!ordenacao?.key || typeof obterValor !== "function") return [...(lista || [])];
  const direcao = ordenacao.direction === "desc" ? -1 : 1;

  return [...(lista || [])]
    .map((item, indice) => ({ item, indice }))
    .sort((a, b) => {
      const valorA = obterValor(a.item, ordenacao.key);
      const valorB = obterValor(b.item, ordenacao.key);
      const vazioA = valorVazio(valorA);
      const vazioB = valorVazio(valorB);
      if (vazioA || vazioB) {
        if (vazioA && vazioB) return a.indice - b.indice;
        return vazioA ? 1 : -1;
      }
      const comparacao = compararValores(valorA, valorB);
      return comparacao === 0 ? a.indice - b.indice : comparacao * direcao;
    })
    .map(({ item }) => item);
};

export const alternarOrdenacao = (atual, key, direcaoInicial = "asc") => ({
  key,
  direction:
    atual?.key === key
      ? atual.direction === "asc"
        ? "desc"
        : "asc"
      : direcaoInicial,
});
