export const PREFIXO_CPU_PROPRIA = "prop.";

export const proximoCodigoCpuPropria = (cpus = []) => {
  const maiorNumero = (cpus || []).reduce((maior, cpu) => {
    const codigo = String(cpu?.codigo || "").trim();
    const correspondencia = codigo.match(/^prop\.(\d+)$/i);
    if (!correspondencia) return maior;

    const numero = Number(correspondencia[1]);
    return Number.isSafeInteger(numero) ? Math.max(maior, numero) : maior;
  }, 0);

  return `${PREFIXO_CPU_PROPRIA}${String(maiorNumero + 1).padStart(4, "0")}`;
};
