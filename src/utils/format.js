export const uid = () =>
  typeof crypto !== "undefined" && crypto.randomUUID
    ? crypto.randomUUID()
    : Math.random().toString(36).slice(2, 9);

export const norm = (s) =>
  String(s ?? "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

export const fmt = (n) =>
  (Number(n) || 0).toLocaleString("pt-BR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

export const num = (v) => (v === "" || v === null || v === undefined ? 0 : Number(v));

export const avaliarExpressaoNumerica = (valor) => {
  let expressao = String(valor ?? "").trim();
  if (!expressao) return null;
  if (expressao.startsWith("=")) expressao = expressao.slice(1);
  expressao = expressao.replace(/[xX×]/g, "*").replace(/,/g, ".");
  if (!expressao || /[^0-9+\-*/().\s]/.test(expressao)) return null;

  let posicao = 0;
  const ignorarEspacos = () => {
    while (/\s/.test(expressao[posicao] || "")) posicao += 1;
  };

  const lerNumero = () => {
    ignorarEspacos();
    const match = expressao.slice(posicao).match(/^(?:\d+(?:\.\d*)?|\.\d+)/);
    if (!match) return null;
    posicao += match[0].length;
    return Number(match[0]);
  };

  const lerFator = () => {
    ignorarEspacos();
    const operador = expressao[posicao];
    if (operador === "+" || operador === "-") {
      posicao += 1;
      const fator = lerFator();
      if (fator === null) return null;
      return operador === "-" ? -fator : fator;
    }

    if (operador === "(") {
      posicao += 1;
      const valorInterno = lerSoma();
      ignorarEspacos();
      if (valorInterno === null || expressao[posicao] !== ")") return null;
      posicao += 1;
      return valorInterno;
    }

    return lerNumero();
  };

  const lerProduto = () => {
    let resultado = lerFator();
    if (resultado === null) return null;

    while (true) {
      ignorarEspacos();
      const operador = expressao[posicao];
      if (operador !== "*" && operador !== "/") break;
      posicao += 1;
      const proximo = lerFator();
      if (proximo === null) return null;
      resultado = operador === "*" ? resultado * proximo : resultado / proximo;
      if (!Number.isFinite(resultado)) return null;
    }
    return resultado;
  };

  function lerSoma() {
    let resultado = lerProduto();
    if (resultado === null) return null;

    while (true) {
      ignorarEspacos();
      const operador = expressao[posicao];
      if (operador !== "+" && operador !== "-") break;
      posicao += 1;
      const proximo = lerProduto();
      if (proximo === null) return null;
      resultado = operador === "+" ? resultado + proximo : resultado - proximo;
    }
    return resultado;
  }

  const resultado = lerSoma();
  ignorarEspacos();
  if (resultado === null || posicao !== expressao.length || !Number.isFinite(resultado)) {
    return null;
  }
  if (Object.is(resultado, -0)) return 0;
  return Number(resultado.toPrecision(15));
};

export const sanitize = (v) => {
  if (v === undefined) return null;
  if (v === null) return null;
  if (Array.isArray(v)) return v.map(sanitize);
  if (v instanceof Date) return v;
  if (typeof v === "object") {
    const out = {};
    for (const k of Object.keys(v)) {
      const s = sanitize(v[k]);
      if (s !== undefined) out[k] = s;
    }
    return out;
  }
  return v;
};
