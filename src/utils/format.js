export const uid = () =>
  typeof crypto !== "undefined" && crypto.randomUUID
    ? crypto.randomUUID()
    : Math.random().toString(36).slice(2, 9);

export const norm = (s) =>
  String(s ?? "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

const cp1252BytePorCaractere = new Map([
  [0x20ac, 0x80],
  [0x201a, 0x82],
  [0x0192, 0x83],
  [0x201e, 0x84],
  [0x2026, 0x85],
  [0x2020, 0x86],
  [0x2021, 0x87],
  [0x02c6, 0x88],
  [0x2030, 0x89],
  [0x0160, 0x8a],
  [0x2039, 0x8b],
  [0x0152, 0x8c],
  [0x017d, 0x8e],
  [0x2018, 0x91],
  [0x2019, 0x92],
  [0x201c, 0x93],
  [0x201d, 0x94],
  [0x2022, 0x95],
  [0x2013, 0x96],
  [0x2014, 0x97],
  [0x02dc, 0x98],
  [0x2122, 0x99],
  [0x0161, 0x9a],
  [0x203a, 0x9b],
  [0x0153, 0x9c],
  [0x017e, 0x9e],
  [0x0178, 0x9f],
]);

const repararMojibake = (valor) => {
  const texto = String(valor ?? "");
  if (!/[\u00c2\u00c3]/.test(texto)) return texto;

  const caracteres = Array.from(texto);
  let corrigido = "";

  for (let indice = 0; indice < caracteres.length; indice += 1) {
    const primeiro = caracteres[indice].codePointAt(0);
    const proximo = caracteres[indice + 1];
    if ((primeiro === 0xc2 || primeiro === 0xc3) && proximo) {
      const codigoProximo = proximo.codePointAt(0);
      const segundo =
        codigoProximo <= 0xff
          ? codigoProximo
          : cp1252BytePorCaractere.get(codigoProximo);

      if (segundo >= 0x80 && segundo <= 0xbf) {
        corrigido += String.fromCodePoint(
          ((primeiro & 0x1f) << 6) | (segundo & 0x3f)
        );
        indice += 1;
        continue;
      }
    }

    corrigido += caracteres[indice];
  }

  return corrigido;
};

export const normalizarBusca = (valor) =>
  repararMojibake(valor)
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\u00df/g, "ss")
    .replace(/\u00e6/g, "ae")
    .replace(/\u0153/g, "oe")
    .replace(/\u00f8/g, "o")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

export const fmt = (n) =>
  (Number(n) || 0).toLocaleString("pt-BR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

export const num = (v) => (v === "" || v === null || v === undefined ? 0 : Number(v));

const somenteDigitos = (valor, limite) =>
  String(valor ?? "").replace(/\D/g, "").slice(0, limite);

export const formatarCpfCnpj = (valor) => {
  const digitos = somenteDigitos(valor, 14);
  if (digitos.length <= 11) {
    return digitos
      .replace(/^(\d{3})(\d)/, "$1.$2")
      .replace(/^(\d{3})\.(\d{3})(\d)/, "$1.$2.$3")
      .replace(/\.(\d{3})(\d)/, ".$1-$2");
  }
  return digitos
    .replace(/^(\d{2})(\d)/, "$1.$2")
    .replace(/^(\d{2})\.(\d{3})(\d)/, "$1.$2.$3")
    .replace(/\.(\d{3})(\d)/, ".$1/$2")
    .replace(/(\d{4})(\d)/, "$1-$2");
};

export const formatarTelefone = (valor) => {
  const digitos = somenteDigitos(valor, 11);
  if (digitos.length <= 10) {
    return digitos
      .replace(/^(\d{2})(\d)/, "($1) $2")
      .replace(/(\d{4})(\d)/, "$1-$2");
  }
  return digitos
    .replace(/^(\d{2})(\d)/, "($1) $2")
    .replace(/(\d{5})(\d)/, "$1-$2");
};

export const formatarCep = (valor) =>
  somenteDigitos(valor, 8).replace(/^(\d{5})(\d)/, "$1-$2");

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
