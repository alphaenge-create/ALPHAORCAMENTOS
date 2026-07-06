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
