import React from "react";
import { ArrowDown, ArrowUp, ArrowUpDown } from "lucide-react";

export function IndicadorOrdenacao({ coluna, ordenacao, size = 12 }) {
  if (ordenacao?.key !== coluna) {
    return <ArrowUpDown size={size} className="shrink-0 text-stone-300" />;
  }
  return ordenacao.direction === "asc" ? (
    <ArrowUp size={size} className="shrink-0 text-stone-700" />
  ) : (
    <ArrowDown size={size} className="shrink-0 text-stone-700" />
  );
}

export function BotaoOrdenacao({
  coluna,
  ordenacao,
  onOrdenar,
  children,
  align = "left",
  className = "",
  direcaoInicial = "asc",
}) {
  const alinhamento =
    align === "right"
      ? "justify-end text-right"
      : align === "center"
        ? "justify-center text-center"
        : "justify-start text-left";

  return (
    <button
      type="button"
      onClick={() => onOrdenar(coluna, direcaoInicial)}
      className={`inline-flex w-full items-center gap-1 rounded-sm hover:text-stone-900 ${alinhamento} ${className}`}
      title="Clique para alternar a ordenação"
    >
      <span>{children}</span>
      <IndicadorOrdenacao coluna={coluna} ordenacao={ordenacao} />
    </button>
  );
}

export function CabecalhoOrdenavel({ className = "", ...props }) {
  return (
    <th className={className}>
      <BotaoOrdenacao {...props} />
    </th>
  );
}
