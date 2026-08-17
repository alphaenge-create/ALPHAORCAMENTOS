import React, { useMemo, useState } from "react";
import { alternarOrdenacao, ordenarLista } from "./utils/ordenacao";
import { CabecalhoOrdenavel } from "./components/Ordenacao";
import * as XLSX from "xlsx-js-style";
import {
  ChevronDown,
  ChevronRight,
  Download,
  RefreshCw,
  Users,
} from "lucide-react";
import { calcularMaoDeObraPorEtapa } from "./utils/maoObra";
import { fmt, num } from "./utils/format";

const CORES_FUNCOES = [
  "#55753a",
  "#2563eb",
  "#d97706",
  "#dc2626",
  "#0891b2",
  "#7c3aed",
  "#be185d",
  "#4d7c0f",
  "#0f766e",
  "#9333ea",
  "#c2410c",
  "#475569",
];

const limitarInteiro = (valor, minimo, maximo, padrao = minimo) => {
  const numero = Math.round(Number(valor));
  if (!Number.isFinite(numero)) return padrao;
  return Math.min(maximo, Math.max(minimo, numero));
};

const distribuirPercentuais = (inicio, duracao) => {
  const percentuais = {};
  const base = Number((100 / duracao).toFixed(6));

  for (let indice = 0; indice < duracao; indice += 1) {
    const semana = inicio + indice;
    percentuais[String(semana)] =
      indice === duracao - 1
        ? Number((100 - base * (duracao - 1)).toFixed(6))
        : base;
  }

  return percentuais;
};

const dataLocal = (valor) => {
  const partes = String(valor || "").split("-").map(Number);
  if (partes.length !== 3 || partes.some((parte) => !Number.isFinite(parte))) {
    return null;
  }
  return new Date(partes[0], partes[1] - 1, partes[2]);
};

const somarDias = (data, dias) => {
  const resultado = new Date(data);
  resultado.setDate(resultado.getDate() + dias);
  return resultado;
};

const rotuloPeriodo = (dataInicio, indice) => {
  if (!dataInicio) return "";
  const inicio = somarDias(dataInicio, indice * 7);
  const fim = somarDias(inicio, 6);
  const formatar = (data) =>
    data.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" });
  return `${formatar(inicio)} a ${formatar(fim)}`;
};

const nomeArquivoSeguro = (valor) =>
  String(valor || "Orcamento")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const estilizarCabecalho = (planilha, linha, quantidadeColunas) => {
  for (let coluna = 0; coluna < quantidadeColunas; coluna += 1) {
    const endereco = XLSX.utils.encode_cell({ r: linha, c: coluna });
    if (!planilha[endereco]) continue;
    planilha[endereco].s = {
      font: { name: "Aptos", sz: 10, bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "55753A" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: {
        bottom: { style: "thin", color: { rgb: "D6D3D1" } },
        right: { style: "thin", color: { rgb: "D6D3D1" } },
      },
    };
  }
};

export default function HistogramaMaoObra({
  projeto,
  etapas,
  cronograma,
  setCronograma,
  cpus,
}) {
  const [modo, setModo] = useState("equivalente");
  const [etapasAbertas, setEtapasAbertas] = useState({});
  const [ordenacaoFuncoes, setOrdenacaoFuncoes] = useState({ key: "horas", direction: "desc" });
  const [ordenacaoEtapas, setOrdenacaoEtapas] = useState({ key: "numero", direction: "asc" });
  const semanas = limitarInteiro(cronograma.semanas, 1, 104, 12);
  const horasPorSemana = Math.max(1, num(cronograma.horasSemana) || 44);
  const inicioProjeto = dataLocal(cronograma.dataInicio);

  const etapasCalculadas = useMemo(() => {
    const maoDeObraPorEtapa = new Map(
      calcularMaoDeObraPorEtapa(etapas, cpus).map((etapa) => [
        etapa.id,
        etapa.profissionais,
      ])
    );

    return (etapas || []).map((etapa, indiceEtapa) => {
        const id = etapa.id || `etapa-${indiceEtapa}`;
        const listaProfissionais = maoDeObraPorEtapa.get(id) || [];

        const configuracao = cronograma.etapas?.[id] || {};
        const inicioPadrao = Math.min(indiceEtapa + 1, semanas);
        const inicio = limitarInteiro(
          configuracao.inicio,
          1,
          semanas,
          inicioPadrao
        );
        const duracao = limitarInteiro(
          configuracao.duracao,
          1,
          semanas - inicio + 1,
          1
        );
        const percentuaisSalvos = configuracao.percentuais || {};
        const percentuaisBase =
          Object.keys(percentuaisSalvos).length > 0
            ? percentuaisSalvos
            : distribuirPercentuais(inicio, duracao);
        const percentuais = Array.from({ length: semanas }, (_, indice) => {
          const semana = indice + 1;
          const ativa = semana >= inicio && semana < inicio + duracao;
          return ativa ? Math.max(0, num(percentuaisBase[String(semana)])) : 0;
        });
        const somaPercentuais = percentuais.reduce(
          (soma, percentual) => soma + percentual,
          0
        );
        const horasTotais = listaProfissionais.reduce(
          (soma, profissional) => soma + profissional.horas,
          0
        );

        const profissionaisCalculados = listaProfissionais.map(
          (profissional) => {
            const horasSemanais = percentuais.map(
              (percentual) => profissional.horas * (percentual / 100)
            );
            const equivalentes = horasSemanais.map(
              (horas) => horas / horasPorSemana
            );
            const recomendados = equivalentes.map((quantidade) =>
              quantidade > 0 ? Math.ceil(quantidade - 1e-9) : 0
            );
            return {
              ...profissional,
              horasSemanais,
              equivalentes,
              recomendados,
            };
          }
        );

        const horasSemanais = Array.from({ length: semanas }, (_, indice) =>
          profissionaisCalculados.reduce(
            (soma, profissional) => soma + profissional.horasSemanais[indice],
            0
          )
        );
        const equivalentes = horasSemanais.map(
          (horas) => horas / horasPorSemana
        );
        const recomendados = Array.from({ length: semanas }, (_, indice) =>
          profissionaisCalculados.reduce(
            (soma, profissional) => soma + profissional.recomendados[indice],
            0
          )
        );

        return {
          id,
          numero: indiceEtapa + 1,
          nome: etapa.nome || `Etapa ${indiceEtapa + 1}`,
          inicio,
          duracao,
          percentuais,
          somaPercentuais,
          horasTotais,
          horasSemanais,
          equivalentes,
          recomendados,
          profissionais: profissionaisCalculados,
        };
      });
  }, [etapas, cpus, cronograma.etapas, semanas, horasPorSemana]);

  const funcoes = useMemo(() => {
    const mapa = new Map();

    etapasCalculadas.forEach((etapa) => {
      etapa.profissionais.forEach((profissional) => {
        const atual = mapa.get(profissional.chave) || {
          chave: profissional.chave,
          descricao: profissional.descricao,
          horas: 0,
          horasSemanais: Array(semanas).fill(0),
        };
        atual.horas += profissional.horas;
        profissional.horasSemanais.forEach((horas, indice) => {
          atual.horasSemanais[indice] += horas;
        });
        mapa.set(profissional.chave, atual);
      });
    });

    return Array.from(mapa.values())
      .sort((a, b) => b.horas - a.horas)
      .map((funcao, indice) => {
        const equivalentes = funcao.horasSemanais.map(
          (horas) => horas / horasPorSemana
        );
        return {
          ...funcao,
          cor: CORES_FUNCOES[indice % CORES_FUNCOES.length],
          equivalentes,
          recomendados: equivalentes.map((quantidade) =>
            quantidade > 0 ? Math.ceil(quantidade - 1e-9) : 0
          ),
        };
      });
  }, [etapasCalculadas, semanas, horasPorSemana]);

  const funcoesOrdenadas = useMemo(
    () =>
      ordenarLista(funcoes, ordenacaoFuncoes, (funcao, key) => {
        if (key.startsWith("semana:")) {
          const indice = Number(key.split(":")[1]);
          return modo === "recomendado"
            ? funcao.recomendados[indice]
            : funcao.equivalentes[indice];
        }
        return funcao[key];
      }),
    [funcoes, ordenacaoFuncoes, modo]
  );

  const etapasOrdenadas = useMemo(
    () =>
      ordenarLista(etapasCalculadas, ordenacaoEtapas, (etapa, key) => {
        if (key === "periodo") return etapa.inicio;
        if (key === "pico") return Math.max(0, ...etapa.equivalentes);
        if (key === "equipe") return Math.max(0, ...etapa.recomendados);
        return etapa[key];
      }),
    [etapasCalculadas, ordenacaoEtapas]
  );

  const ordenarFuncoesPor = (key, direcaoInicial = "asc") =>
    setOrdenacaoFuncoes((atual) => alternarOrdenacao(atual, key, direcaoInicial));
  const ordenarEtapasPor = (key, direcaoInicial = "asc") =>
    setOrdenacaoEtapas((atual) => alternarOrdenacao(atual, key, direcaoInicial));

  const totais = useMemo(() => {
    const horasSemanais = Array.from({ length: semanas }, (_, indice) =>
      funcoes.reduce(
        (soma, funcao) => soma + funcao.horasSemanais[indice],
        0
      )
    );
    const equivalentes = horasSemanais.map(
      (horas) => horas / horasPorSemana
    );
    const recomendados = Array.from({ length: semanas }, (_, indice) =>
      funcoes.reduce(
        (soma, funcao) => soma + funcao.recomendados[indice],
        0
      )
    );
    return { horasSemanais, equivalentes, recomendados };
  }, [funcoes, semanas, horasPorSemana]);

  const horasTotais = etapasCalculadas.reduce(
    (soma, etapa) => soma + etapa.horasTotais,
    0
  );
  const horasDistribuidas = totais.horasSemanais.reduce(
    (soma, horas) => soma + horas,
    0
  );
  const picoEquivalente = Math.max(0, ...totais.equivalentes);
  const picoRecomendado = Math.max(0, ...totais.recomendados);
  const semanaPico =
    picoRecomendado > 0 ? totais.recomendados.indexOf(picoRecomendado) + 1 : 0;
  const valoresGrafico =
    modo === "recomendado" ? totais.recomendados : totais.equivalentes;
  const maximoGrafico = Math.max(1, ...valoresGrafico);

  const redefinirJornada = () => {
    setCronograma((atual) => ({ ...atual, horasSemana: 44 }));
  };

  const exportarExcel = () => {
    const cabecalhoSemanal = ["Função", "Total HH"];
    for (let semana = 1; semana <= semanas; semana += 1) {
      const periodo = rotuloPeriodo(inicioProjeto, semana - 1);
      cabecalhoSemanal.push(
        `S${semana}${periodo ? ` - ${periodo}` : ""} HH`,
        `S${semana} Equiv.`,
        `S${semana} Recom.`
      );
    }

    const linhasSemanais = funcoes.map((funcao) => {
      const linha = [funcao.descricao, funcao.horas];
      for (let indice = 0; indice < semanas; indice += 1) {
        linha.push(
          funcao.horasSemanais[indice],
          funcao.equivalentes[indice],
          funcao.recomendados[indice]
        );
      }
      return linha;
    });
    const linhaTotal = ["TOTAL", horasTotais];
    for (let indice = 0; indice < semanas; indice += 1) {
      linhaTotal.push(
        totais.horasSemanais[indice],
        totais.equivalentes[indice],
        totais.recomendados[indice]
      );
    }

    const dadosSemanais = [
      [`HISTOGRAMA DE MÃO DE OBRA - ${projeto.nome || "ORÇAMENTO"}`],
      [`Jornada considerada: ${fmt(horasPorSemana)} horas por profissional/semana`],
      [],
      cabecalhoSemanal,
      ...linhasSemanais,
      linhaTotal,
    ];
    const planilhaSemanal = XLSX.utils.aoa_to_sheet(dadosSemanais);
    planilhaSemanal["!cols"] = [
      { wch: 36 },
      { wch: 14 },
      ...Array.from({ length: semanas * 3 }, (_, indice) => ({
        wch: indice % 3 === 0 ? 14 : 12,
      })),
    ];
    planilhaSemanal["!freeze"] = { xSplit: 2, ySplit: 4 };
    estilizarCabecalho(planilhaSemanal, 3, cabecalhoSemanal.length);

    const dadosEtapas = [
      ["Etapa", "Semana inicial", "Duração", "Distribuído (%)", "Total HH", "Pico equivalente", "Equipe recomendada"],
    ];
    etapasCalculadas.forEach((etapa) => {
      dadosEtapas.push([
        `${etapa.numero}. ${etapa.nome}`,
        etapa.inicio,
        etapa.duracao,
        etapa.somaPercentuais,
        etapa.horasTotais,
        Math.max(0, ...etapa.equivalentes),
        Math.max(0, ...etapa.recomendados),
      ]);
      etapa.profissionais.forEach((profissional) => {
        dadosEtapas.push([
          `   ${profissional.descricao}`,
          "",
          "",
          "",
          profissional.horas,
          Math.max(0, ...profissional.equivalentes),
          Math.max(0, ...profissional.recomendados),
        ]);
      });
    });
    const planilhaEtapas = XLSX.utils.aoa_to_sheet(dadosEtapas);
    planilhaEtapas["!cols"] = [
      { wch: 48 },
      { wch: 16 },
      { wch: 12 },
      { wch: 16 },
      { wch: 14 },
      { wch: 18 },
      { wch: 20 },
    ];
    estilizarCabecalho(planilhaEtapas, 0, 7);

    const arquivo = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(arquivo, planilhaSemanal, "Histograma Semanal");
    XLSX.utils.book_append_sheet(arquivo, planilhaEtapas, "Resumo por Etapa");
    XLSX.writeFile(
      arquivo,
      `${nomeArquivoSeguro(projeto.nome)}_Histograma_Mao_de_Obra.xlsx`
    );
  };

  if (funcoes.length === 0) {
    return (
      <div className="bg-white border border-stone-200 shadow-sm rounded-lg p-10 text-center">
        <Users size={30} className="mx-auto text-stone-300 mb-3" />
        <h2 className="text-base font-semibold text-stone-800">
          Histograma de Mão de Obra
        </h2>
        <p className="text-sm text-stone-500 mt-1">
          Nenhum insumo de mão de obra foi encontrado nas CPUs deste orçamento.
        </p>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      <section className="bg-white border border-stone-200 shadow-sm rounded-lg overflow-hidden">
        <div className="p-5 border-b border-stone-200 flex flex-wrap items-end justify-between gap-4">
          <div>
            <h2 className="text-base font-semibold text-stone-800">
              Histograma de Mão de Obra
            </h2>
            <p className="text-xs text-stone-500">
              {funcoes.length} função(ões) em {etapasCalculadas.length} etapa(s)
            </p>
          </div>

          <div className="flex flex-wrap items-end gap-2">
            <label className="text-xs text-stone-600">
              <span className="block mb-1">Horas por profissional/semana</span>
              <input
                type="number"
                min="1"
                max="168"
                step="1"
                value={horasPorSemana}
                onChange={(evento) =>
                  setCronograma((atual) => ({
                    ...atual,
                    horasSemana: Math.max(1, num(evento.target.value) || 1),
                  }))
                }
                className="h-9 w-24 border border-stone-300 rounded px-2 text-right font-mono text-sm bg-white"
              />
            </label>
            <button
              type="button"
              onClick={redefinirJornada}
              className="h-9 w-9 inline-flex items-center justify-center border border-stone-300 rounded text-stone-600 hover:bg-stone-50"
              title="Restaurar jornada de 44 horas"
            >
              <RefreshCw size={14} />
            </button>
            <div className="h-9 inline-flex border border-stone-300 rounded overflow-hidden">
              <button
                type="button"
                onClick={() => setModo("equivalente")}
                className={`px-3 text-xs font-medium ${
                  modo === "equivalente"
                    ? "bg-stone-900 text-white"
                    : "bg-white text-stone-600 hover:bg-stone-50"
                }`}
              >
                Equivalente
              </button>
              <button
                type="button"
                onClick={() => setModo("recomendado")}
                className={`px-3 text-xs font-medium border-l border-stone-300 ${
                  modo === "recomendado"
                    ? "bg-stone-900 text-white"
                    : "bg-white text-stone-600 hover:bg-stone-50"
                }`}
              >
                Recomendado
              </button>
            </div>
            <button
              type="button"
              onClick={exportarExcel}
              className="h-9 px-3 inline-flex items-center gap-1.5 border border-emerald-700 bg-emerald-700 rounded text-xs font-medium text-white hover:bg-emerald-800"
            >
              <Download size={14} /> Excel
            </button>
          </div>
        </div>

        <div className="grid grid-cols-2 lg:grid-cols-4 border-b border-stone-200 bg-stone-50">
          <ResumoNumero
            label="Homem-hora do orçamento"
            valor={`${fmt(horasTotais)} h`}
          />
          <ResumoNumero
            label="Homem-hora distribuído"
            valor={`${fmt(horasDistribuidas)} h`}
            alerta={Math.abs(horasTotais - horasDistribuidas) > 0.01}
          />
          <ResumoNumero
            label="Pico equivalente"
            valor={`${fmt(picoEquivalente)} prof.`}
            detalhe={semanaPico ? `Semana ${semanaPico}` : ""}
          />
          <ResumoNumero
            label="Equipe recomendada no pico"
            valor={`${picoRecomendado} prof.`}
            detalhe={semanaPico ? `Semana ${semanaPico}` : ""}
            ultimo
          />
        </div>

        <div className="p-5">
          <div className="flex flex-wrap gap-x-4 gap-y-2 mb-4">
            {funcoes.map((funcao) => (
              <div
                key={funcao.chave}
                className="inline-flex items-center gap-1.5 text-[11px] text-stone-600"
              >
                <span
                  className="w-3 h-3 rounded-sm shrink-0"
                  style={{ backgroundColor: funcao.cor }}
                />
                <span>{funcao.descricao}</span>
              </div>
            ))}
          </div>

          <div className="overflow-x-auto pb-2">
            <div
              className="flex items-end gap-2 h-[290px]"
              style={{ minWidth: `${Math.max(720, semanas * 76)}px` }}
            >
              {Array.from({ length: semanas }, (_, indice) => {
                const total = valoresGrafico[indice];
                return (
                  <div
                    key={indice}
                    className="flex-1 min-w-[64px] h-full flex flex-col justify-end"
                  >
                    <div className="h-[235px] flex flex-col justify-end">
                      <span className="block text-center font-mono text-[10px] font-semibold text-stone-700 mb-1">
                        {modo === "recomendado" ? total : fmt(total)}
                      </span>
                      <div className="h-[210px] flex flex-col-reverse justify-start bg-stone-50 border-b border-stone-300">
                        {funcoes.map((funcao) => {
                          const valor =
                            modo === "recomendado"
                              ? funcao.recomendados[indice]
                              : funcao.equivalentes[indice];
                          if (valor <= 0) return null;
                          return (
                            <div
                              key={funcao.chave}
                              style={{
                                height: `${Math.max(
                                  2,
                                  (valor / maximoGrafico) * 210
                                )}px`,
                                backgroundColor: funcao.cor,
                              }}
                              title={`${funcao.descricao}: ${
                                modo === "recomendado" ? valor : fmt(valor)
                              } profissional(is)`}
                            />
                          );
                        })}
                      </div>
                    </div>
                    <div className="pt-2 text-center">
                      <span className="block text-[10px] font-semibold text-stone-700">
                        S{indice + 1}
                      </span>
                      {rotuloPeriodo(inicioProjeto, indice) && (
                        <span className="block text-[8px] text-stone-400">
                          {rotuloPeriodo(inicioProjeto, indice)}
                        </span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </section>

      <section className="bg-white border border-stone-200 shadow-sm rounded-lg overflow-hidden">
        <div className="px-5 py-4 border-b border-stone-200">
          <h3 className="text-sm font-semibold text-stone-800">
            Necessidade semanal por função
          </h3>
        </div>
        <div className="overflow-x-auto">
          <table
            className="w-full border-separate border-spacing-0 text-xs"
            style={{ minWidth: `${420 + semanas * 92}px` }}
          >
            <thead>
              <tr className="bg-stone-100 text-stone-600">
                <CabecalhoOrdenavel coluna="descricao" ordenacao={ordenacaoFuncoes} onOrdenar={ordenarFuncoesPor} className="sticky left-0 z-20 bg-stone-100 min-w-[250px] px-3 py-3 border-b border-r border-stone-200">Função</CabecalhoOrdenavel>
                <CabecalhoOrdenavel coluna="horas" ordenacao={ordenacaoFuncoes} onOrdenar={ordenarFuncoesPor} className="min-w-[110px] px-3 py-3 border-b border-r border-stone-200" align="right" direcaoInicial="desc">Total HH</CabecalhoOrdenavel>
                {Array.from({ length: semanas }, (_, indice) => (
                  <CabecalhoOrdenavel
                    key={indice}
                    coluna={`semana:${indice}`}
                    ordenacao={ordenacaoFuncoes}
                    onOrdenar={ordenarFuncoesPor}
                    direcaoInicial="desc"
                    align="center"
                    className="min-w-[92px] px-2 py-2 text-center border-b border-r border-stone-200"
                  >
                    <span className="block">S{indice + 1}</span>
                    {rotuloPeriodo(inicioProjeto, indice) && (
                      <span className="block text-[8px] font-normal text-stone-400">
                        {rotuloPeriodo(inicioProjeto, indice)}
                      </span>
                    )}
                  </CabecalhoOrdenavel>
                ))}
              </tr>
            </thead>
            <tbody>
              {funcoesOrdenadas.map((funcao) => (
                <tr key={funcao.chave} className="group">
                  <td className="sticky left-0 z-10 bg-white group-hover:bg-stone-50 px-3 py-2.5 border-b border-r border-stone-200">
                    <span
                      className="inline-block w-2.5 h-2.5 rounded-sm mr-2"
                      style={{ backgroundColor: funcao.cor }}
                    />
                    <span className="font-medium text-stone-800">
                      {funcao.descricao}
                    </span>
                  </td>
                  <td className="px-3 py-2.5 text-right font-mono border-b border-r border-stone-200">
                    {fmt(funcao.horas)}
                  </td>
                  {funcao.equivalentes.map((equivalente, indice) => (
                    <td
                      key={indice}
                      className="px-2 py-2.5 text-right font-mono border-b border-r border-stone-200"
                    >
                      {modo === "recomendado"
                        ? funcao.recomendados[indice] || "-"
                        : equivalente > 0
                          ? fmt(equivalente)
                          : "-"}
                    </td>
                  ))}
                </tr>
              ))}
              <tr className="bg-stone-900 text-white font-semibold">
                <td className="sticky left-0 z-10 bg-stone-900 px-3 py-3 border-r border-stone-700">
                  Total da equipe
                </td>
                <td className="px-3 py-3 text-right font-mono border-r border-stone-700">
                  {fmt(horasTotais)}
                </td>
                {totais.equivalentes.map((equivalente, indice) => (
                  <td
                    key={indice}
                    className="px-2 py-3 text-right font-mono text-amber-300 border-r border-stone-700"
                  >
                    {modo === "recomendado"
                      ? totais.recomendados[indice] || "-"
                      : equivalente > 0
                        ? fmt(equivalente)
                        : "-"}
                  </td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>
      </section>

      <section className="bg-white border border-stone-200 shadow-sm rounded-lg overflow-hidden">
        <div className="px-5 py-4 border-b border-stone-200">
          <h3 className="text-sm font-semibold text-stone-800">
            Resumo por etapa
          </h3>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[760px] text-xs">
            <thead>
              <tr className="bg-stone-100 text-stone-600">
                <CabecalhoOrdenavel coluna="numero" ordenacao={ordenacaoEtapas} onOrdenar={ordenarEtapasPor} className="px-4 py-3">Etapa</CabecalhoOrdenavel>
                <CabecalhoOrdenavel coluna="periodo" ordenacao={ordenacaoEtapas} onOrdenar={ordenarEtapasPor} className="px-3 py-3" align="center">Período</CabecalhoOrdenavel>
                <CabecalhoOrdenavel coluna="somaPercentuais" ordenacao={ordenacaoEtapas} onOrdenar={ordenarEtapasPor} className="px-3 py-3" align="right" direcaoInicial="desc">Distribuído</CabecalhoOrdenavel>
                <CabecalhoOrdenavel coluna="horasTotais" ordenacao={ordenacaoEtapas} onOrdenar={ordenarEtapasPor} className="px-3 py-3" align="right" direcaoInicial="desc">Total HH</CabecalhoOrdenavel>
                <CabecalhoOrdenavel coluna="pico" ordenacao={ordenacaoEtapas} onOrdenar={ordenarEtapasPor} className="px-3 py-3" align="right" direcaoInicial="desc">Pico equivalente</CabecalhoOrdenavel>
                <CabecalhoOrdenavel coluna="equipe" ordenacao={ordenacaoEtapas} onOrdenar={ordenarEtapasPor} className="px-4 py-3" align="right" direcaoInicial="desc">Equipe recomendada</CabecalhoOrdenavel>
              </tr>
            </thead>
            <tbody>
              {etapasOrdenadas.map((etapa) => {
                const aberta = !!etapasAbertas[etapa.id];
                return (
                  <React.Fragment key={etapa.id}>
                    <tr className="border-t border-stone-200 hover:bg-stone-50">
                      <td className="px-4 py-3">
                        <button
                          type="button"
                          onClick={() =>
                            setEtapasAbertas((atual) => ({
                              ...atual,
                              [etapa.id]: !aberta,
                            }))
                          }
                          className="inline-flex items-center gap-2 text-left font-medium text-stone-800"
                        >
                          {aberta ? (
                            <ChevronDown size={14} />
                          ) : (
                            <ChevronRight size={14} />
                          )}
                          {etapa.numero}. {etapa.nome}
                        </button>
                      </td>
                      <td className="px-3 py-3 text-center font-mono text-stone-600">
                        S{etapa.inicio} a S{etapa.inicio + etapa.duracao - 1}
                      </td>
                      <td className={`px-3 py-3 text-right font-mono ${
                        Math.abs(etapa.somaPercentuais - 100) < 0.01
                          ? "text-emerald-700"
                          : "text-amber-700"
                      }`}>
                        {fmt(etapa.somaPercentuais)}%
                      </td>
                      <td className="px-3 py-3 text-right font-mono">
                        {fmt(etapa.horasTotais)}
                      </td>
                      <td className="px-3 py-3 text-right font-mono">
                        {fmt(Math.max(0, ...etapa.equivalentes))}
                      </td>
                      <td className="px-4 py-3 text-right font-mono font-semibold">
                        {Math.max(0, ...etapa.recomendados)}
                      </td>
                    </tr>
                    {aberta && (
                      <tr className="bg-stone-50/70">
                        <td colSpan={6} className="px-10 py-3">
                          {etapa.profissionais.length === 0 ? (
                            <span className="text-stone-400">
                              Nenhuma mão de obra nesta etapa.
                            </span>
                          ) : (
                            <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-x-8 gap-y-2">
                              {etapa.profissionais.map((profissional) => (
                                <div
                                  key={profissional.chave}
                                  className="grid grid-cols-[1fr_auto_auto] gap-3 items-center border-b border-stone-200 py-1.5"
                                >
                                  <span className="truncate text-stone-700">
                                    {profissional.descricao}
                                  </span>
                                  <span className="font-mono text-stone-500">
                                    {fmt(profissional.horas)} HH
                                  </span>
                                  <span className="font-mono font-semibold text-stone-800">
                                    pico {Math.max(0, ...profissional.recomendados)}
                                  </span>
                                </div>
                              ))}
                            </div>
                          )}
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

function ResumoNumero({ label, valor, detalhe = "", alerta = false, ultimo = false }) {
  return (
    <div
      className={`px-5 py-3 border-b lg:border-b-0 ${
        ultimo ? "" : "border-r"
      } border-stone-200`}
    >
      <span className="block text-[10px] uppercase text-stone-500">{label}</span>
      <strong
        className={`font-mono text-sm ${
          alerta ? "text-amber-700" : "text-stone-800"
        }`}
      >
        {valor}
      </strong>
      {detalhe && (
        <span className="block text-[10px] text-stone-400 mt-0.5">{detalhe}</span>
      )}
    </div>
  );
}
