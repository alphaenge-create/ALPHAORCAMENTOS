export const gruposAlternativasDaEtapa = (etapa) =>
  Array.isArray(etapa?.gruposAlternativas) ? etapa.gruposAlternativas : [];

export const grupoAlternativaDoItem = (etapa, item) => {
  if (!item?.alternativaGrupoId) return null;
  return (
    gruposAlternativasDaEtapa(etapa).find(
      (grupo) => grupo.id === item.alternativaGrupoId
    ) || null
  );
};

export const itemIncluidoNoCalculo = (etapa, item) => {
  const grupo = grupoAlternativaDoItem(etapa, item);
  if (!grupo) return true;
  return Boolean(
    grupo.opcaoAtivaId &&
      item.alternativaOpcaoId === grupo.opcaoAtivaId &&
      (grupo.opcoes || []).some((opcao) => opcao.id === grupo.opcaoAtivaId)
  );
};

export const itensAtivosDaEtapa = (etapa) =>
  (etapa?.itens || []).filter((item) => itemIncluidoNoCalculo(etapa, item));

export const etapaComOpcaoAtiva = (etapa, grupoId, opcaoAtivaId) => ({
  ...etapa,
  gruposAlternativas: gruposAlternativasDaEtapa(etapa).map((grupo) =>
    grupo.id === grupoId ? { ...grupo, opcaoAtivaId } : grupo
  ),
});

export const etapasComOpcaoAtiva = (
  etapas,
  etapaId,
  grupoId,
  opcaoAtivaId
) =>
  (etapas || []).map((etapa) =>
    etapa.id === etapaId
      ? etapaComOpcaoAtiva(etapa, grupoId, opcaoAtivaId)
      : etapa
  );

