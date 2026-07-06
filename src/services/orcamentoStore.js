import { doc, getDoc, setDoc, writeBatch } from "firebase/firestore";
import { db } from "../firebase";
import { sanitize } from "../utils/format";

const ROOT_COLLECTION = "orcacpu";
const CPUS_COLLECTION = "orcacpu_cpus";
const CPU_BATCH_SIZE = 400;
const SAVE_TIMEOUT_MS = 45000;

const cpuHash = (cpu) => JSON.stringify(sanitize(cpu));
const buildCpuHashes = (cpus = []) =>
  Object.fromEntries((cpus || []).map((cpu) => [cpu.id, cpuHash(cpu)]));

const withTimeout = (promise, label) =>
  Promise.race([
    promise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(`${label} demorou demais. Verifique a conexão e tente novamente.`)), SAVE_TIMEOUT_MS)
    ),
  ]);

export async function loadOrcamentoData() {
  const [snapProjetos, snapMeta, snapPrecos] = await Promise.all([
    getDoc(doc(db, ROOT_COLLECTION, "projetos")),
    getDoc(doc(db, ROOT_COLLECTION, "cpus_meta")),
    getDoc(doc(db, ROOT_COLLECTION, "precos")),
  ]);

  if (!snapProjetos.exists()) {
    return { empty: true };
  }

  const cpuIds = snapMeta.exists() ? snapMeta.data().cpuIds || [] : [];
  const cpuSnaps = await Promise.all(cpuIds.map((id) => getDoc(doc(db, CPUS_COLLECTION, id))));
  const cpus = cpuSnaps.filter((s) => s.exists()).map((s) => s.data());
  const precosData = snapPrecos.exists() ? snapPrecos.data() : {};

  return {
    empty: false,
    cpus,
    cpuHashes: buildCpuHashes(cpus),
    projetos: snapProjetos.data().projetos || [],
    precos: precosData.precos || [],
    projetoAtivoId: snapProjetos.data().projetoAtivoId || "",
  };
}

export async function saveOrcamentoData({
  cpus,
  projetos,
  precos,
  projetoAtivoId,
  previousCpuHashes = {},
  includeCpus = false,
}) {
  await withTimeout(
    Promise.all([
      setDoc(doc(db, ROOT_COLLECTION, "projetos"), sanitize({ projetos, projetoAtivoId })),
      setDoc(doc(db, ROOT_COLLECTION, "precos"), sanitize({ precos })),
    ]),
    "Salvamento do orçamento"
  );

  if (!includeCpus) {
    return {
      cpuHashes: previousCpuHashes,
      cpusSalvas: 0,
    };
  }

  const nextCpuHashes = buildCpuHashes(cpus);
  const changedCpus = (cpus || []).filter((cpu) => previousCpuHashes[cpu.id] !== nextCpuHashes[cpu.id]);

  await withTimeout(
    setDoc(doc(db, ROOT_COLLECTION, "cpus_meta"), sanitize({ cpuIds: cpus.map((c) => c.id) })),
    "Salvamento do índice da base de CPUs"
  );

  for (let i = 0; i < changedCpus.length; i += CPU_BATCH_SIZE) {
    const batch = writeBatch(db);
    changedCpus.slice(i, i + CPU_BATCH_SIZE).forEach((cpu) => {
      batch.set(doc(db, CPUS_COLLECTION, cpu.id), sanitize(cpu));
    });
    await withTimeout(batch.commit(), "Salvamento da base de CPUs");
  }

  return {
    cpuHashes: nextCpuHashes,
    cpusSalvas: changedCpus.length,
  };
}
