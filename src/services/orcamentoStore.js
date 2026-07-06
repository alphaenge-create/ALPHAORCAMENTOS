import { doc, getDoc, setDoc, writeBatch } from "firebase/firestore";
import { db } from "../firebase";
import { sanitize } from "../utils/format";

const ROOT_COLLECTION = "orcacpu";
const CPUS_COLLECTION = "orcacpu_cpus";
const CPU_BATCH_SIZE = 400;

const cpuHash = (cpu) => JSON.stringify(sanitize(cpu));
const buildCpuHashes = (cpus = []) =>
  Object.fromEntries((cpus || []).map((cpu) => [cpu.id, cpuHash(cpu)]));

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

export async function saveOrcamentoData({ cpus, projetos, precos, projetoAtivoId, previousCpuHashes = {} }) {
  const nextCpuHashes = buildCpuHashes(cpus);
  const changedCpus = (cpus || []).filter((cpu) => previousCpuHashes[cpu.id] !== nextCpuHashes[cpu.id]);

  await Promise.all([
    setDoc(doc(db, ROOT_COLLECTION, "projetos"), sanitize({ projetos, projetoAtivoId })),
    setDoc(doc(db, ROOT_COLLECTION, "precos"), sanitize({ precos })),
    setDoc(doc(db, ROOT_COLLECTION, "cpus_meta"), sanitize({ cpuIds: cpus.map((c) => c.id) })),
  ]);

  for (let i = 0; i < changedCpus.length; i += CPU_BATCH_SIZE) {
    const batch = writeBatch(db);
    changedCpus.slice(i, i + CPU_BATCH_SIZE).forEach((cpu) => {
      batch.set(doc(db, CPUS_COLLECTION, cpu.id), sanitize(cpu));
    });
    await batch.commit();
  }

  return {
    cpuHashes: nextCpuHashes,
    cpusSalvas: changedCpus.length,
  };
}
