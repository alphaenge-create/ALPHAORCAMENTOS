import { doc, getDoc, setDoc } from "firebase/firestore";
import { db } from "../firebase";
import { sanitize } from "../utils/format";

const ROOT_COLLECTION = "orcacpu";
const CPUS_COLLECTION = "orcacpu_cpus";

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
    projetos: snapProjetos.data().projetos || [],
    precos: precosData.precos || [],
    projetoAtivoId: snapProjetos.data().projetoAtivoId || "",
  };
}

export async function saveOrcamentoData({ cpus, projetos, precos, projetoAtivoId }) {
  await Promise.all([
    setDoc(doc(db, ROOT_COLLECTION, "projetos"), sanitize({ projetos, projetoAtivoId })),
    setDoc(doc(db, ROOT_COLLECTION, "precos"), sanitize({ precos })),
    setDoc(doc(db, ROOT_COLLECTION, "cpus_meta"), sanitize({ cpuIds: cpus.map((c) => c.id) })),
  ]);

  const lote = 10;
  for (let i = 0; i < cpus.length; i += lote) {
    await Promise.all(
      cpus.slice(i, i + lote).map((c) => setDoc(doc(db, CPUS_COLLECTION, c.id), sanitize(c)))
    );
  }
}
