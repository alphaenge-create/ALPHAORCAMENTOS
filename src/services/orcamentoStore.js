import { doc, getDoc, setDoc, writeBatch } from "firebase/firestore";
import { db } from "../firebase";
import { sanitize } from "../utils/format";

const ROOT_COLLECTION = "orcacpu";
const CPUS_COLLECTION = "orcacpu_cpus";
const CHUNKS_COLLECTION = "orcacpu_chunks";
const BATCH_SIZE = 400;
const SAVE_TIMEOUT_MS = 90000;
const CHUNK_SIZE = 250000;
const CHUNK_ENCODING = "gzip-base64";

const cpuHash = (cpu) => JSON.stringify(sanitize(cpu));
const buildCpuHashes = (cpus = []) =>
  Object.fromEntries((cpus || []).map((cpu) => [cpu.id, cpuHash(cpu)]));

const withTimeout = (promise, label) =>
  Promise.race([
    promise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(`${label} demorou demais. Verifique a conexao e tente novamente.`)), SAVE_TIMEOUT_MS)
    ),
  ]);

const splitText = (text) => {
  const chunks = [];
  for (let i = 0; i < text.length; i += CHUNK_SIZE) {
    chunks.push(text.slice(i, i + CHUNK_SIZE));
  }
  return chunks;
};

const bytesToBase64 = (bytes) => {
  let binary = "";
  const blockSize = 0x8000;
  for (let i = 0; i < bytes.length; i += blockSize) {
    binary += String.fromCharCode(...bytes.slice(i, i + blockSize));
  }
  return btoa(binary);
};

const base64ToBytes = (base64) => {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
};

async function gzipText(text) {
  if (typeof CompressionStream === "undefined") return { payload: text, encoding: "plain" };

  const stream = new Blob([new TextEncoder().encode(text)])
    .stream()
    .pipeThrough(new CompressionStream("gzip"));
  const compressed = new Uint8Array(await new Response(stream).arrayBuffer());
  return { payload: bytesToBase64(compressed), encoding: CHUNK_ENCODING };
}

async function gunzipText(payload, encoding) {
  if (encoding !== CHUNK_ENCODING) return payload;
  if (typeof DecompressionStream === "undefined") {
    throw new Error("Este navegador nao suporta descompactar dados salvos.");
  }

  const bytes = base64ToBytes(payload);
  const stream = new Blob([bytes])
    .stream()
    .pipeThrough(new DecompressionStream("gzip"));
  return await new Response(stream).text();
}

async function writeChunkedData(name, data) {
  const { payload, encoding } = await gzipText(JSON.stringify(sanitize(data)));
  const chunks = splitText(payload);

  await setDoc(doc(db, ROOT_COLLECTION, `${name}_meta_v2`), {
    chunkCount: chunks.length,
    encoding,
    updatedAt: new Date().toISOString(),
  });

  for (let i = 0; i < chunks.length; i += BATCH_SIZE) {
    const batch = writeBatch(db);
    chunks.slice(i, i + BATCH_SIZE).forEach((chunk, offset) => {
      batch.set(doc(db, CHUNKS_COLLECTION, `${name}_${i + offset}`), { chunk });
    });
    await batch.commit();
  }
}

async function readChunkedData(name) {
  const metaSnap = await getDoc(doc(db, ROOT_COLLECTION, `${name}_meta_v2`));
  if (!metaSnap.exists()) return null;

  const meta = metaSnap.data();
  const chunkCount = meta.chunkCount || 0;
  if (!chunkCount) return null;

  const snaps = await Promise.all(
    Array.from({ length: chunkCount }, (_, index) => getDoc(doc(db, CHUNKS_COLLECTION, `${name}_${index}`)))
  );
  const payload = snaps.map((snap) => snap.data()?.chunk || "").join("");
  const json = payload ? await gunzipText(payload, meta.encoding) : "";
  return json ? JSON.parse(json) : null;
}

export async function loadOrcamentoData() {
  const [chunkedProjetos, chunkedPrecos, snapProjetos, snapMeta, snapPrecos] = await Promise.all([
    readChunkedData("projetos"),
    readChunkedData("precos"),
    getDoc(doc(db, ROOT_COLLECTION, "projetos")),
    getDoc(doc(db, ROOT_COLLECTION, "cpus_meta")),
    getDoc(doc(db, ROOT_COLLECTION, "precos")),
  ]);

  if (!chunkedProjetos && !snapProjetos.exists()) {
    return { empty: true };
  }

  const cpuIds = snapMeta.exists() ? snapMeta.data().cpuIds || [] : [];
  const cpuSnaps = await Promise.all(cpuIds.map((id) => getDoc(doc(db, CPUS_COLLECTION, id))));
  const cpus = cpuSnaps.filter((s) => s.exists()).map((s) => s.data());
  const projetosData = chunkedProjetos || snapProjetos.data() || {};
  const precosData = chunkedPrecos || (snapPrecos.exists() ? snapPrecos.data() : {});

  return {
    empty: false,
    cpus,
    cpuHashes: buildCpuHashes(cpus),
    projetos: projetosData.projetos || [],
    precos: precosData.precos || [],
    projetoAtivoId: projetosData.projetoAtivoId || "",
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
      writeChunkedData("projetos", { projetos, projetoAtivoId }),
      writeChunkedData("precos", { precos }),
    ]),
    "Salvamento do orcamento"
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
    "Salvamento do indice da base de CPUs"
  );

  for (let i = 0; i < changedCpus.length; i += BATCH_SIZE) {
    const batch = writeBatch(db);
    changedCpus.slice(i, i + BATCH_SIZE).forEach((cpu) => {
      batch.set(doc(db, CPUS_COLLECTION, cpu.id), sanitize(cpu));
    });
    await withTimeout(batch.commit(), "Salvamento da base de CPUs");
  }

  return {
    cpuHashes: nextCpuHashes,
    cpusSalvas: changedCpus.length,
  };
}
