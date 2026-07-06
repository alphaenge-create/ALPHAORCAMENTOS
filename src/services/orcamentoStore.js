import { doc, getDoc } from "firebase/firestore";
import { db, firebaseConfig } from "../firebase";
import { sanitize } from "../utils/format";

const ROOT_COLLECTION = "orcacpu";
const CPUS_COLLECTION = "orcacpu_cpus";
const CHUNKS_COLLECTION = "orcacpu_chunks";
const PROJECT_CHUNKS_PREFIX = "projeto";
const BATCH_SIZE = 400;
const SAVE_TIMEOUT_MS = 30000;
const CHUNK_SIZE = 700000;
const CHUNK_ENCODING = "gzip-base64";

const firestoreRestUrl = (path) => {
  const encodedPath = path.split("/").map(encodeURIComponent).join("/");
  return `https://firestore.googleapis.com/v1/projects/${firebaseConfig.projectId}/databases/(default)/documents/${encodedPath}?key=${firebaseConfig.apiKey}`;
};

const firestoreValue = (value) => {
  if (value === null || value === undefined) return { nullValue: "NULL_VALUE" };
  if (typeof value === "boolean") return { booleanValue: value };
  if (typeof value === "number") {
    return Number.isInteger(value) ? { integerValue: String(value) } : { doubleValue: value };
  }
  if (typeof value === "string") return { stringValue: value };
  if (value instanceof Date) return { timestampValue: value.toISOString() };
  if (Array.isArray(value)) return { arrayValue: { values: value.map(firestoreValue) } };
  if (typeof value === "object") return { mapValue: { fields: firestoreFields(value) } };
  return { stringValue: String(value) };
};

const firestoreFields = (data) => {
  const safe = sanitize(data) || {};
  return Object.fromEntries(Object.entries(safe).map(([key, value]) => [key, firestoreValue(value)]));
};

async function restSetDoc(path, data) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), SAVE_TIMEOUT_MS);

  let response;
  try {
    response = await fetch(firestoreRestUrl(path), {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ fields: firestoreFields(data) }),
      signal: controller.signal,
    });
  } finally {
    clearTimeout(timeout);
  }

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Firestore REST ${response.status}: ${text}`);
  }
}

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

  await restSetDoc(`${ROOT_COLLECTION}/${name}_meta_v2`, {
    chunkCount: chunks.length,
    encoding,
    updatedAt: new Date().toISOString(),
  });

  for (let i = 0; i < chunks.length; i += 1) {
    await restSetDoc(`${CHUNKS_COLLECTION}/${name}_${i}`, { chunk: chunks[i] });
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

const projectChunkName = (id) => `${PROJECT_CHUNKS_PREFIX}_${id}`;
const projectEtapaChunkName = (projectId, etapaId) => `${PROJECT_CHUNKS_PREFIX}_${projectId}_etapa_${etapaId}`;

const projetoResumo = (p) => ({
  id: p.id,
  nome: p.nome,
  cliente: p.cliente || "",
});

async function loadProjetosV3() {
  const metaSnap = await getDoc(doc(db, ROOT_COLLECTION, "projetos_index_v3"));
  if (!metaSnap.exists()) return null;

  const meta = metaSnap.data();
  const ids = meta.projetoIds || [];
  const projetos = await Promise.all(
    ids.map(async (id) => {
      const projeto = await readChunkedData(projectChunkName(id));
      if (!projeto || projeto.storageVersion !== 4) return projeto;

      const etapas = await Promise.all(
        (projeto.etapaIds || []).map((etapaId) => readChunkedData(projectEtapaChunkName(id, etapaId)))
      );
      const { etapaIds, etapaResumos, storageVersion, ...baseProjeto } = projeto;
      return {
        ...baseProjeto,
        etapas: etapas.filter(Boolean),
      };
    })
  );
  const resumosPorId = new Map((meta.resumos || []).map((resumo) => [resumo.id, resumo]));
  return {
    projetos: projetos
      .filter(Boolean)
      .map((projeto) => ({
        ...projeto,
        ...(resumosPorId.get(projeto.id) || {}),
      })),
    projetoAtivoId: meta.projetoAtivoId || ids[0] || "",
  };
}

async function saveProjetosV3(projetos, projetoAtivoId) {
  const projetoAtivo = (projetos || []).find((p) => p.id === projetoAtivoId) || (projetos || [])[0];

  await withTimeout(
    restSetDoc(`${ROOT_COLLECTION}/projetos_index_v3`, {
      storageVersion: 4,
      projetoAtivoId: projetoAtivo?.id || projetoAtivoId || "",
      projetoIds: (projetos || []).map((p) => p.id),
      updatedAt: new Date().toISOString(),
    }),
    "Salvamento REST do indice dos orcamentos"
  );

  if (projetoAtivo) {
    const { etapas = [], ...projetoBase } = projetoAtivo;
    await withTimeout(
      writeChunkedData(projectChunkName(projetoAtivo.id), {
        ...projetoBase,
        storageVersion: 4,
        etapaIds: etapas.map((etapa) => etapa.id),
        etapaResumos: etapas.map((etapa) => ({ id: etapa.id, nome: etapa.nome })),
      }),
      "Salvamento dos dados principais do orcamento"
    );

    for (let i = 0; i < etapas.length; i += 1) {
      await withTimeout(
        writeChunkedData(projectEtapaChunkName(projetoAtivo.id, etapas[i].id), etapas[i]),
        `Salvamento da etapa ${i + 1}`
      );
    }
  }
}

export async function loadOrcamentoData() {
  const [projetosV3, chunkedProjetos, chunkedPrecos, snapProjetos, snapMeta, snapPrecos] = await Promise.all([
    loadProjetosV3(),
    readChunkedData("projetos"),
    readChunkedData("precos"),
    getDoc(doc(db, ROOT_COLLECTION, "projetos")),
    getDoc(doc(db, ROOT_COLLECTION, "cpus_meta")),
    getDoc(doc(db, ROOT_COLLECTION, "precos")),
  ]);

  if (!projetosV3 && !chunkedProjetos && !snapProjetos.exists()) {
    return { empty: true };
  }

  const cpuIds = snapMeta.exists() ? snapMeta.data().cpuIds || [] : [];
  const cpuSnaps = await Promise.all(cpuIds.map((id) => getDoc(doc(db, CPUS_COLLECTION, id))));
  const cpus = cpuSnaps.filter((s) => s.exists()).map((s) => s.data());
  const projetosData = projetosV3 || chunkedProjetos || snapProjetos.data() || {};
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
  await saveProjetosV3(projetos, projetoAtivoId);
  await withTimeout(writeChunkedData("precos", { precos }), "Salvamento do banco de precos");

  if (!includeCpus) {
    return {
      cpuHashes: previousCpuHashes,
      cpusSalvas: 0,
    };
  }

  const nextCpuHashes = buildCpuHashes(cpus);
  const changedCpus = (cpus || []).filter((cpu) => previousCpuHashes[cpu.id] !== nextCpuHashes[cpu.id]);

  await withTimeout(
    restSetDoc(`${ROOT_COLLECTION}/cpus_meta`, { cpuIds: cpus.map((c) => c.id) }),
    "Salvamento do indice da base de CPUs"
  );

  for (let i = 0; i < changedCpus.length; i += 1) {
    await withTimeout(restSetDoc(`${CPUS_COLLECTION}/${changedCpus[i].id}`, changedCpus[i]), "Salvamento da base de CPUs");
  }

  return {
    cpuHashes: nextCpuHashes,
    cpusSalvas: changedCpus.length,
  };
}
