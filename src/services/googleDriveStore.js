const DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.file";
const DRIVE_FOLDER_NAME = "AlphaOrcamentos";
const DRIVE_BACKUP_NAME = "alphaorc-backup.json.gz";
const DRIVE_BASE_NAME = "alphaorc-base.json.gz";
const DRIVE_PROJECT_PREFIX = "alphaorc-projeto-";
const DRIVE_PROJECT_SUFFIX = ".json.gz";
const DRIVE_STORAGE_VERSION = 2;
const CLIENT_ID_KEY = "alphaorc-google-client-id";
const DEFAULT_GOOGLE_CLIENT_ID = "376035181065-3gdnkppglnag1s1u3ue7kkflfe2iv5ln.apps.googleusercontent.com";

let tokenClient = null;
let accessToken = "";
let accessTokenExpiresAt = 0;

const hasValidAccessToken = () =>
  !!accessToken && Date.now() < accessTokenExpiresAt - 60_000;

const getClientId = () => {
  const configured = import.meta.env.VITE_GOOGLE_CLIENT_ID || DEFAULT_GOOGLE_CLIENT_ID;
  if (configured && configured !== "COLE_SEU_CLIENT_ID_AQUI") return configured;
  return localStorage.getItem(CLIENT_ID_KEY) || "";
};

export const hasGoogleDriveClientId = () => !!getClientId();

export const configureGoogleDriveClientId = () => {
  const current = getClientId();
  const value = window.prompt("Cole o OAuth Client ID do Google Drive:", current);
  if (!value) return false;
  localStorage.setItem(CLIENT_ID_KEY, value.trim());
  tokenClient = null;
  accessToken = "";
  return true;
};

const loadGoogleIdentity = () =>
  new Promise((resolve, reject) => {
    if (window.google?.accounts?.oauth2) {
      resolve();
      return;
    }

    const existing = document.querySelector('script[src="https://accounts.google.com/gsi/client"]');
    if (existing) {
      existing.addEventListener("load", resolve, { once: true });
      existing.addEventListener("error", reject, { once: true });
      return;
    }

    const script = document.createElement("script");
    script.src = "https://accounts.google.com/gsi/client";
    script.async = true;
    script.defer = true;
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
  });

export async function requestGoogleDriveAccess() {
  await loadGoogleIdentity();

  const clientId = getClientId();
  if (!clientId) {
    if (!configureGoogleDriveClientId()) throw new Error("Client ID do Google Drive nao configurado.");
  }

  if (hasValidAccessToken()) return accessToken;

  return await new Promise((resolve, reject) => {
    tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: getClientId(),
      scope: DRIVE_SCOPE,
      callback: (response) => {
        if (response?.error) {
          reject(new Error(response.error_description || response.error));
          return;
        }
        accessToken = response.access_token;
        accessTokenExpiresAt = Date.now() + Number(response.expires_in || 3600) * 1000;
        resolve(accessToken);
      },
      error_callback: (error) => {
        reject(new Error(error?.message || error?.type || "A janela de acesso ao Google Drive foi fechada."));
      },
    });
    tokenClient.requestAccessToken({ prompt: "" });
  });
}

const driveFetch = async (url, options = {}) => {
  if (!accessToken) await requestGoogleDriveAccess();
  const response = await fetch(url, {
    ...options,
    headers: {
      ...(options.headers || {}),
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (response.status === 401) {
    accessToken = "";
    accessTokenExpiresAt = 0;
    await requestGoogleDriveAccess();
    return driveFetch(url, options);
  }

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Google Drive ${response.status}: ${text}`);
  }

  return response;
};

const escapeDriveQuery = (value) => String(value).replace(/\\/g, "\\\\").replace(/'/g, "\\'");

const driveSearch = async (q, fields = "files(id,name,modifiedTime)") => {
  const params = new URLSearchParams({
    q,
    fields,
    spaces: "drive",
    pageSize: "1000",
  });
  const response = await driveFetch(`https://www.googleapis.com/drive/v3/files?${params.toString()}`);
  const data = await response.json();
  return data.files || [];
};

const findDriveFolder = async () => {
  const files = await driveSearch(
    `name='${escapeDriveQuery(DRIVE_FOLDER_NAME)}' and mimeType='application/vnd.google-apps.folder' and trashed=false`
  );
  return files[0] || null;
};

const createDriveFolder = async () => {
  const response = await driveFetch("https://www.googleapis.com/drive/v3/files?fields=id,name", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: DRIVE_FOLDER_NAME,
      mimeType: "application/vnd.google-apps.folder",
    }),
  });
  return await response.json();
};

const ensureDriveFolder = async () => (await findDriveFolder()) || (await createDriveFolder());

const findBackupFile = async (folderId) => {
  const files = await driveSearch(
    `name='${escapeDriveQuery(DRIVE_BACKUP_NAME)}' and '${folderId}' in parents and trashed=false`,
    "files(id,name,modifiedTime,size)"
  );
  return files[0] || null;
};

const findNamedFile = async (folderId, name) => {
  const files = await driveSearch(
    `name='${escapeDriveQuery(name)}' and '${folderId}' in parents and trashed=false`,
    "files(id,name,modifiedTime,size)"
  );
  return files[0] || null;
};

const projectFileName = (projectId) =>
  `${DRIVE_PROJECT_PREFIX}${String(projectId || "").replace(/[^a-zA-Z0-9_-]/g, "_")}${DRIVE_PROJECT_SUFFIX}`;

const listProjectFiles = async (folderId) => {
  const files = await driveSearch(
    `name contains '${escapeDriveQuery(DRIVE_PROJECT_PREFIX)}' and '${folderId}' in parents and trashed=false`,
    "files(id,name,modifiedTime,size)"
  );
  return files.filter(
    (file) => file.name.startsWith(DRIVE_PROJECT_PREFIX) && file.name.endsWith(DRIVE_PROJECT_SUFFIX)
  );
};

const gzipJson = async (data) => {
  const json = JSON.stringify(data);
  if (typeof CompressionStream === "undefined") {
    return new Blob([json], { type: "application/json" });
  }
  const stream = new Blob([new TextEncoder().encode(json)]).stream().pipeThrough(new CompressionStream("gzip"));
  return new Blob([await new Response(stream).arrayBuffer()], { type: "application/gzip" });
};

const readBackupBlob = async (blob) => {
  if (typeof DecompressionStream !== "undefined") {
    try {
      const stream = blob.stream().pipeThrough(new DecompressionStream("gzip"));
      return JSON.parse(await new Response(stream).text());
    } catch {
      // Fallback below handles non-gzip backups.
    }
  }
  return JSON.parse(await blob.text());
};

const createMultipartBody = (metadata, mediaBlob) => {
  const boundary = `alphaorc_${Date.now()}`;
  return {
    boundary,
    body: new Blob(
      [
        `--${boundary}\r\n`,
        "Content-Type: application/json; charset=UTF-8\r\n\r\n",
        JSON.stringify(metadata),
        `\r\n--${boundary}\r\n`,
        "Content-Type: application/gzip\r\n\r\n",
        mediaBlob,
        `\r\n--${boundary}--`,
      ],
      { type: `multipart/related; boundary=${boundary}` }
    ),
  };
};

const readDriveJsonFile = async (fileId) => {
  const response = await driveFetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`);
  return await readBackupBlob(await response.blob());
};

const upsertDriveJsonFile = async (folderId, name, data, knownFile = null) => {
  const existing = knownFile || await findNamedFile(folderId, name);
  const blob = await gzipJson(data);
  if (existing) {
    await driveFetch(`https://www.googleapis.com/upload/drive/v3/files/${existing.id}?uploadType=media`, {
      method: "PATCH",
      headers: { "Content-Type": "application/gzip" },
      body: blob,
    });
    return { fileId: existing.id, updated: true };
  }

  const { boundary, body } = createMultipartBody(
    {
      name,
      parents: [folderId],
      mimeType: "application/gzip",
    },
    blob
  );
  const response = await driveFetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id", {
    method: "POST",
    headers: { "Content-Type": `multipart/related; boundary=${boundary}` },
    body,
  });
  const created = await response.json();
  return { fileId: created.id, updated: false };
};

const saveSeparatedSnapshot = async (
  folder,
  data,
  { includeBase = false, projectIds = [], migrateAll = false } = {}
) => {
  const baseFile = await findNamedFile(folder.id, DRIVE_BASE_NAME);
  const firstSeparatedSave = !baseFile;
  const idsToSave = new Set(projectIds || []);
  const projectsToSave = firstSeparatedSave || migrateAll
    ? (data.projetos || [])
    : (data.projetos || []).filter((project) => idsToSave.has(project.id));
  const savedAt = new Date().toISOString();

  for (const project of projectsToSave) {
    await upsertDriveJsonFile(folder.id, projectFileName(project.id), {
      storageVersion: DRIVE_STORAGE_VERSION,
      storage: "google-drive-separated",
      savedAt,
      projectId: project.id,
      projectOrder: (data.projetos || []).findIndex((item) => item.id === project.id),
      projeto: project,
    });
  }

  if (firstSeparatedSave || includeBase || migrateAll) {
    await upsertDriveJsonFile(
      folder.id,
      DRIVE_BASE_NAME,
      {
        storageVersion: DRIVE_STORAGE_VERSION,
        storage: "google-drive-separated",
        savedAt,
        projetoAtivoId: data.projetoAtivoId || "",
        cpus: data.cpus || [],
        precos: data.precos || [],
      },
      baseFile
    );
  }

  return {
    storageVersion: DRIVE_STORAGE_VERSION,
    projetosSalvos: projectsToSave.length,
    baseSalva: firstSeparatedSave || includeBase || migrateAll,
  };
};

export async function saveGoogleDriveSnapshot(data, options = {}) {
  await requestGoogleDriveAccess();
  const folder = await ensureDriveFolder();
  return await saveSeparatedSnapshot(folder, data, options);
}

export async function loadGoogleDriveSnapshot() {
  await requestGoogleDriveAccess();
  const folder = await ensureDriveFolder();
  const baseFile = await findNamedFile(folder.id, DRIVE_BASE_NAME);

  if (baseFile) {
    const [base, projectFiles] = await Promise.all([
      readDriveJsonFile(baseFile.id),
      listProjectFiles(folder.id),
    ]);
    const projectRecords = (
      await Promise.all(
        projectFiles.map(async (file) => {
          try {
            const data = await readDriveJsonFile(file.id);
            return data?.projeto
              ? {
                  project: data.projeto,
                  order: Number.isFinite(Number(data.projectOrder)) ? Number(data.projectOrder) : Number.MAX_SAFE_INTEGER,
                  modifiedTime: file.modifiedTime || "",
                }
              : null;
          } catch (error) {
            console.warn(`Não foi possível carregar ${file.name}:`, error);
            return null;
          }
        })
      )
    ).filter(Boolean);
    const projectsById = new Map();
    projectRecords
      .sort((a, b) => String(b.modifiedTime).localeCompare(String(a.modifiedTime)))
      .forEach((record) => {
        if (!projectsById.has(record.project.id)) projectsById.set(record.project.id, record);
      });
    const projetos = Array.from(projectsById.values())
      .sort((a, b) => a.order - b.order)
      .map((record) => record.project);

    return {
      empty: false,
      storageVersion: DRIVE_STORAGE_VERSION,
      storage: "google-drive-separated",
      cpus: base.cpus || [],
      precos: base.precos || [],
      projetos,
      projetoAtivoId:
        projetos.some((project) => project.id === base.projetoAtivoId)
          ? base.projetoAtivoId
          : projetos[0]?.id || "",
    };
  }

  const legacyFile = await findBackupFile(folder.id);
  if (!legacyFile) return null;

  const legacyData = await readDriveJsonFile(legacyFile.id);
  await saveSeparatedSnapshot(folder, legacyData, {
    includeBase: true,
    migrateAll: true,
    projectIds: (legacyData.projetos || []).map((project) => project.id),
  });

  return {
    ...legacyData,
    empty: false,
    storageVersion: DRIVE_STORAGE_VERSION,
    storage: "google-drive-separated",
  };
}

export async function deleteGoogleDriveProject(projectId) {
  await requestGoogleDriveAccess();
  const folder = await ensureDriveFolder();
  const file = await findNamedFile(folder.id, projectFileName(projectId));
  if (!file) return { deleted: false };

  await driveFetch(`https://www.googleapis.com/drive/v3/files/${file.id}`, {
    method: "DELETE",
  });
  return { deleted: true };
}
