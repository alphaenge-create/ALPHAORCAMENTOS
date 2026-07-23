const DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.file";
const DRIVE_FOLDER_NAME = "AlphaOrcamentos";
const DRIVE_BACKUP_NAME = "alphaorc-backup.json.gz";
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

export async function saveGoogleDriveSnapshot(data) {
  await requestGoogleDriveAccess();
  const folder = await ensureDriveFolder();
  const existing = await findBackupFile(folder.id);
  const blob = await gzipJson({
    ...data,
    savedAt: new Date().toISOString(),
    storage: "google-drive",
  });

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
      name: DRIVE_BACKUP_NAME,
      parents: [folder.id],
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
}

export async function loadGoogleDriveSnapshot() {
  await requestGoogleDriveAccess();
  const folder = await ensureDriveFolder();
  const file = await findBackupFile(folder.id);
  if (!file) return null;

  const response = await driveFetch(`https://www.googleapis.com/drive/v3/files/${file.id}?alt=media`);
  const blob = await response.blob();
  return await readBackupBlob(blob);
}
