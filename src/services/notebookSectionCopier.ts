import { graphFetch, sleep } from "./graphClient";

// The OneNote API rejects app-only tokens (Graph error 40001), so sections are copied as
// the underlying .one files through the SharePoint Drive API instead.

export type TemplateSource =
  | { kind: "site"; siteUrl: string; notebookName?: string }
  | { kind: "group"; groupId: string; notebookName?: string };

export interface CopyTemplateOptions {
  token: string;
  source: TemplateSource;
  sections: { from: string; to: string }[];
  targetGroupId: string;
}

export interface CopyTemplateResult {
  sectionsCopied: { from: string; to: string }[];
  templateNotebook: string;
  targetNotebook: string;
  filesInTargetNotebook: string[];
}

interface DriveItem {
  id: string;
  name: string;
  folder?: { childCount?: number };
  file?: { mimeType?: string };
}

interface NotebookLocation {
  driveId: string;
  driveName: string;
  folderId: string;
  folderName: string;
}

const SECTION_EXTENSION = ".one";
const NOTEBOOK_MARKER_EXTENSION = ".onetoc2";

// EasyLife fires the webhook before SharePoint has finished provisioning the group notebook.
const PROVISIONING_ATTEMPTS = 6;
const PROVISIONING_DELAY_MS = 5000;

function normalizeName(value: string): string {
  return value.trim().replace(/\.one$/i, "").toLowerCase();
}

async function getJson<T>(path: string, token: string, what: string): Promise<T> {
  const response = await graphFetch(path, token);
  if (!response.ok) {
    throw new Error(`${what} failed: ${response.status} ${await response.text()}`);
  }
  return (await response.json()) as T;
}

async function resolveSiteIdFromUrl(siteUrl: string, token: string): Promise<string> {
  const url = new URL(siteUrl);
  const sitePath = url.pathname.replace(/\/+$/, "");
  const site = await getJson<{ id: string }>(
    `/sites/${url.hostname}:${sitePath}?$select=id`,
    token,
    `Resolving SharePoint site ${siteUrl}`
  );
  return site.id;
}

async function resolveSiteIdFromGroup(groupId: string, token: string): Promise<string> {
  const site = await getJson<{ id: string }>(
    `/groups/${groupId}/sites/root?$select=id`,
    token,
    `Resolving site of group ${groupId}`
  );
  return site.id;
}

async function resolveSiteId(source: TemplateSource, token: string): Promise<string> {
  return source.kind === "site"
    ? resolveSiteIdFromUrl(source.siteUrl, token)
    : resolveSiteIdFromGroup(source.groupId, token);
}

async function listDrives(siteId: string, token: string): Promise<{ id: string; name: string }[]> {
  const drives = new Map<string, { id: string; name: string }>();

  const direct = await getJson<{ value: { id: string; name: string }[] }>(
    `/sites/${siteId}/drives?$select=id,name`,
    token,
    `Listing document libraries of site ${siteId}`
  );
  for (const drive of direct.value) {
    drives.set(drive.id, drive);
  }

  // /drives omits some libraries (notably "Site Assets"), so also resolve them through /lists.
  const lists = await getJson<{ value: { id: string; displayName: string; list?: { template?: string } }[] }>(
    `/sites/${siteId}/lists?$select=id,displayName,list&$top=200`,
    token,
    `Listing lists of site ${siteId}`
  );

  for (const list of lists.value) {
    try {
      const drive = await getJson<{ id: string; name?: string }>(
        `/sites/${siteId}/lists/${list.id}/drive?$select=id,name`,
        token,
        `Resolving drive of list ${list.displayName}`
      );
      if (!drives.has(drive.id)) {
        drives.set(drive.id, { id: drive.id, name: drive.name ?? list.displayName });
      }
    } catch {
      // Lists without an associated drive are not relevant here.
    }
  }

  // Graph hides some system libraries from /drives and /lists; SiteAssets holds the notebooks.
  for (const knownList of ["SiteAssets", "Site Assets", "Shared Documents", "Documents"]) {
    try {
      const drive = await getJson<{ id: string; name?: string }>(
        `/sites/${siteId}/lists/${encodeURIComponent(knownList)}/drive?$select=id,name`,
        token,
        `Resolving drive of list ${knownList}`
      );
      if (!drives.has(drive.id)) {
        drives.set(drive.id, { id: drive.id, name: drive.name ?? knownList });
      }
    } catch {
      // Library does not exist under this name.
    }
  }

  return [...drives.values()];
}

async function listChildren(driveId: string, itemPath: string, token: string): Promise<DriveItem[]> {
  const body = await getJson<{ value: DriveItem[] }>(
    `/drives/${driveId}/${itemPath}?$select=id,name,folder,file&$top=200`,
    token,
    `Listing ${itemPath} of drive ${driveId}`
  );
  return body.value;
}

function isNotebookFolder(children: DriveItem[]): boolean {
  return children.some((c) => c.name.toLowerCase().endsWith(NOTEBOOK_MARKER_EXTENSION));
}

const MAX_FOLDER_DEPTH = 3;

async function findNotebookInDrive(
  drive: { id: string; name: string },
  notebookName: string | undefined,
  token: string,
  inspected: string[]
): Promise<NotebookLocation | undefined> {
  async function walk(folders: DriveItem[], depth: number): Promise<NotebookLocation | undefined> {
    for (const folder of folders) {
      const children = await listChildren(drive.id, `items/${folder.id}/children`, token);
      inspected.push(`${drive.name}/${folder.name}`);

      const nameMatches = !notebookName || normalizeName(folder.name) === normalizeName(notebookName);
      if (nameMatches && isNotebookFolder(children)) {
        return { driveId: drive.id, driveName: drive.name, folderId: folder.id, folderName: folder.name };
      }

      if (depth < MAX_FOLDER_DEPTH) {
        const nested = await walk(children.filter((c) => c.folder), depth + 1);
        if (nested) {
          return nested;
        }
      }
    }
    return undefined;
  }

  const rootFolders = (await listChildren(drive.id, "root/children", token)).filter((item) => item.folder);
  return walk(rootFolders, 1);
}

/** Fallback when the notebook sits deeper than the folder walk reaches. */
async function searchNotebookInDrive(
  drive: { id: string; name: string },
  notebookName: string,
  token: string
): Promise<NotebookLocation | undefined> {
  const response = await graphFetch(
    `/drives/${drive.id}/root/search(q='${encodeURIComponent(notebookName)}')?$select=id,name,folder`,
    token
  );
  if (!response.ok) {
    return undefined;
  }

  const body = (await response.json()) as { value: DriveItem[] };
  for (const item of body.value.filter((i) => i.folder)) {
    const children = await listChildren(drive.id, `items/${item.id}/children`, token);
    if (isNotebookFolder(children)) {
      return { driveId: drive.id, driveName: drive.name, folderId: item.id, folderName: item.name };
    }
  }
  return undefined;
}

/** Notebooks may live in any document library of the site, not only in "Site Assets". */
async function findNotebook(
  siteId: string,
  notebookName: string | undefined,
  token: string
): Promise<NotebookLocation> {
  const drives = await listDrives(siteId, token);
  // "Site Assets" holds notebooks in most tenants, so check it first.
  const ordered = [...drives].sort((a, b) => Number(/site\s*assets/i.test(b.name)) - Number(/site\s*assets/i.test(a.name)));
  const inspected: string[] = [];

  for (const drive of ordered) {
    const found = await findNotebookInDrive(drive, notebookName, token, inspected);
    if (found) {
      return found;
    }
  }

  if (notebookName) {
    for (const drive of ordered) {
      const found = await searchNotebookInDrive(drive, notebookName, token);
      if (found) {
        return found;
      }
    }
  }

  const libraries = drives.map((d) => d.name).join(", ") || "none";
  const folders = inspected.join(", ") || "none";
  throw new Error(
    `Notebook ${notebookName ? `"${notebookName}" ` : ""}not found in site ${siteId}. ` +
      `Libraries: ${libraries}. Folders inspected: ${folders}`
  );
}

/** Retries while the group's site or notebook is still being provisioned. */
async function waitForProvisioned<T>(what: string, resolve: () => Promise<T>): Promise<T> {
  let lastError: unknown;

  for (let attempt = 0; attempt < PROVISIONING_ATTEMPTS; attempt++) {
    try {
      return await resolve();
    } catch (err) {
      lastError = err;
      if (attempt < PROVISIONING_ATTEMPTS - 1) {
        await sleep(PROVISIONING_DELAY_MS);
      }
    }
  }

  throw new Error(
    `${what} is not available after ${PROVISIONING_ATTEMPTS} attempts: ${(lastError as Error).message}`
  );
}

/** Graph reports copy progress on an unauthenticated monitor URL. */
async function waitForCopyToFinish(monitorUrl: string): Promise<void> {
  for (let attempt = 0; attempt < 10; attempt++) {
    const response = await fetch(monitorUrl);
    if (!response.ok) {
      return;
    }

    const status = (await response.json()) as { status?: string; error?: unknown };
    if (status.status === "completed") {
      return;
    }
    if (status.status === "failed") {
      throw new Error(`Copy failed: ${JSON.stringify(status.error)}`);
    }
    await sleep(2000);
  }
}

async function copySectionFile(
  source: NotebookLocation,
  sectionItemId: string,
  target: NotebookLocation,
  newName: string,
  token: string
): Promise<void> {
  const fileName = newName.toLowerCase().endsWith(SECTION_EXTENSION) ? newName : `${newName}${SECTION_EXTENSION}`;

  async function requestCopy(): Promise<Response> {
    return graphFetch(`/drives/${source.driveId}/items/${sectionItemId}/copy`, token, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        parentReference: { driveId: target.driveId, id: target.folderId },
        name: fileName,
        // EasyLife pre-creates empty sections with the same name, so overwrite them.
        "@microsoft.graph.conflictBehavior": "replace",
      }),
    });
  }

  let response = await requestCopy();

  // Not every drive honours conflictBehavior on copy, so remove the placeholder and retry.
  if (response.status === 409) {
    const existing = (await listChildren(target.driveId, `items/${target.folderId}/children`, token)).find(
      (item) => item.name.toLowerCase() === fileName.toLowerCase()
    );
    if (existing) {
      await graphFetch(`/drives/${target.driveId}/items/${existing.id}`, token, { method: "DELETE" });
      response = await requestCopy();
    }
  }

  // Graph answers 202 Accepted and completes the copy asynchronously.
  if (!response.ok && response.status !== 202) {
    throw new Error(`Copying section "${newName}" failed: ${response.status} ${await response.text()}`);
  }

  const monitorUrl = response.headers.get("Location");
  if (monitorUrl) {
    await waitForCopyToFinish(monitorUrl);
  }
}

/** Copies template sections (.one files) into the notebook of the newly provisioned group. */
export async function copyTemplateSectionsToGroup(options: CopyTemplateOptions): Promise<CopyTemplateResult> {
  const { token, source, sections, targetGroupId } = options;

  const sourceSiteId = await resolveSiteId(source, token);
  const sourceNotebook = await findNotebook(sourceSiteId, source.notebookName, token);
  const sourceSections = (await listChildren(sourceNotebook.driveId, `items/${sourceNotebook.folderId}/children`, token))
    .filter((item) => item.file && item.name.toLowerCase().endsWith(SECTION_EXTENSION));

  const targetSiteId = await waitForProvisioned(`Site of group ${targetGroupId}`, () =>
    resolveSiteIdFromGroup(targetGroupId, token)
  );
  const targetNotebook = await waitForProvisioned(`Notebook of group ${targetGroupId}`, () =>
    findNotebook(targetSiteId, undefined, token)
  );

  const mappings = sections.length ? sections : sourceSections.map((s) => ({ from: s.name, to: s.name }));
  const copied: { from: string; to: string }[] = [];

  for (const mapping of mappings) {
    const sourceSection = sourceSections.find((s) => normalizeName(s.name) === normalizeName(mapping.from));
    if (!sourceSection) {
      const available = sourceSections.map((s) => normalizeName(s.name)).join(", ") || "none";
      throw new Error(`Template section "${mapping.from}" not found. Available sections: ${available}`);
    }

    await copySectionFile(sourceNotebook, sourceSection.id, targetNotebook, mapping.to, token);
    copied.push({ from: normalizeName(sourceSection.name), to: normalizeName(mapping.to) });
  }

  const filesInTargetNotebook = (
    await listChildren(targetNotebook.driveId, `items/${targetNotebook.folderId}/children`, token)
  ).map((item) => item.name);

  return {
    sectionsCopied: copied,
    templateNotebook: `${sourceNotebook.driveName}/${sourceNotebook.folderName}`,
    targetNotebook: `${targetNotebook.driveName}/${targetNotebook.folderName}`,
    filesInTargetNotebook,
  };
}
