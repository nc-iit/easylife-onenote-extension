import { graphFetch } from "./graphClient";

// The OneNote API rejects app-only tokens (Graph error 40001), so sections are copied as
// the underlying .one files from the SharePoint "Site Assets" library instead.

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
}

interface DriveItem {
  id: string;
  name: string;
  folder?: { childCount?: number };
  file?: { mimeType?: string };
}

const SECTION_EXTENSION = ".one";
const NOTEBOOK_MARKER_EXTENSION = ".onetoc2";

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

/** OneNote notebooks live in the "Site Assets" document library of a SharePoint site. */
async function getSiteAssetsDriveId(siteId: string, token: string): Promise<string> {
  const body = await getJson<{ value: { id: string; name: string }[] }>(
    `/sites/${siteId}/drives?$select=id,name`,
    token,
    `Listing document libraries of site ${siteId}`
  );
  const drive = body.value.find((d) => /site\s*assets/i.test(d.name));
  if (!drive) {
    const available = body.value.map((d) => d.name).join(", ") || "none";
    throw new Error(`No "Site Assets" library found. Available libraries: ${available}`);
  }
  return drive.id;
}

async function listChildren(driveId: string, itemPath: string, token: string): Promise<DriveItem[]> {
  const body = await getJson<{ value: DriveItem[] }>(
    `/drives/${driveId}/${itemPath}?$select=id,name,folder,file&$top=200`,
    token,
    `Listing ${itemPath} of drive ${driveId}`
  );
  return body.value;
}

/** Finds the notebook folder by name, or the only folder that actually contains a notebook. */
async function findNotebookFolder(
  driveId: string,
  notebookName: string | undefined,
  token: string
): Promise<DriveItem> {
  const rootItems = await listChildren(driveId, "root/children", token);
  const folders = rootItems.filter((item) => item.folder);

  if (notebookName) {
    const match = folders.find((f) => normalizeName(f.name) === normalizeName(notebookName));
    if (!match) {
      const available = folders.map((f) => f.name).join(", ") || "none";
      throw new Error(`Notebook "${notebookName}" not found. Available folders: ${available}`);
    }
    return match;
  }

  for (const folder of folders) {
    const children = await listChildren(driveId, `items/${folder.id}/children`, token);
    if (children.some((c) => c.name.toLowerCase().endsWith(NOTEBOOK_MARKER_EXTENSION))) {
      return folder;
    }
  }

  const available = folders.map((f) => f.name).join(", ") || "none";
  throw new Error(`No OneNote notebook found in Site Assets. Available folders: ${available}`);
}

async function copySectionFile(
  sourceDriveId: string,
  sourceItemId: string,
  targetDriveId: string,
  targetFolderId: string,
  newName: string,
  token: string
): Promise<void> {
  const response = await graphFetch(`/drives/${sourceDriveId}/items/${sourceItemId}/copy`, token, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      parentReference: { driveId: targetDriveId, id: targetFolderId },
      name: newName.toLowerCase().endsWith(SECTION_EXTENSION) ? newName : `${newName}${SECTION_EXTENSION}`,
    }),
  });

  // Graph answers 202 Accepted and finishes the copy asynchronously.
  if (!response.ok && response.status !== 202) {
    throw new Error(`Copying section "${newName}" failed: ${response.status} ${await response.text()}`);
  }
}

/** Copies template sections (.one files) into the notebook of the newly provisioned group. */
export async function copyTemplateSectionsToGroup(options: CopyTemplateOptions): Promise<CopyTemplateResult> {
  const { token, source, sections, targetGroupId } = options;

  const sourceSiteId = await resolveSiteId(source, token);
  const sourceDriveId = await getSiteAssetsDriveId(sourceSiteId, token);
  const sourceNotebook = await findNotebookFolder(sourceDriveId, source.notebookName, token);
  const sourceSections = (await listChildren(sourceDriveId, `items/${sourceNotebook.id}/children`, token)).filter(
    (item) => item.file && item.name.toLowerCase().endsWith(SECTION_EXTENSION)
  );

  const targetSiteId = await resolveSiteIdFromGroup(targetGroupId, token);
  const targetDriveId = await getSiteAssetsDriveId(targetSiteId, token);
  const targetNotebook = await findNotebookFolder(targetDriveId, undefined, token);

  const mappings = sections.length
    ? sections
    : sourceSections.map((s) => ({ from: s.name, to: s.name }));

  const copied: { from: string; to: string }[] = [];

  for (const mapping of mappings) {
    const sourceSection = sourceSections.find((s) => normalizeName(s.name) === normalizeName(mapping.from));
    if (!sourceSection) {
      const available = sourceSections.map((s) => normalizeName(s.name)).join(", ") || "none";
      throw new Error(`Template section "${mapping.from}" not found. Available sections: ${available}`);
    }

    await copySectionFile(sourceDriveId, sourceSection.id, targetDriveId, targetNotebook.id, mapping.to, token);
    copied.push({ from: normalizeName(sourceSection.name), to: normalizeName(mapping.to) });
  }

  return {
    sectionsCopied: copied,
    templateNotebook: sourceNotebook.name,
    targetNotebook: targetNotebook.name,
  };
}
