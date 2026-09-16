import { graphFetch } from "./graphClient";

export type TemplateSource =
  | { kind: "group"; groupId: string; notebookName?: string }
  | { kind: "site"; siteUrl: string; notebookName?: string };

export interface CopyTemplateOptions {
  token: string;
  source: TemplateSource;
  sections: { from: string; to: string }[];
  targetGroupId: string;
}

export interface CopyTemplateResult {
  pagesCopied: number;
  sections: { from: string; to: string; pagesCopied: number }[];
}

interface GraphOneNoteSection {
  id: string;
  displayName: string;
}

interface GraphOneNotePage {
  id: string;
  title: string;
}

function normalizeName(value: string): string {
  return value.trim().replace(/\.one$/i, "").toLowerCase();
}

async function resolveSiteId(siteUrl: string, token: string): Promise<string> {
  const url = new URL(siteUrl);
  const sitePath = url.pathname.replace(/\/+$/, "");
  const response = await graphFetch(`/sites/${url.hostname}:${sitePath}?$select=id`, token);
  if (!response.ok) {
    throw new Error(`Failed to resolve SharePoint site ${siteUrl}: ${response.status} ${await response.text()}`);
  }
  const body = (await response.json()) as { id: string };
  return body.id;
}

async function resolveSourceBasePath(source: TemplateSource, token: string): Promise<string> {
  if (source.kind === "group") {
    return `/groups/${source.groupId}`;
  }
  return `/sites/${await resolveSiteId(source.siteUrl, token)}`;
}

async function resolveSectionsPath(
  basePath: string,
  notebookName: string | undefined,
  token: string
): Promise<string> {
  if (!notebookName) {
    return `${basePath}/onenote/sections`;
  }

  const response = await graphFetch(`${basePath}/onenote/notebooks?$select=id,displayName`, token);
  if (!response.ok) {
    throw new Error(`Failed to list notebooks at ${basePath}: ${response.status} ${await response.text()}`);
  }
  const body = (await response.json()) as { value: { id: string; displayName: string }[] };
  const notebook = body.value.find((n) => normalizeName(n.displayName) === normalizeName(notebookName));
  if (!notebook) {
    const available = body.value.map((n) => n.displayName).join(", ") || "none";
    throw new Error(`Notebook "${notebookName}" not found at ${basePath}. Available: ${available}`);
  }
  return `${basePath}/onenote/notebooks/${notebook.id}/sections`;
}

async function listSections(sectionsPath: string, token: string): Promise<GraphOneNoteSection[]> {
  const response = await graphFetch(`${sectionsPath}?$select=id,displayName&$top=100`, token);
  if (!response.ok) {
    throw new Error(`Failed to list sections at ${sectionsPath}: ${response.status} ${await response.text()}`);
  }
  const body = (await response.json()) as { value: GraphOneNoteSection[] };
  return body.value;
}

function findByName(sections: GraphOneNoteSection[], name: string): GraphOneNoteSection | undefined {
  return sections.find((s) => normalizeName(s.displayName) === normalizeName(name));
}

async function getDefaultNotebookSectionsUrl(groupId: string, token: string): Promise<string> {
  const response = await graphFetch(`/groups/${groupId}/onenote/notebooks?$select=id,isDefault`, token);
  if (!response.ok) {
    throw new Error(`Failed to list notebooks for group ${groupId}: ${response.status} ${await response.text()}`);
  }
  const body = (await response.json()) as { value: { id: string; isDefault: boolean }[] };
  const defaultNotebook = body.value.find((n) => n.isDefault) ?? body.value[0];
  if (!defaultNotebook) {
    throw new Error(`Group ${groupId} has no OneNote notebook yet.`);
  }
  return `/groups/${groupId}/onenote/notebooks/${defaultNotebook.id}/sections`;
}

async function createSection(
  groupId: string,
  sectionName: string,
  token: string
): Promise<GraphOneNoteSection> {
  const sectionsUrl = await getDefaultNotebookSectionsUrl(groupId, token);
  const response = await graphFetch(sectionsUrl, token, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ displayName: sectionName }),
  });
  if (!response.ok) {
    throw new Error(`Failed to create section "${sectionName}" for group ${groupId}: ${response.status} ${await response.text()}`);
  }
  return (await response.json()) as GraphOneNoteSection;
}

async function listPages(sectionId: string, token: string): Promise<GraphOneNotePage[]> {
  const response = await graphFetch(`/onenote/sections/${sectionId}/pages?$select=id,title`, token);
  if (!response.ok) {
    throw new Error(`Failed to list pages for section ${sectionId}: ${response.status} ${await response.text()}`);
  }
  const body = (await response.json()) as { value: GraphOneNotePage[] };
  return body.value;
}

async function getPageContent(pageId: string, token: string): Promise<string> {
  const response = await graphFetch(`/onenote/pages/${pageId}/content`, token);
  if (!response.ok) {
    throw new Error(`Failed to read content of page ${pageId}: ${response.status} ${await response.text()}`);
  }
  return response.text();
}

async function createPage(sectionId: string, html: string, token: string): Promise<void> {
  const response = await graphFetch(`/onenote/sections/${sectionId}/pages`, token, {
    method: "POST",
    headers: { "Content-Type": "application/xhtml+xml" },
    body: html,
  });
  if (!response.ok) {
    throw new Error(`Failed to create page in section ${sectionId}: ${response.status} ${await response.text()}`);
  }
}

/**
 * Copies the pages of each template section into the target group's notebook.
 * Target sections are created on demand.
 */
export async function copyTemplateSectionToGroup(options: CopyTemplateOptions): Promise<CopyTemplateResult> {
  const { token, source, sections, targetGroupId } = options;

  const basePath = await resolveSourceBasePath(source, token);
  const templateSectionsPath = await resolveSectionsPath(basePath, source.notebookName, token);
  const templateSections = await listSections(templateSectionsPath, token);

  const targetSectionsPath = await getDefaultNotebookSectionsUrl(targetGroupId, token);
  let targetSections = await listSections(targetSectionsPath, token);

  const result: CopyTemplateResult = { pagesCopied: 0, sections: [] };

  // No explicit selection means: mirror every section of the template notebook.
  const mappings = sections.length
    ? sections
    : templateSections.map((s) => ({ from: s.displayName, to: s.displayName }));

  for (const mapping of mappings) {
    const templateSection = findByName(templateSections, mapping.from);
    if (!templateSection) {
      const available = templateSections.map((s) => s.displayName).join(", ") || "none";
      throw new Error(`Template section "${mapping.from}" not found. Available: ${available}`);
    }

    let targetSection = findByName(targetSections, mapping.to);
    if (!targetSection) {
      targetSection = await createSection(targetGroupId, mapping.to, token);
      targetSections = [...targetSections, targetSection];
    }

    const pages = await listPages(templateSection.id, token);
    // Graph returns newest pages first; copy in reverse to keep the template order.
    for (const page of [...pages].reverse()) {
      const html = await getPageContent(page.id, token);
      await createPage(targetSection.id, html, token);
    }

    result.pagesCopied += pages.length;
    result.sections.push({ from: mapping.from, to: mapping.to, pagesCopied: pages.length });
  }

  return result;
}
