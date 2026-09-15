import { graphFetch } from "./graphClient";

export interface CopyTemplateOptions {
  token: string;
  templateGroupId: string;
  templateSectionName: string;
  targetGroupId: string;
  targetSectionName: string;
}

interface GraphOneNoteSection {
  id: string;
  displayName: string;
}

interface GraphOneNotePage {
  id: string;
  title: string;
}

/** Finds a OneNote section by display name within a group's notebook(s). Creates target sections on demand. */
async function findSection(
  groupId: string,
  sectionName: string,
  token: string
): Promise<GraphOneNoteSection | undefined> {
  const response = await graphFetch(
    `/groups/${groupId}/onenote/sections?$filter=displayName eq '${sectionName.replace(/'/g, "''")}'`,
    token
  );
  if (!response.ok) {
    throw new Error(`Failed to list sections for group ${groupId}: ${response.status} ${await response.text()}`);
  }
  const body = (await response.json()) as { value: GraphOneNoteSection[] };
  return body.value[0];
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
 * Copies every page of the template section into the target group's section (created if missing),
 * so the target notebook ends up with the same content (e.g. a "Traktandenliste").
 */
export async function copyTemplateSectionToGroup(options: CopyTemplateOptions): Promise<{ pagesCopied: number }> {
  const { token, templateGroupId, templateSectionName, targetGroupId, targetSectionName } = options;

  const templateSection = await findSection(templateGroupId, templateSectionName, token);
  if (!templateSection) {
    throw new Error(`Template section "${templateSectionName}" not found in group ${templateGroupId}.`);
  }

  const targetSection =
    (await findSection(targetGroupId, targetSectionName, token)) ??
    (await createSection(targetGroupId, targetSectionName, token));

  const pages = await listPages(templateSection.id, token);

  for (const page of pages) {
    const html = await getPageContent(page.id, token);
    await createPage(targetSection.id, html, token);
  }

  return { pagesCopied: pages.length };
}
