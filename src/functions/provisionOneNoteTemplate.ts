import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { getGraphAccessToken } from "../services/graphClient";
import { copyTemplateSectionToGroup, TemplateSource } from "../services/oneNoteTemplateCopier";

/** Extracts the newly created group's id from the EasyLife 365 webhook payload. */
function extractGroupId(body: unknown): string | undefined {
  if (!body || typeof body !== "object") {
    return undefined;
  }
  const record = body as Record<string, unknown>;
  const group = record.group as Record<string, unknown> | undefined;
  const candidate = group?.id ?? record.groupId ?? record.resourceId ?? record.id;
  return typeof candidate === "string" ? candidate : undefined;
}

function splitNames(value: string | null | undefined): string[] {
  return (value ?? "")
    .split(",")
    .map((name) => name.trim())
    .filter((name) => name.length > 0);
}

function buildSectionMappings(from: string[], to: string[]): { from: string; to: string }[] {
  return from.map((name, index) => ({ from: name, to: to[index] ?? name }));
}

export async function provisionOneNoteTemplate(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const body = await request.json().catch(() => undefined);
  context.log("EasyLife webhook payload", JSON.stringify(body));

  const targetGroupId = extractGroupId(body);
  if (!targetGroupId) {
    context.warn("No group id found in webhook payload.");
    return { status: 400, jsonBody: { error: "Could not determine groupId from webhook payload." } };
  }

  // Everything below can be set per EasyLife automation step via the webhook URL query string.
  const templateSiteUrl = request.query.get("templateSiteUrl") ?? process.env.DEFAULT_TEMPLATE_SITE_URL;
  const templateGroupId = request.query.get("templateGroupId") ?? process.env.DEFAULT_TEMPLATE_GROUP_ID;
  const notebookName =
    request.query.get("templateNotebookName") ?? process.env.DEFAULT_TEMPLATE_NOTEBOOK_NAME ?? undefined;

  const templateSectionNames = splitNames(
    request.query.get("templateSectionName") ?? process.env.DEFAULT_TEMPLATE_SECTION_NAMES
  );
  const targetSectionNames = splitNames(
    request.query.get("targetSectionName") ?? process.env.DEFAULT_TARGET_SECTION_NAMES
  );

  let source: TemplateSource;
  if (templateSiteUrl) {
    source = { kind: "site", siteUrl: templateSiteUrl, notebookName };
  } else if (templateGroupId) {
    source = { kind: "group", groupId: templateGroupId, notebookName };
  } else {
    context.warn(
      "No template source configured. Set templateSiteUrl/templateGroupId in the webhook URL or DEFAULT_TEMPLATE_SITE_URL/DEFAULT_TEMPLATE_GROUP_ID in the app settings."
    );
    return {
      status: 400,
      jsonBody: { error: "No template source configured. Provide templateSiteUrl or templateGroupId." },
    };
  }

  context.log(
    `Copying into group ${targetGroupId} from ${source.kind} source`,
    JSON.stringify({ notebookName, templateSectionNames, targetSectionNames })
  );

  try {
    const token = await getGraphAccessToken();
    const result = await copyTemplateSectionToGroup({
      token,
      source,
      sections: buildSectionMappings(templateSectionNames, targetSectionNames),
      targetGroupId,
    });

    context.log(`Copied ${result.pagesCopied} page(s) into group ${targetGroupId}.`, JSON.stringify(result.sections));
    return { status: 200, jsonBody: { status: "ok", ...result } };
  } catch (err) {
    context.error("Failed to copy OneNote template", err);
    return { status: 500, jsonBody: { error: (err as Error).message } };
  }
}

app.http("provisionOneNoteTemplate", {
  methods: ["POST"],
  authLevel: "function",
  route: "onenote-template",
  handler: provisionOneNoteTemplate,
});
