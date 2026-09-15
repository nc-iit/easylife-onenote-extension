import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { getGraphAccessToken } from "../services/graphClient";
import { copyTemplateSectionToGroup } from "../services/oneNoteTemplateCopier";

/**
 * Extracts the newly created group's id from the EasyLife 365 webhook payload.
 * EasyLife's exact field name may vary by automation type - adjust once confirmed
 * against a real payload (check the raw body logged below on first test run).
 */
function extractGroupId(body: unknown): string | undefined {
  if (!body || typeof body !== "object") {
    return undefined;
  }
  const record = body as Record<string, unknown>;
  const candidate =
    record.groupId ?? record.id ?? record.resourceId ?? (record.group as Record<string, unknown> | undefined)?.id;
  return typeof candidate === "string" ? candidate : undefined;
}

export async function provisionOneNoteTemplate(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const body = await request.json().catch(() => undefined);
  context.log("EasyLife webhook payload", JSON.stringify(body));

  const targetGroupId = extractGroupId(body);
  if (!targetGroupId) {
    return { status: 400, jsonBody: { error: "Could not determine groupId from webhook payload." } };
  }

  // Template source/target section can be overridden per automation step via the webhook URL's
  // query string, e.g. ?templateGroupId=...&templateSectionName=Traktandenliste&targetSectionName=Traktandenliste
  const templateGroupId = request.query.get("templateGroupId") ?? process.env.DEFAULT_TEMPLATE_GROUP_ID;
  const templateSectionName =
    request.query.get("templateSectionName") ?? process.env.DEFAULT_TEMPLATE_SECTION_NAME ?? "Traktandenliste";
  const targetSectionName =
    request.query.get("targetSectionName") ?? process.env.DEFAULT_TARGET_SECTION_NAME ?? templateSectionName;

  if (!templateGroupId) {
    return { status: 400, jsonBody: { error: "No templateGroupId configured (query string or DEFAULT_TEMPLATE_GROUP_ID)." } };
  }

  try {
    const token = await getGraphAccessToken();
    const result = await copyTemplateSectionToGroup({
      token,
      templateGroupId,
      templateSectionName,
      targetGroupId,
      targetSectionName,
    });

    context.log(`Copied ${result.pagesCopied} page(s) into group ${targetGroupId}, section "${targetSectionName}".`);
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
