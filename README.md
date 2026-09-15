# EasyLife 365 OneNote Template Function

Azure Function in Node.js/TypeScript, die nach der Bereitstellung einer EasyLife 365 Team-/Gruppen-Automation eine OneNote-Vorlage in das Notizbuch der neuen Gruppe kopiert.

Beispiel: Eine Section namens `Traktandenliste` mit allen darin enthaltenen Seiten wird aus einer Vorlage in die neue Gruppe kopiert.

**Autoren:** Andy Bui, uniQconsulting ag

## Funktionsweise

1. EasyLife 365 erstellt eine neue Microsoft 365-Gruppe beziehungsweise ein Team.
2. EasyLife ruft den HTTP-Webhook dieser Function auf.
3. Azure Functions prüft den Function Key.
4. Die Function liest die Gruppen-ID aus dem EasyLife-Payload.
5. Die Function sucht die gewünschte Vorlage-Section im OneNote-Notizbuch der Vorlagen-Gruppe.
6. Die Ziel-Section wird im Standard-Notizbuch der neuen Gruppe angelegt, falls sie noch nicht existiert.
7. Jede Vorlagenseite wird inklusive ihres XHTML-Inhalts in die Ziel-Section kopiert.

## Voraussetzungen

- Azure-Abonnement
- Azure Function App mit Node.js und dem Node.js v4 Programming Model
- Microsoft Entra App-Registrierung für den Zugriff der Function auf Microsoft Graph
- EasyLife 365 mit einer Automation für Team-/Gruppenbereitstellung
- Node.js LTS und Azure Functions Core Tools für lokale Entwicklung
- Ein OneNote-Notizbuch mit der gewünschten Vorlage

## Lokale Entwicklung

```powershell
npm install
npm run build
```

Für die lokale Ausführung werden Azure Functions Core Tools und gegebenenfalls Azurite für `AzureWebJobsStorage` benötigt:

```powershell
func start
```

Die lokale Datei `local.settings.json` ist absichtlich in `.gitignore` eingetragen. Sie darf keine Secrets in Git enthalten.

## Microsoft Graph App-Registrierung

Die Function verwendet den Client-Credentials-Flow und ruft Microsoft Graph mit einer eigenen App-Identität auf. Die bestehende EasyLife-App-Registrierung ist für diese Webhook-Variante nicht erforderlich.

In der App-Registrierung unter **API permissions** sind als **Application permissions** mindestens die für den verwendeten Graph-Zugriff erforderlichen OneNote-/Gruppenberechtigungen zu konfigurieren. Für dieses Projekt sind vorgesehen:

- `Notes.ReadWrite.All`
- `Group.ReadWrite.All`

Danach muss ein Administrator **Grant admin consent** für den Tenant erteilen.

Unter **Certificates & secrets** wird ein Client Secret erstellt. Der Secret-Wert wird nur einmal angezeigt und muss sicher gespeichert werden.

> Für den produktiven Betrieb sollte statt eines langlebigen Client Secrets nach Möglichkeit eine sicher verwaltete Identität beziehungsweise ein Zertifikat verwendet werden. Die konkrete Möglichkeit hängt vom gewählten Azure-Hosting und der Tenant-Konfiguration ab.

## Azure Function App konfigurieren

Unter **Function App → Settings → Environment variables** beziehungsweise **Configuration → Application settings** müssen folgende Einstellungen hinterlegt werden:

| Einstellung | Wert |
|---|---|
| `FUNCTIONS_WORKER_RUNTIME` | `node` |
| `GRAPH_TENANT_ID` | Microsoft-Entra-Tenant-ID des Microsoft-365-Tenants |
| `GRAPH_CLIENT_ID` | Application (client) ID der Graph-App-Registrierung |
| `GRAPH_CLIENT_SECRET` | Client-Secret-Wert der Graph-App-Registrierung |
| `DEFAULT_TEMPLATE_GROUP_ID` | Gruppen-ID der Gruppe, in deren OneNote die Vorlage liegt |
| `DEFAULT_TEMPLATE_SECTION_NAME` | Standardmässig `Traktandenliste` |
| `DEFAULT_TARGET_SECTION_NAME` | Standardmässig `Traktandenliste` |

Nach Änderungen an Application Settings die Function App neu starten, falls Azure dies nicht automatisch erledigt.

## Webhook in EasyLife 365

In der EasyLife-Automation unter **OneNote**:

1. **Notify via webhook** aktivieren.
2. Bei **Authentication** die Option **Code authentication** wählen.
3. Die URL der Function mit Function Key eintragen:

```text
https://<function-app-name>.azurewebsites.net/api/onenote-template?code=<function-key>
```

Den Function Key erhält man unter:

**Function App → Functions → provisionOneNoteTemplate → Function keys**

Der Key kann alternativ als HTTP-Header `x-functions-key` übermittelt werden. Der Key ist ein Secret und darf nicht in öffentlichen Dokumenten, Screenshots oder Git committed werden.

## Dynamische Vorlagen

Die Vorlage kann pro EasyLife-Webhook über Query-Parameter ausgewählt werden. Dadurch können mehrere Automationen unterschiedliche Vorlagen verwenden:

```text
https://<function-app-name>.azurewebsites.net/api/onenote-template?code=<function-key>&templateGroupId=<template-group-id>&templateSectionName=Traktandenliste&targetSectionName=Traktandenliste
```

Unterstützte Parameter:

- `templateGroupId`: Gruppen-ID der Vorlagen-Gruppe
- `templateSectionName`: Name der Vorlagen-Section
- `targetSectionName`: Name der Section in der neuen Gruppe

Wenn Parameter fehlen, verwendet die Function die Werte aus den Application Settings.

## EasyLife-Payload prüfen

Die Function erwartet die neue Gruppen-ID in einem der folgenden Felder:

- `groupId`
- `id`
- `resourceId`
- `group.id`

Der rohe Payload wird beim ersten Testaufruf im Application-Insights- beziehungsweise Function-Log protokolliert. Falls EasyLife ein anderes Feld verwendet, muss `extractGroupId()` in [src/functions/provisionOneNoteTemplate.ts](src/functions/provisionOneNoteTemplate.ts) angepasst werden.

Für einen erfolgreichen Aufruf antwortet die Function beispielsweise:

```json
{
  "status": "ok",
  "pagesCopied": 3
}
```

## Testablauf

1. Prüfen, dass das Vorlagen-Notizbuch existiert und die gewünschte Section exakt benannt ist.
2. Graph-App-Registrierung und Admin Consent prüfen.
3. Application Settings der Function App hinterlegen.
4. Function-Key-URL in EasyLife eintragen.
5. Eine Test-Gruppe über EasyLife bereitstellen.
6. Im Function-Log prüfen, ob die Gruppen-ID erkannt wurde.
7. Das OneNote-Notizbuch der neuen Gruppe öffnen und die kopierte Section sowie ihre Seiten prüfen.

## GitHub-Deployment

Dieses Repository ist für die Bereitstellung aus GitHub vorgesehen. Der Produktionsbranch ist `main`.

Nach der Konfiguration von **Deployment Center** der Azure Function App mit diesem Repository wird bei Pushes auf `main` automatisch deployt:

```powershell
git add .
git commit -m "Describe the change"
git push origin main
```

Die Secrets und Application Settings werden nicht aus `local.settings.json` deployt. Sie müssen in Azure separat gepflegt werden.

### Fehler 401 beim Deploy mit `Azure/functions-action`

Wenn der Build erfolgreich ist, der Deploy-Schritt aber bei `ValidateAzureResource` mit einem Fehler wie dem folgenden abbricht, betrifft das normalerweise die SCM-/Kudu-Anmeldung:

```text
Failed to fetch Kudu App Settings
https://<function-app>.scm.azurewebsites.net/api/settings
Unauthorized (CODE: 401)
```

Mögliche Ursachen sind ein veraltetes Publish Profile oder deaktivierte SCM-Basic-Authentication. In diesem Fall hilft ein neu heruntergeladenes Publish Profile unter **Overview → Download publish profile** und die Option **SCM Basic Auth Publishing Credentials** unter **Configuration → General settings → Platform settings**.

### Fehler 405 und 404 bei Kudu und ZipDeploy

Zeigt das Log dagegen folgendes Muster, liegt es nicht an den Zugangsdaten:

```text
updateAppSettingViaKudu
Response with status code 405
App setting SCM_DO_BUILD_DURING_DEPLOYMENT has not been propagated to Kudu container yet
...
zipDeploy
Not Found (CODE: 404)
```

Diese Function App läuft dann auf **Flex Consumption**. Dort existieren die Kudu-Endpunkte `/api/settings` und `/api/zipdeploy` nicht, und Publish Profiles beziehungsweise Basic Auth werden nicht unterstützt. Das Deployment muss über Microsoft Entra ID (OIDC) und die ARM-Deployment-API erfolgen.

Der Workflow in diesem Repository verwendet deshalb `azure/login` mit folgenden GitHub Secrets:

| Secret | Inhalt |
|---|---|
| `AZURE_CLIENT_ID` | Client ID der Deployment-App-Registrierung beziehungsweise User-assigned Managed Identity |
| `AZURE_TENANT_ID` | Tenant ID |
| `AZURE_SUBSCRIPTION_ID` | Subscription ID der Function App |

Die Identität benötigt auf der Function App die Rolle **Contributor** oder **Website Contributor** sowie eine Federated Credential für dieses Repository und den Branch `main`.

Am einfachsten wird das im Azure Portal erzeugt: **Function App → Deployment Center → Source: GitHub → Authentication: User-assigned identity**. Azure legt Identität, Rollenzuweisung, Federated Credential und die Secrets automatisch an.

Publish Profile und Function Key sind Geheimnisse. Sie dürfen nicht in Quellcode, README oder Logs gelangen. Ein versehentlich offengelegtes Publish Profile muss in Azure neu generiert werden.

## Projektstruktur

```text
src/
  functions/
    provisionOneNoteTemplate.ts    HTTP-Trigger und EasyLife-Payload
  services/
    graphClient.ts                 Microsoft-Graph-Authentifizierung und HTTP-Aufrufe
    oneNoteTemplateCopier.ts       Lesen und Kopieren der OneNote-Seiten
  index.ts                         Registrierung der Functions
host.json                          Azure-Functions-Laufzeitkonfiguration
package.json                       Abhängigkeiten und Scripts
local.settings.json                lokale Einstellungen, nicht versionieren
```

## Hinweise und Grenzen

- Die Vorlage wird aktuell seitenweise kopiert. Anhänge oder bestimmte eingebettete Inhalte können je nach OneNote-/Graph-XHTML-Repräsentation spezielle Behandlung benötigen.
- Die Seiten werden in der Reihenfolge der Graph-Antwort kopiert; eine garantierte Reihenfolge sollte bei Bedarf separat implementiert und getestet werden.
- Wird die Ziel-Section bereits gefunden, werden die Seiten zusätzlich eingefügt. Für wiederholte Webhook-Aufrufe kann dadurch Duplikation entstehen.
- Die Function wartet synchron auf das Kopieren aller Seiten. Bei sehr grossen Notizbüchern oder vielen Seiten sollte der Prozess auf eine Queue beziehungsweise Durable Function umgestellt werden.

## Lizenz und Eigentum

Dieses Projekt wurde für uniQconsulting ag erstellt. Lizenzierung und Weiterverwendung sind mit uniQconsulting ag zu klären.
