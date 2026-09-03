# Enterprise Deployment Runbook: Office Agents

This runbook guides enterprise IT administrators and security teams through deploying, configuring, and maintaining Office Agents (Excel, PowerPoint, Word Add-ins) in an internal enterprise Microsoft 365 environment.

---

## 1. Architecture Overview

Office Agents is designed for single-organization internal enterprise deployment:
- **Client Execution**: Office Web Add-in running inside Microsoft Office desktop and web taskpanes.
- **Authentication**: Native Microsoft Office Single Sign-On (Office SSO / Entra ID) using `OfficeRuntime.auth.getAccessToken`.
- **Zero-BYOK Storage**: In `enterprise` mode, no API keys or third-party OAuth tokens are stored in the browser. Credentials route strictly through corporate gateways using Entra ID bearer tokens.
- **Corporate LLM Gateway**: Requests route to an enterprise OpenAI- or Anthropic-compatible gateway that handles compliance, DLP, auditing, and corporate authorization.
- **Host Sandboxing & Policy**: Sandboxed code execution with configurable tool allow/deny policies, destructive operation guardrails, and cryptographic SHA-256 audit telemetry logs.

---

## 2. Entra ID App Registration

To enable Office SSO, register an application in the Microsoft Entra admin center:

1. Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com/) as an Application Administrator.
2. Navigate to **Identity** > **Applications** > **App registrations** > **New registration**.
3. Configure the registration:
   - **Name**: `Office Agents Enterprise`
   - **Supported account types**: `Accounts in this organizational directory only (Single tenant)`
   - **Redirect URI**: Select `Single-page application (SPA)` and enter the production taskpane URL (e.g. `https://office-agents.corp.example.com/taskpane.html`).
4. Click **Register** and record the **Application (client) ID** and **Directory (tenant) ID**.

### 2.1 Configure Application ID URI & API Scope
1. Navigate to **Expose an API**.
2. Set **Application ID URI** to `api://{HOST}/{CLIENT_ID}` (e.g. `api://office-agents.corp.example.com/00000000-0000-0000-0000-000000000000`).
3. Under **Scopes defined by this API**, click **Add a scope**:
   - **Scope name**: `access_as_user`
   - **Who can consent?**: `Admins and users`
   - **Admin consent display name**: `Office Agents User Impersonation`
   - **Admin consent description**: `Allows the Office Add-in to access enterprise LLM services as the signed-in user.`
   - **State**: `Enabled`

### 2.2 Pre-authorize Microsoft Office Applications
Under **Authorized client applications**, add the Microsoft Office application client IDs with the `access_as_user` scope enabled:
- `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Office on the web)
- `d3590433-44b4-42c9-90c9-04728de7c710` (Office desktop)
- `bc59ab90-3ae6-4566-bfb4-123456789abc` (Outlook web/desktop if applicable)

---

## 3. Manifest Configuration & M365 Centralized Deployment

### 3.1 Update Production Manifests
In `packages/excel/manifest.prod.xml`, `packages/powerpoint/manifest.prod.xml`, and `packages/word/manifest.prod.xml`:
1. Verify the `<WebApplicationInfo>` block matches your Entra ID registration:
```xml
<WebApplicationInfo>
  <Id>YOUR_ENTRA_CLIENT_ID</Id>
  <Resource>api://office-agents.corp.example.com/YOUR_ENTRA_CLIENT_ID</Resource>
  <Scopes>
    <Scope>access_as_user</Scope>
  </Scopes>
</WebApplicationInfo>
```
2. Ensure placeholder domain references (`contoso.com`) are replaced with your enterprise domains and internal support URLs:
   - `<SupportUrl DefaultValue="https://support.corp.example.com/help"/>`
   - `<AppDomain>https://office-agents.corp.example.com</AppDomain>`

### 3.2 Validate Manifests
Before deployment, validate all production manifests:
```bash
pnpm validate
```

### 3.3 Deploy via Microsoft 365 Admin Center
1. Sign in to the [Microsoft 365 admin center](https://admin.microsoft.com/) as a Global Administrator or Exchange Administrator.
2. Go to **Settings** > **Integrated apps**.
3. Click **Upload custom apps**.
4. Choose **Upload manifest file (.xml) from device** and select `manifest.prod.xml` for each application (Excel, PowerPoint, Word).
5. Specify deployment scope:
   - Assign to **Specific users/groups** (e.g. pilot user group) or **Entire organization**.
6. Review and complete deployment. Centralized deployment propagates to Office desktop and web within 12–24 hours.

---

## 4. Build Configuration & Static Host Headers

### 4.1 Environment Variables
Configure build-time environment variables in your CI/CD pipeline:
| Variable | Description | Example |
|---|---|---|
| `VITE_APP_MODE` | Build mode (`enterprise` disables BYOK inputs and enforces SSO) | `enterprise` |
| `VITE_DEPLOY_URL` | Base HTTPS URL hosting add-in assets | `https://office-agents.corp.example.com` |
| `VITE_GATEWAY_URL` | Corporate LLM gateway endpoint | `https://ai-gateway.corp.example.com` |
| `VITE_MCP_HOSTS` | Space-separated list of approved MCP endpoints | `https://mcp.corp.example.com` |

### 4.2 Content Security Policy & Host Headers
Cloudflare Pages and static web servers deliver CSP and frame ancestors via `public/_headers`:
```http
/*
  X-Content-Type-Options: nosniff
  Referrer-Policy: strict-origin-when-cross-origin
  Content-Security-Policy: default-src 'self'; frame-ancestors 'self' https://*.office.com https://*.officeapps.live.com https://*.sharepoint.com https://*.office365.com; script-src 'self' https://appsforoffice.microsoft.com; connect-src 'self' https://appsforoffice.microsoft.com https://*.office.com https://*.officeapps.live.com https://ai-gateway.corp.example.com https://mcp.corp.example.com; object-src 'none';
```
*Note: `frame-ancestors` is ignored in `<meta http-equiv>` tags by browsers and must be delivered via HTTP response headers.*

---

## 5. Runtime Policy, Sandboxing & Telemetry

### 5.1 Tool Execution Policy
The SDK supports per-build policy configuration via `ToolPolicyConfig`:
```typescript
interface ToolPolicyConfig {
  allowedTools?: string[];        // Explicit whitelist of callable tools
  deniedTools?: string[];         // Blacklist of forbidden tools
  allowDestructiveOps?: boolean;  // Guardrail against code with .delete()/.clear()
}
```

### 5.2 Pluggable Audit Telemetry Sink
All tool calls emit structured audit events before and after execution:
```typescript
interface ToolAuditEvent {
  toolCallId: string;
  toolName: string;
  argsHash: string;      // SHA-256 digest of serialized arguments
  documentId: string;
  userId: string | null; // Entra ID user principal
  timestamp: number;
  durationMs: number;
  status: "success" | "error" | "policy_denied";
  errorMessage?: string;
}
```
Inject an enterprise `TelemetrySink` into the app adapter or runtime options to forward events to your centralized SIEM (Splunk, Datadog, Microsoft Sentinel).

---

## 6. Troubleshooting & Diagnostics

### 6.1 Diagnostic Bundle Export
When a user encounters issues, they can export a diagnostic bundle from the Settings panel or Error Boundary:
- Click **Export Diagnostics** or **Copy Error Details**.
- The bundle includes:
  - Add-in and SDK version information
  - Host application and Office platform version
  - Redacted session statistics and recent error logs
  - Network and gateway connectivity status

### 6.2 Common Office SSO Error Codes
| Error Code | Meaning | Remediation |
|---|---|---|
| `13001` | User is not signed in to Office | Prompt user to sign into Office desktop or web with corporate account. |
| `13003` | User canceled sign-in or consent prompt | Re-prompt user or grant admin consent in Entra ID app registration. |
| `13007` | SSO not supported on current platform | Ensure client meets minimum Office requirement (Office 365 version 16.0+). |
| `13008` | Entra ID application ID URI mismatch | Verify `<Resource>` in manifest matches `api://{HOST}/{CLIENT_ID}`. |
