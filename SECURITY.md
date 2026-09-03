# Security Policy

## Enterprise Security Context

Office Agents is deployed in enterprise Microsoft 365 environments via Centralized Deployment.
In production enterprise environments:
- **Zero In-Browser Provider Keys**: Authentication to LLM infrastructure routes exclusively through enterprise gateways authenticated via Entra ID tokens (`OfficeRuntime.auth.getAccessToken`). Bring-Your-Own-Key (BYOK) mode is strictly limited to local development builds.
- **Sandboxed Execution & Tool Auditing**: Document operations and script evaluations execute within bounded sandboxes with explicit allow/deny policies and immutable audit trails.
- **Supply Chain Hardening**: Releases require automated static analysis (CodeQL), secret detection (Gitleaks), pinned CI/CD action commit hashes, dependency auditing, and CycloneDX Software Bill of Materials (SBOM) tracking.

## Supported Versions

Only the latest release of each active package receives security updates.

| Package | Supported Version |
| ------- | ----------------- |
| `@office-agents/excel` | Latest tag (`excel-v*`) |
| `@office-agents/powerpoint` | Latest tag (`ppt-v*`) |
| `@office-agents/word` | Latest tag (`word-v*`) |
| `@office-agents/sdk` | Latest tag (`sdk-v*`) |
| `@office-agents/bridge` | Latest tag (`bridge-v*`) |

## Reporting a Vulnerability

We take the security of Office Agents seriously. If you identify or suspect a security vulnerability, please disclose it responsibly.

### Enterprise Disclosure Process

1. **Do not report security vulnerabilities through public GitHub issues or discussions.**
2. Email your vulnerability report to `security@nucleix.com` (or contact the repository maintainers directly via private channels).
3. Include detailed information in your report:
   - Affected package(s) and version(s)
   - Description of the vulnerability and attack vector
   - Step-by-step reproduction instructions or proof-of-concept (PoC)
   - Impact assessment on enterprise M365 environments, document integrity, or credential isolation
   - Suggested mitigations, if known

### Response Timelines & SLAs

- **Initial Acknowledgment**: Within 24 hours of receipt.
- **Triage & Severity Assessment**: Within 72 hours of acknowledgment.
- **Resolution & Patch Target**: Critical/High severity issues are targeted for remediation within 7 business days; Moderate/Low severity within 30 business days.
- **Coordinated Disclosure**: We request a 90-day embargo period before public disclosure to allow enterprise tenants sufficient time to apply patches.

## Security Governance & Code Ownership

Changes to the following sensitive components require explicit security review and code owner approval:
- GitHub Actions workflows (`.github/workflows/`)
- Toolchain configuration (`package.json`, `pnpm-workspace.yaml`, `.github/dependabot.yml`)
- Office Add-in manifests (`manifest.xml`, `manifest.prod.xml`)
- Authentication, SSO tokens, and LLM gateway routing modules
- Sandboxed evaluation and tool execution audit logging
