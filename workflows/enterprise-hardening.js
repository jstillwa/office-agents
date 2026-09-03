export const meta = {
  name: "enterprise_hardening",
  description:
    "Take office-agents to an internal-enterprise production state: parallel audit, adversarial verification, plan synthesis, human gate, worktree-isolated implementation with validated gates, then integration.",
  phases: [
    { title: "Audit" },
    { title: "Verify" },
    { title: "Plan" },
    { title: "Approve" },
    { title: "Implement" },
    { title: "Integrate" },
  ],
};

// ---------------------------------------------------------------------------
// Args (all nondeterminism and human decisions enter here)
//   mode:            "audit" | "plan" | "full"          (default "plan")
//   runDate:         ISO date string for report headers  (required for full)
//   maxWorkstreams:  cap on implementation lanes         (default 8, max 10)
//   implementRounds: gate attempts per lane              (default 3, max 4)
//   skipVerify:      skip adversarial pass               (default false)
//   plan:            pre-approved plan object {summary, workstreams, deferred}
//                    (e.g. contents of workflows/enterprise-hardening.plan.json).
//                    When set, Audit/Verify/Plan are skipped.
//   model:           exact provider/modelId for ALL subagents
//                    Default: bifrost/gemini/gemini-flash-latest (owner request).
//                    An unavailable explicit model THROWS (no silent fallback).
// ---------------------------------------------------------------------------
const mode = ["audit", "plan", "full"].includes(args?.mode) ? args.mode : "plan";
const runDate = typeof args?.runDate === "string" ? args.runDate : "unspecified";
const maxWorkstreams = Number.isInteger(args?.maxWorkstreams)
  ? Math.max(1, Math.min(args.maxWorkstreams, 10))
  : 8;
const implementRounds = Number.isInteger(args?.implementRounds)
  ? Math.max(1, Math.min(args.implementRounds, 4))
  : 3;
const skipVerify = args?.skipVerify === true;
const presetPlan = args?.plan && Array.isArray(args.plan.workstreams) ? args.plan : null;
const MODEL =
  typeof args?.model === "string" && args.model.trim()
    ? args.model.trim()
    : "bifrost/gemini/gemini-flash-latest";
// Every agent()/verify() call below passes { model: MODEL } so the main
// orchestrating session spends no tokens on subagent work.

// Decisions already made with the owner. Every agent sees these.
const DECISIONS = `
PRODUCT DECISIONS (fixed — do not re-litigate):
- Audience: single-org internal enterprise deployment via M365 centralized deployment. Not multi-tenant SaaS.
- Credentials: production uses a corporate LLM gateway (OpenAI-compatible or Anthropic-compatible endpoint) authenticated with an Entra ID token from Office SSO (OfficeRuntime.auth.getAccessToken). Zero provider API keys in the browser in "enterprise" build mode. BYOK stays available only in "dev" build mode.
- Sandbox: keep eval_officejs / edit_slide_xml / bash. Add a per-build policy (allow/deny tool list, destructive-op flags) and an audit log of every tool execution (tool name, args hash, doc id, user, timestamp) sent to a pluggable sink.
- Style: shortest working diff. Reuse what exists in the repo before writing new code. Stdlib/platform before dependencies. No new dependency without justification. Every non-trivial change leaves one runnable test.
- Quality bar: pnpm check && pnpm test && pnpm build must pass on every lane.
`;

// Owner corrections after reviewing the first plan synthesis (2026-09-02).
const PLAN_CONSTRAINTS = `
OWNER PLAN CONSTRAINTS (apply when synthesizing workstreams):
- CSP: frame-ancestors is ignored in <meta http-equiv>; deliver it ONLY via public/_headers. connect-src must be built from VITE_* vars (gateway host, MCP hosts) in vite.config.ts, never hardcoded. Do not put CSP in a <meta> tag at all if _headers covers it.
- SES lockdown: do NOT mandate eager lockdown() before Office.onReady; it is known to break Office.js init. Make it a spike step with acceptance "documented go/no-go", keep lazy lockdown as the default.
- Dev bridge: drop bearer-token auth. Origin allowlist (https://localhost:3000-3002) + 127.0.0.1 bind + devDependencies move is sufficient for a dev-only tool.
- Destructive-op detection on eval code strings is a guardrail against accidental LLM behaviour, NOT a security control. Say so in the step text. The real controls are the tool allow/deny policy and audit log from the SDK seam.
- Split the SSO/gateway lane: one workstream for Office SSO -> Entra token -> gateway routing + zero-BYOK storage + settings-panel lockdown; a separate workstream for SSRF guard on web fetch/search + MCP HTTPS enforcement. Neither should be effort L.
- Do not exceed 8 workstreams; merge or defer to stay within the cap.
`;

const REPO_FACTS = `
REPO FACTS (verified ${runDate}):
- pnpm monorepo: packages/{sdk,core,bridge,excel,powerpoint,word}. Svelte 5 + Vite + Tailwind. Biome lint/format. Vitest.
- CI (.github/workflows/ci.yml) pins pnpm 9 and Node 20; release-*.yml use Node 24; package.json has no "packageManager". pnpm-workspace.yaml already has pnpm-11 "allowBuilds".
- Secrets: ProviderConfig.apiKey and OAuth tokens stored plaintext in localStorage (sdk/src/provider-config.ts, sdk/src/oauth/index.ts). MCP config likewise (sdk/src/mcp/index.ts).
- Prod URLs hardcoded to https://open{excel,ppt-9p7,word}.pages.dev in packages/*/vite.config.ts and manifest.prod.xml; SupportUrl/AppDomain use contoso.com placeholders.
- No CSP, no _headers, no security scanning, no dependency audit, no SBOM, no CODEOWNERS, no SECURITY.md.
- ~140 explicit "any", ~51 raw console.* calls; biome has noExplicitAny off.
- 18 test files, almost all in packages/sdk/tests. Excel/PPT/Word tool packages have near-zero tests.
- Sandbox: sdk/src/lockdown.ts (SES) + core sandboxedEval. eval_officejs (excel), edit_slide_xml (ppt), bash (all) can run arbitrary code inside the taskpane.
- Runtime: sdk/src/runtime.ts (AgentRuntime, ~900 lines) owns agent lifecycle, tool wiring, session persistence (IndexedDB via idb), and onToolResult adapter callback.
- Recently merged: edit_file tool + edit-slide-xml CLI (PR #20), MCP client (PR #24).
`;

// ---------------------------------------------------------------------------
// Phase 1 — Audit: independent read-only lenses
// ---------------------------------------------------------------------------
const LENSES = [
  {
    id: "secrets-auth",
    focus:
      "Credential storage, OAuth flows, MCP server headers, proxy config, and what an enterprise gateway + Office SSO integration must touch. Enumerate every place apiKey/access tokens are read or written.",
  },
  {
    id: "sandbox-exec",
    focus:
      "Arbitrary code execution surface: eval_officejs, edit_slide_xml, edit-slide-xml CLI, bash custom commands, sandboxedEval, SES lockdown config. What can escape, what can destroy user data, where a policy check and audit hook would sit with the fewest edits.",
  },
  {
    id: "supply-chain-ci",
    focus:
      "CI/CD, pnpm/Node version drift, lockfile integrity, postinstall scripts (allowBuilds), release workflows, missing scanning (CodeQL, gitleaks, pnpm audit), Dependabot/Renovate, SBOM, provenance, branch protection needs.",
  },
  {
    id: "web-security",
    focus:
      "CSP, iframe/taskpane threat model, mixed content, CORS proxy usage, fetch of untrusted URLs (web/fetch.ts, search.ts), image handling, markdown rendering XSS (core/src/chat/markdown.ts), postMessage usage.",
  },
  {
    id: "type-safety-errors",
    focus:
      "Explicit any usage, unchecked casts, swallowed errors, raw console.* logging, missing error boundaries in Svelte components, unhandled promise rejections in runtime.ts and adapters. Rank by blast radius.",
  },
  {
    id: "test-coverage",
    focus:
      "Existing tests vs. risk. Identify the 15 highest-value missing tests across sdk runtime, mcp client (SSE parsing, session id), edit-file tool, and Office.js tool wrappers in excel/powerpoint/word (with a minimal Office.js mock strategy that already fits vitest + happy-dom).",
  },
  {
    id: "deploy-manifests",
    focus:
      "manifest.xml / manifest.prod.xml correctness for centralized deployment, hardcoded URLs, version sync, icon/asset hosting, static host headers, environment parameterisation (VITE_* vars), and what an M365 admin needs to deploy.",
  },
  {
    id: "observability-ops",
    focus:
      "What exists for telemetry, error reporting, version display, session diagnostics, bridge (packages/bridge) exposure in prod builds. Propose the minimal pluggable telemetry sink interface and where runtime.ts should emit.",
  },
];

const findingSchema = {
  type: "object",
  properties: {
    lens: { type: "string" },
    findings: {
      type: "array",
      items: {
        type: "object",
        properties: {
          id: { type: "string" },
          title: { type: "string" },
          severity: { type: "string", enum: ["critical", "high", "medium", "low"] },
          files: { type: "array", items: { type: "string" } },
          evidence: { type: "string" },
          fix: { type: "string" },
          effort: { type: "string", enum: ["S", "M", "L"] },
        },
        required: ["id", "title", "severity", "files", "evidence", "fix", "effort"],
      },
    },
  },
  required: ["lens", "findings"],
};

phase("Audit");
log(`mode=${mode} model=${MODEL} lenses=${LENSES.length} presetPlan=${!!presetPlan}`);

const auditResults = presetPlan ? LENSES.map(() => null) : await parallel(
  LENSES.map((lens, i) => () =>
    agent(
      `You are auditing the repository at ${cwd} (read-only: do NOT modify files).
${DECISIONS}
${REPO_FACTS}

LENS: ${lens.id}
${lens.focus}

Read the actual code. Cite exact file paths and line-level evidence. Max 12 findings; prefer fewer, sharper ones. Each finding id must be "${lens.id}-<n>". For "fix", describe the shortest diff that resolves the root cause, naming existing helpers/patterns in the repo to reuse. Do not propose new dependencies unless nothing in the repo or platform covers it.`,
      { label: `audit:${i}:${lens.id}`, schema: findingSchema, model: MODEL },
    ),
  ),
);

const auditLedger = LENSES.map((lens, i) => ({
  id: lens.id,
  status: auditResults[i] === null ? "failed" : "complete",
  result: auditResults[i],
}));

const allFindings = auditLedger
  .filter((e) => e.status === "complete")
  .flatMap((e) => e.result.findings.map((f) => ({ ...f, lens: e.id })));

log(`audit complete: ${allFindings.length} findings, ${auditLedger.filter((e) => e.status === "failed").length} lenses failed`);

if (mode === "audit") {
  return { mode, runDate, auditLedger, findings: allFindings };
}

// ---------------------------------------------------------------------------
// Phase 2 — Verify: skeptics attack critical/high findings
// ---------------------------------------------------------------------------
phase("Verify");

const toVerify = allFindings
  .filter((f) => f.severity === "critical" || f.severity === "high")
  .slice(0, 24);

const verifyLedger = [];
if (!skipVerify && !presetPlan) {
  const verdicts = await parallel(
    toVerify.map((f) => () =>
      verify(
        { finding: f, repo: cwd, decisions: DECISIONS },
        {
          model: MODEL,
          reviewers: 2,
          threshold: 0.5,
          lens: [
            "Is this finding real in the current code? Open the cited files and confirm the evidence. Reject if the cited behaviour does not exist or is already mitigated elsewhere in the repo.",
            "Is the proposed fix the shortest root-cause fix, or does it treat a symptom / add unrequested abstraction? Reject if a smaller, already-present mechanism covers it.",
          ],
        },
      ),
    ),
  );
  toVerify.forEach((f, i) => {
    verifyLedger.push({
      id: f.id,
      status: verdicts[i] ? "verified" : "failed",
      real: verdicts[i] ? verdicts[i].real : null,
      votes: verdicts[i] ? verdicts[i].votes : [],
    });
  });
}

const rejected = new Set(verifyLedger.filter((v) => v.real === false).map((v) => v.id));
const vettedFindings = allFindings.filter((f) => !rejected.has(f.id));
log(`verify: ${toVerify.length} checked, ${rejected.size} rejected`);

// ---------------------------------------------------------------------------
// Phase 3 — Plan: synthesize into ordered workstreams
// ---------------------------------------------------------------------------
phase("Plan");

const planSchema = {
  type: "object",
  properties: {
    summary: { type: "string" },
    workstreams: {
      type: "array",
      items: {
        type: "object",
        properties: {
          id: { type: "string" },
          title: { type: "string" },
          goal: { type: "string" },
          findingIds: { type: "array", items: { type: "string" } },
          files: { type: "array", items: { type: "string" } },
          steps: { type: "array", items: { type: "string" } },
          acceptance: { type: "array", items: { type: "string" } },
          dependsOn: { type: "array", items: { type: "string" } },
          effort: { type: "string", enum: ["S", "M", "L"] },
          wave: { type: "integer" },
        },
        required: ["id", "title", "goal", "findingIds", "files", "steps", "acceptance", "dependsOn", "effort", "wave"],
      },
    },
    deferred: { type: "array", items: { type: "string" } },
  },
  required: ["summary", "workstreams", "deferred"],
};

const plan = presetPlan ?? await agent(
  `Synthesize an implementation plan for taking this repo to an internal-enterprise production state.
${DECISIONS}
${PLAN_CONSTRAINTS}
${REPO_FACTS}

VETTED FINDINGS (${vettedFindings.length}):
${JSON.stringify(vettedFindings)}

REJECTED BY SKEPTICS (do not include): ${JSON.stringify([...rejected])}
AUDIT LENSES THAT FAILED (coverage gap — say so in summary): ${JSON.stringify(auditLedger.filter((e) => e.status === "failed").map((e) => e.id))}

Rules:
- Produce at most ${maxWorkstreams} workstreams. Group findings by the files they touch so lanes do not collide; two workstreams must not edit the same file in the same wave.
- Assign "wave" 1..3. Wave 1 = foundations (CI/toolchain pin, config/env parameterisation, logger/telemetry seam, policy seam). Wave 2 = features that depend on those seams (gateway auth, audit log, CSP, tests). Wave 3 = docs/runbooks and anything depending on wave 2.
- "steps" are concrete: file, what changes, what existing helper to reuse. "acceptance" is machine-checkable where possible (a test name, a grep that must return zero, a command that must exit 0).
- Every workstream's acceptance must include: pnpm check, pnpm test, pnpm build pass.
- Put genuinely speculative items in "deferred" with one-line reasons (YAGNI).
- Use stable ids "ws-01".."ws-${String(maxWorkstreams).padStart(2, "0")}".`,
  { label: "plan:synthesize", schema: planSchema, model: MODEL },
);

if (plan === null) {
  return { mode, runDate, auditLedger, verifyLedger, findings: vettedFindings, plan: null, error: "plan synthesis returned no result" };
}

const workstreams = plan.workstreams.slice(0, maxWorkstreams);
log(`plan: ${workstreams.length} workstreams, ${plan.deferred.length} deferred`);

if (mode === "plan") {
  return { mode, runDate, auditLedger, verifyLedger, findings: vettedFindings, plan: { ...plan, workstreams } };
}

// ---------------------------------------------------------------------------
// Phase 4 — Approve: human gate (foreground only; headless aborts)
// ---------------------------------------------------------------------------
phase("Approve");

const approved = presetPlan
  ? true
  : await checkpoint(
      `Implement ${workstreams.length} workstreams in isolated worktrees?\n\n${workstreams
        .map((w) => `[wave ${w.wave}] ${w.id} ${w.title} (${w.effort})`)
        .join("\n")}\n\nDeferred: ${plan.deferred.length} items.`,
      { headless: "abort" },
    );

if (approved !== true) {
  return { mode, runDate, auditLedger, verifyLedger, plan: { ...plan, workstreams }, implementation: null, aborted: "not approved" };
}

// ---------------------------------------------------------------------------
// Phase 5 — Implement: wave-ordered, worktree-isolated, gated by review
// ---------------------------------------------------------------------------
phase("Implement");

const implSchema = {
  type: "object",
  properties: {
    branch: { type: "string" },
    changedFiles: { type: "array", items: { type: "string" } },
    testsAdded: { type: "array", items: { type: "string" } },
    checksPassed: { type: "boolean" },
    testsPassed: { type: "boolean" },
    buildPassed: { type: "boolean" },
    summary: { type: "string" },
    residualRisks: { type: "array", items: { type: "string" } },
  },
  required: ["branch", "changedFiles", "testsAdded", "checksPassed", "testsPassed", "buildPassed", "summary", "residualRisks"],
};

const reviewSchema = {
  type: "object",
  properties: {
    approve: { type: "boolean" },
    blocking: { type: "array", items: { type: "string" } },
    nits: { type: "array", items: { type: "string" } },
  },
  required: ["approve", "blocking", "nits"],
};

const implLedger = [];
const waves = [...new Set(workstreams.map((w) => w.wave))].sort((a, b) => a - b);

for (const wave of waves) {
  const lanes = workstreams.filter((w) => w.wave === wave);
  log(`wave ${wave}: ${lanes.length} lanes`);

  const priorSummaries = implLedger
    .filter((e) => e.status === "merged-ready")
    .map((e) => `${e.id} (${e.result.branch}): ${e.result.summary}`)
    .join("\n");

  const results = await parallel(
    lanes.map((ws) => async () => {
      const attempts = [];
      const outcome = await gate(
        async (feedback, attempt) => {
          const value = await agent(
            `You are implementing one workstream in an isolated git worktree of ${cwd}. Commit your work on a branch named "hardening/${ws.id}".
${DECISIONS}

WORKSTREAM: ${JSON.stringify(ws)}

PRIOR WAVES ALREADY IMPLEMENTED (their branches exist; rebase onto or cherry-pick from them if your work depends on them, otherwise ignore):
${priorSummaries || "(none)"}

${feedback ? `REVIEWER FEEDBACK TO ADDRESS:\n${feedback}\n` : ""}
Process:
1. Read every file you will touch first. Trace the real flow. Then pick the smallest change that satisfies "acceptance".
2. Reuse existing helpers/types/patterns in the repo. No new dependencies unless a step explicitly names one and nothing existing covers it.
3. Leave one runnable test per non-trivial change (vitest, no new frameworks).
4. Run: pnpm check && pnpm test && pnpm build. Report honestly; do not claim green if not green.
5. Commit with a conventional message. Report the branch name and changed files.`,
            {
              label: `impl:${ws.id}:${attempt + 1}`,
              isolation: "worktree",
              schema: implSchema,
              model: MODEL,
            },
          );
          attempts.push({ attempt: attempt + 1, feedback: feedback ?? null, value, review: null });
          return value;
        },
        async (value) => {
          const entry = attempts[attempts.length - 1];
          if (value === null) {
            return { ok: false, feedback: "Implementation returned no result. Retry from scratch." };
          }
          if (!value.checksPassed || !value.testsPassed || !value.buildPassed) {
            return { ok: false, feedback: `Quality bar not met (check=${value.checksPassed} test=${value.testsPassed} build=${value.buildPassed}). Fix and rerun.` };
          }
          if (value.changedFiles.length === 0) {
            return { ok: false, feedback: "No files changed. The workstream requires code changes." };
          }
          const review = await agent(
            `Review branch "${value.branch}" in ${cwd} against main (git diff main...${value.branch}). Read-only.
${DECISIONS}
WORKSTREAM: ${JSON.stringify({ id: ws.id, title: ws.title, goal: ws.goal, acceptance: ws.acceptance })}
IMPLEMENTER REPORT: ${JSON.stringify(value)}

Block only on: acceptance criteria not met, security regression, data-loss risk, unrequested abstraction/dependency, missing test for non-trivial logic, or dishonest report (verify claims by running pnpm check/test yourself if cheap). Everything else is a nit.`,
            { label: `review:${ws.id}:${attempts.length}`, schema: reviewSchema, model: MODEL },
          );
          entry.review = review;
          if (review === null) {
            return { ok: false, feedback: "Reviewer produced no verdict; re-run checks and resubmit." };
          }
          return review.approve
            ? { ok: true }
            : { ok: false, feedback: review.blocking.join("\n") };
        },
        { attempts: implementRounds },
      );
      return { outcome, attempts };
    }),
  );

  lanes.forEach((ws, i) => {
    const r = results[i];
    implLedger.push({
      id: ws.id,
      wave,
      status: r === null ? "failed" : r.outcome.ok ? "merged-ready" : "exhausted",
      result: r === null ? null : r.outcome.value,
      attempts: r === null ? [] : r.attempts,
    });
  });
}

// ---------------------------------------------------------------------------
// Phase 6 — Integrate: merge ready branches, final green, release notes
// ---------------------------------------------------------------------------
phase("Integrate");

const ready = implLedger.filter((e) => e.status === "merged-ready");
const integration =
  ready.length === 0
    ? null
    : await agent(
        `Integrate the following branches into a single branch "hardening/integration" off main in ${cwd}, in wave order. Resolve conflicts minimally, preferring the later wave. Then run pnpm check && pnpm test && pnpm build. Do NOT push and do NOT merge into main.
${DECISIONS}

BRANCHES (wave order):
${ready.map((e) => `wave ${e.wave} ${e.id} -> ${e.result.branch}`).join("\n")}

LANES THAT DID NOT LAND (mention in release notes as follow-ups): ${JSON.stringify(implLedger.filter((e) => e.status !== "merged-ready").map((e) => e.id))}

Report: integration branch name, final check/test/build status, conflicts resolved, and release notes in CHANGELOG style grouped by package.`,
        {
          label: "integrate:final",
          model: MODEL,
          schema: {
            type: "object",
            properties: {
              branch: { type: "string" },
              checksPassed: { type: "boolean" },
              testsPassed: { type: "boolean" },
              buildPassed: { type: "boolean" },
              conflicts: { type: "array", items: { type: "string" } },
              releaseNotes: { type: "string" },
            },
            required: ["branch", "checksPassed", "testsPassed", "buildPassed", "conflicts", "releaseNotes"],
          },
        },
      );

return {
  mode,
  runDate,
  model: MODEL,
  auditLedger,
  verifyLedger,
  plan: { ...plan, workstreams },
  implementation: implLedger,
  integration,
};
