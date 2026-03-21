# Adaptive Platform Maturity Roadmap

This document defines what each major platform stream must achieve to reach `5/5` maturity, where the platform stands now, and the sequence of work needed to climb each rung.

## Maturity Scale

| Score | Meaning | Operational Reality |
|------|---------|---------------------|
| `1/5` | Concept / scaffold | Mostly shapes, mocks, or contracts |
| `2/5` | Working prototype | Internal alpha, usable by builders with sharp edges |
| `3/5` | Design-partner alpha | End-to-end flows work for pilot tenants, still support-heavy |
| `4/5` | Beta-ready | Reliable, observable, governed, and test-covered |
| `5/5` | Enterprise-ready | Secure, resilient, scalable, operable, polished, and compliant |

## Current Platform Snapshot

Overall maturity: `2/5`

What exists now:
- Metadata-driven control plane for objects, fields, pages, layouts, navigation, workflows, agents, models, security, forms, branding, and profiles.
- Publish/version flow with immutable runtime manifests and Git-backed publication.
- Draft preview, admin preview, published runtime, and view-as-user foundation.
- Branded public form builder/runtime foundation.
- Workflow builder with node palette, edge wiring, config editing, run history, and draft test.
- Agent builder with provider controls, masked preview, and eval scaffold.
- Queue abstraction for BullMQ/Redis with local fallback worker.

What is still missing for a true beta:
- Real enterprise authentication and session lifecycle.
- World-class page/form/workflow/agent builder depth.
- Full workflow runtime durability and operator tooling.
- Real provider-backed agent execution plane.
- Strong white-label automation and asset processing.
- Formal UX/design gate, SLOs, and enterprise operational controls.

## Cross-Platform Requirements For `5/5`

These are mandatory for the whole product, regardless of stream.

### Security And Identity
- Real OIDC and SAML SSO.
- Short-lived signed sessions with rotation and revocation.
- RBAC plus policy enforcement at API, runtime, workflow, and agent layers.
- Audit coverage for every privileged action and impersonation event.
- Secrets stored in a managed secret system, never in manifest payloads.

### Tenancy And Isolation
- Hard tenant isolation across metadata, records, assets, workflow execution, agent execution, logs, and exports.
- Environment isolation for dev, preview, staging, and production.
- Safe tenant cloning, export, backup, restore, and rollback.

### Runtime Discipline
- Draft and published states strictly separated.
- Runtime consumes published definitions only.
- Deterministic compile/activation pipeline from draft to published manifest.
- Backward-compatible contract evolution for manifests and APIs.

### Operability
- Structured logs, metrics, traces, alerting, SLOs, and runbooks.
- Background job durability, replay, dead-letter queues, and operational controls.
- Health dashboards for builder, runtime, workflow, agent, and asset systems.

### Quality
- Unit, integration, and E2E coverage around every builder and runtime loop.
- Accessibility and responsive acceptance criteria.
- Performance budgets for studio, runtime, and public forms.
- UX review gate on all major builder flows.

### Enterprise Packaging
- Feature flags and entitlement controls.
- Admin analytics and change management.
- Documentation, onboarding, and support tooling.
- Clear compliance posture for data handling, masking, retention, and audit.

## Portfolio Scorecard

| Stream | Current | Target | Short Truth |
|------|---------|--------|-------------|
| 1. Platform kernel & contracts | `3/5` | `5/5` | Strong base, not yet a hardened contract platform |
| 2. Identity, profiles, and view-as | `2/5` | `5/5` | Working local/session foundation, no real enterprise auth yet |
| 3. Branding and white-label | `2/5` | `5/5` | Theme tokens and uploads exist, automation is shallow |
| 4. Data model and rules engine | `2.5/5` | `5/5` | Real metadata core exists, rule tooling is still early |
| 5. Page builder and admin preview | `2.5/5` | `5/5` | Useful builder, not yet world class |
| 6. Forms product | `2/5` | `5/5` | Real foundation, not yet a serious forms product |
| 7. Workflow studio | `2/5` | `5/5` | Good alpha editor, not yet enterprise orchestration design |
| 8. Workflow runtime and operator console | `1.5/5` | `5/5` | Queue scaffold exists, runtime depth does not |
| 9. Agentforce platform | `1.5/5` | `5/5` | Builder scaffolding exists, real execution and governance do not |
| 10. UX review and design quality gate | `1/5` | `5/5` | Intent exists, operating discipline does not |

---

## 1. Platform Kernel & Contracts

Current maturity: `3/5`

Current state:
- `PlatformManifest v2` exists.
- Shared service and API surfaces exist for the main platform domains.
- Publish/versioning and Git-backed manifest output exist.
- Queue abstraction exists.

### Rung Requirements

#### `1/5`
- Define core manifest types.
- Define tenant/environment/version concepts.
- Define publish as a first-class concept.

#### `2/5`
- Persist draft metadata.
- Publish immutable versions.
- Render runtime from published metadata.
- Provide basic APIs per major domain.

#### `3/5`
- Backfill older manifest versions safely.
- Provide preview vs published runtime separation.
- Establish additive schema evolution discipline.
- Add audit taxonomy and feature flags.

#### `4/5`
- Versioned public API contracts with compatibility rules.
- Formal migration framework for manifest and DB schema evolution.
- Deterministic compile pipeline with diffing, validation, and rollback guarantees.
- Full contract tests for manifest compile and runtime activation.

#### `5/5`
- Contract governance with deprecation policy and migration windows.
- Stable SDK/client contracts for builder, worker, and runtime layers.
- Multi-environment release promotion with deterministic artifact integrity.
- Disaster recovery, tenant export/import, and platform-wide rollback drills.

### What Must Happen Next

To reach `4/5`:
- Freeze manifest/API change rules.
- Add compile-stage validation and compatibility checks.
- Add manifest migration test fixtures for old tenant states.
- Add explicit artifact integrity checks around publish/rollback.

To reach `5/5`:
- Introduce versioned API/manifest compatibility policy.
- Build tenant export/import and recovery tooling.
- Add release promotion across environments with audit and approvals.

---

## 2. Identity, Profiles, And View-As

Current maturity: `2/5`

Current state:
- Local signed session flow exists.
- Membership-aware session bootstrap exists.
- Profile configuration exists.
- View-as-user cookie and runtime banner exist.

### Rung Requirements

#### `1/5`
- Local actor resolution.
- Role model defined.
- Tenant membership concept defined.

#### `2/5`
- Signed sessions.
- Tenant-aware actor context.
- Profile/settings metadata foundation.
- View-as-user builder/admin lens.

#### `3/5`
- OIDC-ready login flow.
- Tenant switcher with real session persistence.
- Profile page and settings page runtime surfaces.
- Impersonation audit trail and exit controls.

#### `4/5`
- SAML support and enterprise IdP configuration.
- Session expiry, rotation, revocation, and idle timeout handling.
- Self-service profile editing, passwordless/session management if relevant.
- Admin access review and membership lifecycle tooling.

#### `5/5`
- Full enterprise SSO matrix: OIDC, SAML, SCIM, JIT provisioning.
- Fine-grained policy controls for impersonation, tenant admin delegation, and break-glass access.
- Security reviews, penetration test coverage, and audit-grade evidence.
- Strong compliance posture for auth events and user lifecycle.

### What Must Happen Next

To reach `3/5`:
- Replace local-dev auth as primary path with real OIDC login.
- Persist tenant/user sessions through a proper auth adapter.
- Ship runtime profile and settings pages backed by live user state.
- Add view-as audit and exit surfaces everywhere.

To reach `4/5`:
- Add SAML and IdP config UI.
- Add session governance, revocation, and access review.
- Add membership invites, acceptance, suspension, and removal as complete flows.

To reach `5/5`:
- Add SCIM and enterprise provisioning.
- Add advanced impersonation controls and policy exceptions.
- Complete security hardening and compliance evidence.

---

## 3. Branding And White-Label System

Current maturity: `2/5`

Current state:
- Tenant branding model exists.
- Theme tokens drive runtime and preview shell styling.
- Asset upload routes exist.
- Studio has branding workspace.

### Rung Requirements

#### `1/5`
- Branding schema exists.
- Theme tokens exist.
- Tenant-level shell styling exists.

#### `2/5`
- Admin can edit colors, fonts, and notes.
- Admin can upload assets.
- Branding affects runtime shell and draft preview.

#### `3/5`
- Logo/icon selection and assignment flows work cleanly.
- Branding flows through pages, forms, and app shell consistently.
- Theme draft review and publish diff exist.
- Basic accessibility checks for color contrast exist.

#### `4/5`
- Brand-book ingestion with assisted extraction of palette, typography, and logo usage.
- Generated theme suggestions with manual approval.
- Asset processing pipeline for variants, favicon, crops, and responsive formats.
- Theme QA with contrast and spacing checks across core builder/runtime surfaces.

#### `5/5`
- Enterprise white-label management with multi-brand packs per tenant if needed.
- Automated asset pipeline with approval workflows and rollback.
- Full token system across shell, components, forms, and generated apps.
- Brand compliance toolkit and preview pack for stakeholder review.

### What Must Happen Next

To reach `3/5`:
- Add selected logo/icon assignment in the builder.
- Apply branding tokens consistently to forms, page builder, and runtime widgets.
- Add publish preview for branding impact.

To reach `4/5`:
- Add brand-book ingestion and theme suggestion service.
- Add asset processing for variants and metadata extraction.
- Add automated contrast/accessibility validation.

To reach `5/5`:
- Add enterprise brand packs and approval workflows.
- Add comprehensive token coverage across every rendered surface.

---

## 4. Data Model And Rules Engine

Current maturity: `2.5/5`

Current state:
- Object and field metadata builder exists.
- Validation, calculations, tooltips, help text, and mandatory rule foundations exist.
- Shared rule expression system exists.

### Rung Requirements

#### `1/5`
- Objects and fields can be defined.
- Required and basic validation rules exist.

#### `2/5`
- Objects, fields, calculations, and validations persist and publish.
- Runtime record validation works.
- Shared rule-expression model exists.

#### `3/5`
- Reusable validation sets and field groups.
- Relationship-aware rules and object bindings.
- Builder UX for advanced conditions and calculated fields.
- Server-side enforcement across forms, records, workflows, and agents.

#### `4/5`
- Formal rules engine with traceability and debugging.
- Expression library with references, test cases, and evaluation previews.
- Migration-safe object/field evolution with impact analysis.
- Safe schema change rollout for live tenants.

#### `5/5`
- Enterprise-grade schema governance, lineage, and change control.
- Reusable business-rule registry across tenants or packages where appropriate.
- Full explainability for rule evaluation in UI and logs.
- Performance guarantees for large schemas and heavy rule usage.

### What Must Happen Next

To reach `3/5`:
- Add field groups, reusable validation sets, and richer rule builder UX.
- Add relationship-aware binding UI.
- Add rule preview/test tooling in the builder.

To reach `4/5`:
- Add rule trace logs and evaluation explainability.
- Add schema impact analysis before publish.
- Add live-tenant-safe schema evolution rules.

To reach `5/5`:
- Add governance workflows, lineage, and packageable rule libraries.
- Add performance and audit guarantees for large enterprise tenants.

---

## 5. Page Builder And Admin Preview

Current maturity: `2.5/5`

Current state:
- Metadata-driven page builder exists.
- Admin preview and published runtime exist.
- Hybrid layout model exists with sections, grids, components, and placement metadata.

### Rung Requirements

#### `1/5`
- Pages, layouts, and menus exist as metadata.
- Runtime can render page definitions.

#### `2/5`
- Builder can create/edit pages and layouts.
- Draft preview works.
- Admin preview exists.

#### `3/5`
- Strong page tree, reusable templates, saved tenant templates, and sticky inspector.
- Clear drag/reorder flows for sections and components.
- Better responsive breakpoints and preview modes.
- Better page diff before publish.

#### `4/5`
- World-class authoring ergonomics: keyboard flow, fast duplication, reusable blocks, layout diagnostics, content bindings, visibility rules, and preview QA.
- Rich catalog of enterprise-ready component presets.
- Performance optimized for large page trees and multi-page apps.
- Strong accessibility guidance inside builder and runtime.

#### `5/5`
- Best-in-class business app designer with reusable packages, design token awareness, deep preview tooling, audit-friendly diffs, and high author confidence.
- Operational confidence for large tenant apps with many pages/components.
- Sophisticated preview lenses for admin, support, and end user contexts.

### What Must Happen Next

To reach `3/5`:
- Strengthen page tree and reusable template workflow.
- Improve hierarchy, component inspector, and responsive preview controls.
- Add publish diff focused on pages/routes/layouts.

To reach `4/5`:
- Add component packages, advanced visibility/binding diagnostics, and fast authoring controls.
- Add accessibility and runtime QA inside the builder.

To reach `5/5`:
- Add package/versioning model for reusable builder assets.
- Add enterprise-scale performance and governance tooling.

---

## 6. Forms Product

Current maturity: `2/5`

Current state:
- Form definitions exist in the manifest.
- Form builder exists in the studio.
- Draft preview and live published form routes exist.
- Submissions persist against the published runtime.

### Rung Requirements

#### `1/5`
- Forms are metadata, not hard-coded pages.
- Runtime can render a simple form.

#### `2/5`
- Multi-step form builder exists.
- Branded form rendering exists.
- Save draft and submit foundations exist.
- Submission capture exists.

#### `3/5`
- Conditional branching between steps.
- Rich field settings: tooltips, mandatory logic, calculations, validation sets.
- Embed mode, authenticated mode, and public mode fully supported.
- Submission inbox and detail views are usable.

#### `4/5`
- Typeform/Jotform-class authoring UX.
- Save-and-resume with durable respondent sessions.
- Share links, embed codes, confirmation flows, analytics, and operational filters.
- Form testing and preview with sample paths.

#### `5/5`
- Enterprise forms platform with workflow triggers, auditability, access controls, advanced analytics, versioned releases, and scale/performance guarantees.
- Rich widgets, conditional sections, uploads, localization, accessibility, and governance.

### What Must Happen Next

To reach `3/5`:
- Add branching and conditional visibility in the builder/runtime.
- Add submission inbox/detail views in the studio.
- Add authenticated and embedded delivery flows properly.

To reach `4/5`:
- Add durable respondent sessions and save-and-resume restoration.
- Add embed/share tooling and submission analytics.
- Add richer field library and form test runner.

To reach `5/5`:
- Add enterprise governance, localization, analytics depth, and large-scale ops tooling.

---

## 7. Workflow Studio

Current maturity: `2/5`

Current state:
- Visual node palette exists.
- Nodes, edges, and config editing exist.
- Draft test exists.
- Run history is visible.

### Rung Requirements

#### `1/5`
- Workflow schema and node types exist.
- Builder can define a graph.

#### `2/5`
- Node palette, edge creation, config editing, and run launch exist.
- Basic run visibility exists.

#### `3/5`
- Trigger builder, templates, branch labels, validation hints, and test harness are strong.
- Better graph ergonomics and readable edge semantics.
- Workflow compare and publish diff exist.

#### `4/5`
- N8N-class authoring quality for business automation flows.
- Reusable subflows/templates, sample payload testing, dry runs, and diagnostics.
- Strong validation, change impact, and linting before publish.

#### `5/5`
- Enterprise orchestration design surface with packaged workflow assets, governance, approvals, reuse, simulation, and confidence tooling.
- High usability at large graph sizes.

### What Must Happen Next

To reach `3/5`:
- Add better trigger builder and template system.
- Add graph validation and test payload workflows.
- Add workflow diff before publish.

To reach `4/5`:
- Add reusable subflows, better graph UX, richer branch semantics, and workflow linting.
- Add scenario simulation and operator handoff views.

To reach `5/5`:
- Add packaging, approvals, governance, and scale handling for large enterprise workflow libraries.

---

## 8. Workflow Runtime And Operator Console

Current maturity: `1.5/5`

Current state:
- Worker loop exists.
- BullMQ/Redis queue abstraction exists.
- Basic run creation and logs exist.
- Published-only discipline is partially enforced.

### Rung Requirements

#### `1/5`
- Workflow runs can be queued and processed.
- Logs can be written.

#### `2/5`
- Runs persist.
- Local worker and queue abstraction both work.
- Basic status changes are visible.

#### `3/5`
- Retries, idempotency, pause/resume, wait states, approval tasks, and scheduling work correctly.
- Operator console supports run inspection and replay.
- Published-only runtime execution is fully enforced.

#### `4/5`
- Dead-letter queues, stuck-run detection, step-level replay, approval inboxes, and operational alerting exist.
- Strong runtime SLOs and metrics exist.
- Large-run-volume performance is acceptable.

#### `5/5`
- Enterprise-grade durable execution with recoverability, throttling, scheduling, replay safety, runbooks, observability, and compliance-grade audit.
- Operational tooling is strong enough for support and platform teams to trust it under load.

### What Must Happen Next

To reach `3/5`:
- Finish BullMQ-backed execution semantics.
- Add retries, pause/resume, wait nodes, approval tasks, and scheduling.
- Add run detail UI and replay basics.

To reach `4/5`:
- Add DLQ, alerts, stuck-run detection, step replay, and approval inboxes.
- Add metrics and tracing.

To reach `5/5`:
- Add enterprise runbooks, capacity controls, throttling, and recovery operations.

---

## 9. Agentforce Platform

Current maturity: `1.5/5`

Current state:
- Agent metadata builder exists.
- Provider policy, masked preview, and eval scaffold exist.
- Security policy and zero-retention gating foundations exist.

### Rung Requirements

#### `1/5`
- Agent metadata model exists.
- Provider registry exists.
- Masking policy model exists.

#### `2/5`
- Builder can define prompts, object scope, tool scope, and provider settings.
- Preview and eval scaffolding exists.

#### `3/5`
- Real provider-backed execution through a server gateway.
- Tool invocation, workflow handoff, and scoped access actually work.
- Simulation and eval workflows are usable for builders.

#### `4/5`
- Agentforce-class builder depth with prompt blocks, role instructions, output schemas, handoffs, policy diagnostics, and runtime observability.
- Policy enforcement and redacted logs are robust.
- Operator views show agent usage, failures, and quality metrics.

#### `5/5`
- Enterprise agent platform with provider routing, eval suites, safe rollout controls, approval chains for sensitive actions, policy explainability, and model governance.
- Strong trust, audit, and reliability posture around AI execution.

### What Must Happen Next

To reach `3/5`:
- Add real provider-backed execution service.
- Add tool execution policy and workflow handoff.
- Add simulation/eval UX that feels first class.

To reach `4/5`:
- Add output schemas, prompt blocks, routing diagnostics, operator telemetry, and stronger safety posture controls.

To reach `5/5`:
- Add eval suites, rollout controls, governance workflows, and enterprise AI audit posture.

---

## 10. UX Review And Design Quality Gate

Current maturity: `1/5`

Current state:
- Intent exists.
- Some visual refinement has happened.
- No enforced design review operating model exists.

### Rung Requirements

#### `1/5`
- Visual direction exists informally.
- Builder is usable by the core team.

#### `2/5`
- Shared design tokens and UI patterns are documented.
- Builder flows are reviewed for obvious friction and accessibility issues.

#### `3/5`
- Every stream has UX acceptance criteria, empty-state standards, copy standards, and accessibility checks.
- Review gate exists before merge for major builder flows.

#### `4/5`
- Dedicated design QA across studio, runtime, and public surfaces.
- Task-completion benchmarks for new admins.
- Consistent visual grammar and polished interaction model.

#### `5/5`
- Design-system-level governance with measurable usability goals, accessibility compliance, strong information architecture, and consistent beauty across the whole platform.
- The platform feels simple despite depth.

### What Must Happen Next

To reach `2/5`:
- Write design guardrails for hierarchy, copy, spacing, empty states, and motion.
- Define review checklist for every major builder flow.

To reach `3/5`:
- Make UX signoff a merge requirement for each pillar.
- Add accessibility and first-task usability acceptance criteria.

To reach `4/5`:
- Add task benchmarks, heuristic reviews, and design QA loops.
- Tighten the platform-wide visual grammar.

To reach `5/5`:
- Turn the design gate into an operating discipline with measurable standards and recurring audits.

---

## Recommended Climb Sequence

If the goal is to reach `5/5` everywhere as fast as possible without building on sand, the order should be:

1. Platform kernel and contracts to stable `4/5`
2. Identity and tenancy to `3/5`
3. Workflow runtime to `3/5`
4. Page builder and forms to `3/5`
5. Agentforce to `3/5`
6. Branding to `3/5`
7. UX review gate to `3/5`
8. Then drive all pillars from `3/5` to `4/5`
9. Reserve `5/5` work for the streams that prove product-market pull first

## Immediate Next Tranche

The next highest-value tranche is:

1. Real OIDC authentication and membership-backed sessions
2. Workflow runtime hardening: retries, pause/resume, approvals, scheduling, DLQ base
3. Forms depth: branching, inbox, authenticated/embedded delivery
4. Page builder depth: templates, component hierarchy, responsive diagnostics
5. Agent execution: real provider-backed gateway and workflow handoffs
6. UX gate: documented standards and enforced review checklist

If those land, the platform moves from `2/5` internal alpha to an honest `3/5` design-partner alpha.
