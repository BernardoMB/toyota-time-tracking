# .NET 10 Upgrade Plan

## Table of Contents

- [Executive Summary](#executive-summary)
- [Migration Strategy](#migration-strategy)
- [Detailed Dependency Analysis](#detailed-dependency-analysis)
- [Project-by-Project Plans](#project-by-project-plans)
- [Package Update Reference](#package-update-reference)
- [Breaking Changes Catalog](#breaking-changes-catalog)
- [Testing & Validation Strategy](#testing--validation-strategy)
- [Risk Management](#risk-management)
- [Complexity & Effort Assessment](#complexity--effort-assessment)
- [Source Control Strategy](#source-control-strategy)
- [Success Criteria](#success-criteria)
- [Appendix](#appendix)

---

## Executive Summary

[To be filled]

---

## Migration Strategy

[To be filled]

---

## Detailed Dependency Analysis

[To be filled]

---

## Project-by-Project Plans

[To be filled]

---

## Package Update Reference

[To be filled]

---

## Breaking Changes Catalog

[To be filled]

---

## Testing & Validation Strategy

[To be filled]

---

## Risk Management

[To be filled]

---

## Complexity & Effort Assessment

[To be filled]

---

## Source Control Strategy

[To be filled]

---

## Success Criteria

[To be filled]

---

## Appendix

[To be filled]


---

## Implementation Notes (meta)

- Generated from assessment: `C:\Users\bmond\Documents\Toyota\toyota-time-tracking\.github\upgrades\scenarios\new-dotnet-version_ff1d79\assessment.md`.
- Upgrade target: `.NET 10.0 (net10.0)` as recommended by initialization tooling.
- Source branch: `master` → Upgrade branch: `upgrade-to-NET10` (recommended).

---

## Executive Summary

### Scenario
Upgrade solution `HoursApp.sln` to target `.NET 10.0 (net10.0)` using an All-At-Once strategy.

### High-level Metrics (from assessment)
- Total projects: 1 (all require upgrade)
- Total NuGet packages: 9 (6 need upgrade)
- Total code files: 4
- Total LOC: 382
- Projects with issues: `HoursApp.csproj` (source incompatibility + package updates + deprecated package)

### Selected Strategy
**All-At-Once Strategy** — Rationale:
- Single-project solution (simple)
- Low codebase size (382 LOC)
- Low dependency depth (no project-to-project dependencies)
- Assessment shows package updates are available and no critical blocking cycles

This strategy minimizes orchestration overhead and upgrades all project files and packages in one coordinated operation.

---

## Migration Strategy

### Approach
- Atomic upgrade: update all project TargetFramework values and all recommended NuGet package versions across the solution in a single coordinated pass.
- After updates, restore dependencies and build the entire solution to observe and address compilation issues.
- Run all test projects (none present besides solution itself) and validation steps after atomic upgrade.

### Key Principles (All-At-Once)
- Treat the upgrade as a single atomic operation affecting all projects simultaneously.
- Consolidate project file changes, package updates, dependency restore and initial build & fixes into one upgrade pass.
- Use a single upgrade branch (`upgrade-to-NET10`) for all changes, with one logical commit (or a few related commits) describing the atomic upgrade.

---

## Detailed Dependency Analysis

### Dependency Graph Summary
- Number of projects: 1
- No inter-project dependencies detected.
- The single project `HoursApp.csproj` is both leaf and root; it depends on NuGet packages only.

### Migration Phasing (All-At-Once)
- Phase 0: Preparation — SDK, branch, pending changes handling
- Phase 1: Atomic Upgrade — update project TargetFramework and package references in `HoursApp.csproj` and any shared MSBuild imports
- Phase 2: Build & Fix — restore, build, resolve compilation errors and API changes
- Phase 3: Test & Validate — run available tests and validation checks

---

## Project-by-Project Plans

### Project: `HoursApp.csproj`

**Current State**
- Current Target Framework: `net8.0-windows`
- SDK-style: True
- Project kind: `DotNetCoreApp`
- Files: 4
- LOC: 382
- Issues found in assessment: source incompatibility for selected .NET version; recommended NuGet upgrades; one deprecated NuGet package

**Target State**
- Target Framework: `net10.0-windows` (net10.0)
- All package references updated to suggested versions (see §Package Update Reference)

**Migration Steps (what the executor will do — planner records these steps only)**
1. Ensure the upgrade branch `upgrade-to-NET10` exists and working copy is on that branch (handle pending changes per repo policy — none presently).
2. Update the `<TargetFramework>` element to `net10.0-windows` in `HoursApp.csproj`.
3. Update `PackageReference` versions per §Package Update Reference.
4. Restore dependencies (dotnet restore) and build solution to identify compilation errors.
5. Address compilation errors caused by framework/API changes and package API changes.
6. Rebuild and verify solution builds with 0 errors.
7. Run unit/integration tests (if any) and confirm passing results.

**Expected Breaking Areas**
- Use of `System.Net.ServicePointManager` (source incompatible) — requires replacement with supported APIs (HttpClientHandler or SocketsHttpHandler configuration) where applicable.
- APIs deprecated by package upgrades or Microsoft.Identity.Client deprecation — review usages and replace or update packages to supported alternatives.

**Validation Checklist**
- [ ] `HoursApp.csproj` `TargetFramework` set to `net10.0-windows`
- [ ] All `PackageReference` versions updated per plan
- [ ] Solution restores and builds with 0 errors
- [ ] No critical runtime exceptions in smoke runs
- [ ] No package security vulnerabilities remain (addressed or deferred with rationale)

---

## Package Update Reference

### Common Package Updates (from assessment)

| Package | Current Version | Target Version | Projects Affected | Notes |
|---|---:|---:|---|---|
| Microsoft.Extensions.Configuration | 9.0.6 | 10.0.3 | HoursApp.csproj | Recommended upgrade for framework alignment |
| Microsoft.Extensions.Configuration.Json | 9.0.6 | 10.0.3 | HoursApp.csproj | Recommended upgrade |
| Microsoft.Extensions.Configuration.UserSecrets | 9.0.6 | 10.0.3 | HoursApp.csproj | Recommended upgrade |
| Microsoft.Extensions.Hosting | 9.0.6 | 10.0.3 | HoursApp.csproj | Recommended upgrade |
| Microsoft.Extensions.Logging | 9.0.6 | 10.0.3 | HoursApp.csproj | Recommended upgrade |
| EPPlus | 8.0.7 | (no change) | HoursApp.csproj | Compatible with net10.0 |
| MailKit | 4.13.0 | (no change) | HoursApp.csproj | Compatible |
| Microsoft.Graph | 5.82.0 | (no change) | HoursApp.csproj | Compatible |
| Microsoft.Identity.Client | 4.73.0 | (deprecated) | HoursApp.csproj | Marked deprecated — evaluate replacement or supported fork

Notes:
- Include all package updates listed above in the atomic upgrade. Do not skip updates flagged as "Recommended" in the assessment.
- For `Microsoft.Identity.Client` marked as deprecated, investigate supported alternatives or newer major versions before/after the atomic upgrade. If no direct replacement is available, document required code changes to migrate to the newer auth library.

---

## Breaking Changes Catalog

These are expected or likely items to encounter during build and code fixes.

1. T:System.Net.ServicePointManager — categorized as Source Incompatible in assessment. Code using `ServicePointManager` settings (e.g., `ServicePointManager.SecurityProtocol` or connection management) must be migrated to HttpClient patterns and `SocketsHttpHandler` settings. Replace usage with modern HttpClient configuration.

2. Package API changes for Microsoft.Extensions.* 9.x → 10.0.3 may require minor adapter changes in configuration/host initialization code (verify Program.cs/HostBuilder patterns).

3. `Microsoft.Identity.Client` deprecation — API surface may change if migrating to a replacement auth library; plan for authentication flow review.

Notes:
- The exact list of breaking changes will be finalized after the first build on net10.0; update this catalog with concrete compilation errors and suggested fixes.

---

## Testing & Validation Strategy

### Phase Validation
- Phase 0 (Preparation): confirm branch and prerequisites.
- Phase 1 (Atomic Upgrade): Apply project and package updates.
- Phase 2 (Build & Fix): Restore and build; fix compilation errors found.
- Phase 3 (Validation): Run all automated tests and smoke validations.

### Validation Checklist
- Solution builds with 0 errors
- No new high-severity warnings from analyzers (address or document)
- Unit tests pass (if present)
- Configuration and secrets loading behave as expected (verify `appsettings.json` and secrets handling)

---

## Risk Management

### Risk Summary
- Overall solution risk: **Low** (single small project, limited LOC)
- Noted risks:
  - `Microsoft.Identity.Client` deprecated — **Medium** risk if heavily used for auth flows
  - `ServicePointManager` usage — **Low/Medium** risk depending on extent of usage

### Mitigation
- Address `Microsoft.Identity.Client` deprecation by researching newest supported auth libraries and including migration notes. If replacement is large, plan for targeted follow-up changes.
- Replace `ServicePointManager` usages with HttpClient/SocketsHttpHandler in a single pass as code changes are identified during build.
- Keep a feature branch `upgrade-to-NET10` and a single cohesive PR for review to simplify rollback if needed.

---

## Complexity & Effort Assessment

- `HoursApp.csproj`: Complexity = **Low** (small LOC, limited package updates). Expected number of code edits: minimal — primarily API replacements for ServicePointManager and any deprecated auth library adjustments.

---

## Source Control Strategy

- Branch: create `upgrade-to-NET10` from `master` (recommended name returned by initialization).
- Pending changes: none detected. If pending changes exist, commit or stash per repo policy before branching.
- Commit strategy: group all upgrade changes in one atomic change (or a few logically grouped commits):
  1. Project file TargetFramework updates
  2. PackageReference updates
  3. Code fixes for compilation and API changes
- Create a single PR from `upgrade-to-NET10` to `master` describing the atomic upgrade, referencing this plan.
- Require code review and CI build before merging.

---

## Success Criteria

The migration is complete when the following are satisfied:
1. All projects target `net10.0` as specified in this plan.
2. All package updates listed in §Package Update Reference are applied.
3. Solution builds with 0 compilation errors.
4. Tests (if any) pass.
5. No outstanding known security vulnerabilities in packages remain unaddressed without documented rationale.

---

## Appendix

### Assessment Source
- `C:\Users\bmond\Documents\Toyota\toyota-time-tracking\.github\upgrades\scenarios\new-dotnet-version_ff1d79\assessment.md`


