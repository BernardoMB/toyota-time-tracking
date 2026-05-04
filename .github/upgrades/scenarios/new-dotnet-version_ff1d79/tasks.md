# HoursApp .NET 10 Upgrade Tasks

## Overview

This document tracks the execution of the upgrade of `HoursApp.sln` to .NET 10.0. The upgrade will update the project TargetFramework and NuGet package versions in a single coordinated pass, then build and validate the solution.

**Progress**: 4/4 tasks complete (100%) ![0%](https://progress-bar.xyz/100)

---

## Tasks

### [✓] TASK-001: Verify prerequisites *(Completed: 2026-03-08 16:23)*
**References**: Plan §Migration Strategy, Plan §Implementation Notes

- [✓] (1) Verify required .NET SDK (net10.0) is installed per Plan §Implementation Notes
- [✓] (2) Runtime/SDK version meets minimum requirements (**Verify**)
- [✓] (3) If a `global.json` exists, verify compatibility or update it per Plan §Migration Strategy
- [✓] (4) Verify configuration files and build tools are compatible with target framework (**Verify**)

### [✓] TASK-002: Atomic framework and package upgrade with compilation fixes *(Completed: 2026-03-08 16:24)*
**References**: Plan §Project-by-Project Plans, Plan §Package Update Reference, Plan §Breaking Changes Catalog

- [✓] (1) Update `HoursApp.csproj` `<TargetFramework>` to `net10.0-windows` per Plan §Project-by-Project Plans
- [✓] (2) Update all `PackageReference` versions per Plan §Package Update Reference (do not duplicate list)
- [✓] (3) Restore dependencies (e.g., `dotnet restore`) per Plan §Migration Strategy
- [✓] (4) Build solution and fix all compilation errors caused by framework/package upgrades (reference Plan §Breaking Changes Catalog for common issues)
- [✓] (5) Solution builds with 0 errors (**Verify**)

### [✓] TASK-003: Run test suite and validate upgrade *(Completed: 2026-03-08 12:24)*
**References**: Plan §Testing & Validation Strategy, Plan §Project-by-Project Plans

- [✓] (1) Run all test projects for the solution per Plan §Testing & Validation Strategy (if none exist, this step completes with no tests)
- [✓] (2) Fix any test failures (reference Plan §Breaking Changes Catalog for guidance)
- [✓] (3) Re-run tests after fixes
- [✓] (4) All tests pass with 0 failures (**Verify**)

### [✓] TASK-004: Final commit *(Completed: 2026-03-08 16:26)*
**References**: Plan §Source Control Strategy

- [✓] (1) Commit all remaining changes with message: "TASK-004: Complete upgrade to .NET 10.0 (net10.0)"


