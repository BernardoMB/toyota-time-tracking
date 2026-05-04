# Projects and dependencies analysis

This document provides a comprehensive overview of the projects and their dependencies in the context of upgrading to .NETCoreApp,Version=v10.0.

## Table of Contents

- [Executive Summary](#executive-Summary)
  - [Highlevel Metrics](#highlevel-metrics)
  - [Projects Compatibility](#projects-compatibility)
  - [Package Compatibility](#package-compatibility)
  - [API Compatibility](#api-compatibility)
- [Aggregate NuGet packages details](#aggregate-nuget-packages-details)
- [Top API Migration Challenges](#top-api-migration-challenges)
  - [Technologies and Features](#technologies-and-features)
  - [Most Frequent API Issues](#most-frequent-api-issues)
- [Projects Relationship Graph](#projects-relationship-graph)
- [Project Details](#project-details)

  - [HoursApp.csproj](#hoursappcsproj)


## Executive Summary

### Highlevel Metrics

| Metric | Count | Status |
| :--- | :---: | :--- |
| Total Projects | 1 | All require upgrade |
| Total NuGet Packages | 9 | 6 need upgrade |
| Total Code Files | 4 |  |
| Total Code Files with Incidents | 2 |  |
| Total Lines of Code | 382 |  |
| Total Number of Issues | 8 |  |
| Estimated LOC to modify | 1+ | at least 0.3% of codebase |

### Projects Compatibility

| Project | Target Framework | Difficulty | Package Issues | API Issues | Est. LOC Impact | Description |
| :--- | :---: | :---: | :---: | :---: | :---: | :--- |
| [HoursApp.csproj](#hoursappcsproj) | net8.0-windows | 🟢 Low | 6 | 1 | 1+ | DotNetCoreApp, Sdk Style = True |

### Package Compatibility

| Status | Count | Percentage |
| :--- | :---: | :---: |
| ✅ Compatible | 3 | 33.3% |
| ⚠️ Incompatible | 1 | 11.1% |
| 🔄 Upgrade Recommended | 5 | 55.6% |
| ***Total NuGet Packages*** | ***9*** | ***100%*** |

### API Compatibility

| Category | Count | Impact |
| :--- | :---: | :--- |
| 🔴 Binary Incompatible | 0 | High - Require code changes |
| 🟡 Source Incompatible | 1 | Medium - Needs re-compilation and potential conflicting API error fixing |
| 🔵 Behavioral change | 0 | Low - Behavioral changes that may require testing at runtime |
| ✅ Compatible | 545 |  |
| ***Total APIs Analyzed*** | ***546*** |  |

## Aggregate NuGet packages details

| Package | Current Version | Suggested Version | Projects | Description |
| :--- | :---: | :---: | :--- | :--- |
| EPPlus | 8.0.7 |  | [HoursApp.csproj](#hoursappcsproj) | ✅Compatible |
| MailKit | 4.13.0 |  | [HoursApp.csproj](#hoursappcsproj) | ✅Compatible |
| Microsoft.Extensions.Configuration | 9.0.6 | 10.0.3 | [HoursApp.csproj](#hoursappcsproj) | NuGet package upgrade is recommended |
| Microsoft.Extensions.Configuration.Json | 9.0.6 | 10.0.3 | [HoursApp.csproj](#hoursappcsproj) | NuGet package upgrade is recommended |
| Microsoft.Extensions.Configuration.UserSecrets | 9.0.6 | 10.0.3 | [HoursApp.csproj](#hoursappcsproj) | NuGet package upgrade is recommended |
| Microsoft.Extensions.Hosting | 9.0.6 | 10.0.3 | [HoursApp.csproj](#hoursappcsproj) | NuGet package upgrade is recommended |
| Microsoft.Extensions.Logging | 9.0.6 | 10.0.3 | [HoursApp.csproj](#hoursappcsproj) | NuGet package upgrade is recommended |
| Microsoft.Graph | 5.82.0 |  | [HoursApp.csproj](#hoursappcsproj) | ✅Compatible |
| Microsoft.Identity.Client | 4.73.0 |  | [HoursApp.csproj](#hoursappcsproj) | ⚠️NuGet package is deprecated |

## Top API Migration Challenges

### Technologies and Features

| Technology | Issues | Percentage | Migration Path |
| :--- | :---: | :---: | :--- |

### Most Frequent API Issues

| API | Count | Percentage | Category |
| :--- | :---: | :---: | :--- |
| T:System.Net.ServicePointManager | 1 | 100.0% | Source Incompatible |

## Projects Relationship Graph

Legend:
📦 SDK-style project
⚙️ Classic project

```mermaid
flowchart LR
    P1["<b>📦&nbsp;HoursApp.csproj</b><br/><small>net8.0-windows</small>"]
    click P1 "#hoursappcsproj"

```

## Project Details

<a id="hoursappcsproj"></a>
### HoursApp.csproj

#### Project Info

- **Current Target Framework:** net8.0-windows
- **Proposed Target Framework:** net10.0--windows
- **SDK-style**: True
- **Project Kind:** DotNetCoreApp
- **Dependencies**: 0
- **Dependants**: 0
- **Number of Files**: 4
- **Number of Files with Incidents**: 2
- **Lines of Code**: 382
- **Estimated LOC to modify**: 1+ (at least 0.3% of the project)

#### Dependency Graph

Legend:
📦 SDK-style project
⚙️ Classic project

```mermaid
flowchart TB
    subgraph current["HoursApp.csproj"]
        MAIN["<b>📦&nbsp;HoursApp.csproj</b><br/><small>net8.0-windows</small>"]
        click MAIN "#hoursappcsproj"
    end

```

### API Compatibility

| Category | Count | Impact |
| :--- | :---: | :--- |
| 🔴 Binary Incompatible | 0 | High - Require code changes |
| 🟡 Source Incompatible | 1 | Medium - Needs re-compilation and potential conflicting API error fixing |
| 🔵 Behavioral change | 0 | Low - Behavioral changes that may require testing at runtime |
| ✅ Compatible | 545 |  |
| ***Total APIs Analyzed*** | ***546*** |  |

