# Public Group Visibility Auto-Remediation

Daily PowerShell automation that finds Microsoft 365 Groups with visibility set to Public and remediates them to Private by applying a sensitivity label to the associated site or group.  
All actions are logged to text files and uploaded to SharePoint. A Power Automate flow watches for new log files and posts key Public to Private actions to a Teams channel.

---

## Why this exists

Bank policy forbids public Groups or Teams. While sensitivity labels normally block public visibility, groups created via PowerShell or Graph can bypass auto-labeling. This job closes that gap.

---

## What it does

- Connects using an Azure App Registration – cert or secret
- Uses Microsoft Graph to enumerate all Unified groups
- Filters groups with visibility set to Public
- Resolves the site or group and applies the sensitivity label that enforces Private
- Writes logs to disk and uploads to SharePoint:
  - logs files – steps, errors, and Public to Private actions
  - report files – inventory of public groups found

---

## 🏗️ Architecture

```mermaid
flowchart TD
    A[Azure Automation - scheduled runbook/Task Schedule] --> B[Hybrid Worker/Azure VM]
    B --> PS[Public Group Remediation ps1]

    PS --> KV[Azure Key Vault - service account secrets]
    PS --> GCONN[Connect Microsoft Graph - app only]
    PS --> ENUM[Enumerate Unified groups]
    ENUM --> FILT{Visibility is Public?}
    FILT -- No --> SKIP[Skip]
    FILT -- Yes --> RESOLVE[Resolve site or group]
    RESOLVE --> APPLY[Apply sensitivity label to enforce Private]
    APPLY --> LOG1[Write action logs]
    ENUM --> LOG2[Write report of public groups]
    LOG1 --> UP[Upload logs to SharePoint library]
    LOG2 --> UP
    UP --> FLOW[Power Automate - on new file]
    FLOW --> TEAMS[Teams channel post]
