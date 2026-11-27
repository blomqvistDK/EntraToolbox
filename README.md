# Entra Toolbox

The existing six-page GUI is backed by standalone, parameter-driven feature scripts. Keep the whole `EntraToolbox` folder together when moving or copying it. Internal feature and worker paths resolve relative to the launcher, not the current working directory.

```text
EntraToolbox/
|-- Invoke-EntraToolbox-GUI.ps1
|-- Features/
|   |-- Add-MembersToAADGroup/Add-MembersToEntraGroup.ps1
|   |-- Create-AADGroups/Create-EntraSecurityGroupsFromCsv.ps1
|   |-- Get-PrimaryPCFromEntraUsers/Get-PrimaryPCFromEntraUsers.ps1
|   |-- Get-DevicesLastUser.ps1
|   |-- ExtensionAttributes/Set-ExtensionAttribute.ps1
|   `-- Invoke-EntraDeviceCleanup.ps1
|-- Shared/Invoke-EntraToolboxOperation.ps1
|-- Tests/Test-EntraToolbox.ps1
`-- README.md
```

No bundled runtime templates or resources are required, so there is no `Resources` directory. Existing sample files, reports, unrelated tools, and dated backups remain outside this folder. The old launcher and active feature scripts were moved completely; update existing shortcuts and automation paths.

## PowerShell Modules

Install the Graph modules used by the toolbox features. Modules are loaded as needed; the GUI does not install them automatically.

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
```

`Microsoft.Graph.Reports` is additionally required when using Find Primary PCs with `-DeviceSource SignInLogs` or `-SignInFallback`:

```powershell
Install-Module Microsoft.Graph.Reports -Scope CurrentUser
```

The signed-in account still needs the required delegated Graph permissions and consent, Intune roles, and licensing for the selected operation. `Shared/Invoke-EntraToolboxOperation.ps1` is included in this folder and does not need to be installed separately.

## Launch

Run these examples from the toolbox root, or supply an absolute launcher path:

```powershell
powershell.exe -NoProfile -STA -File .\Invoke-EntraToolbox-GUI.ps1
pwsh.exe -NoProfile -STA -File .\Invoke-EntraToolbox-GUI.ps1 -TenantId 'tenant-guid'
```

Requires Windows and Windows PowerShell 5.1 or newer in STA mode. Install Microsoft.Graph.Authentication and the task-specific Microsoft.Graph.Groups, Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement, and Microsoft.Graph.Reports modules as needed. The GUI checks required modules without installing them automatically. Delegated sign-in, appropriate Graph consent, administrator roles, and licensing for the selected operation are still required.

On startup, the GUI loads `Microsoft.Graph.Authentication` and reads the local `Get-MgContext` once. An existing context immediately populates the connection indicator, account, tenant, and Sign out button; no context leaves the GUI disconnected. This check does not prompt for sign-in, request permissions, or make a Graph request. The indicator reflects context presence, not token validity or sufficient feature permissions; operations still check the required scopes and tenant. Connection-check failures are shown in the sidebar and activity log.

To reuse a live connection, launch the GUI from the same STA PowerShell process where you connected:

```powershell
Connect-MgGraph -Scopes 'User.Read.All'
.\Invoke-EntraToolbox-GUI.ps1
```

Starting `powershell.exe` or `pwsh.exe` creates a separate process and does not inherit another process's live Graph context. If an existing context belongs to a different tenant than `-TenantId`, the GUI shows the actual connection and a warning; operations remain blocked until you sign out and reconnect to the requested tenant.

`-ScriptsRoot` optionally selects an alternative feature directory with the same nested layout as `Features`. It does not change where the GUI loads its bundled `Shared` worker:

```powershell
pwsh.exe -NoProfile -STA -File .\Invoke-EntraToolbox-GUI.ps1 -ScriptsRoot 'D:\EntraFeatures'
```

## Feature Inventory

The GUI's left menu groups features under section headings in this order, and opens on the first item (Find Primary PCs).

| Menu section | Feature | Standalone script | Behavior |
| --- | --- | --- | --- |
| User-Device Affinity | Find Primary PCs | [Get-PrimaryPCFromEntraUsers.ps1](Features/Get-PrimaryPCFromEntraUsers/Get-PrimaryPCFromEntraUsers.ps1) | Read-only Entra Latest/All or sign-in-only Most/All Windows-PC discovery and CSV export |
| User-Device Affinity | Device Last User(s) | [Get-DevicesLastUser.ps1](Features/Get-DevicesLastUser.ps1) | Read-only Intune Latest/Most/All logged-on user report and CSV export |
| Groups | Add Group Members | [Add-MembersToEntraGroup.ps1](Features/Add-MembersToAADGroup/Add-MembersToEntraGroup.ps1) | Adds users, devices, or groups; supports WhatIf |
| Groups | Create Security Groups | [Create-EntraSecurityGroupsFromCsv.ps1](Features/Create-AADGroups/Create-EntraSecurityGroupsFromCsv.ps1) | Creates/reuses groups and adds optional nested memberships |
| Users | User Extension Attributes | [Set-ExtensionAttribute.ps1](Features/ExtensionAttributes/Set-ExtensionAttribute.ps1) | Sets/clears extension attributes; supports WhatIf |
| Devices | Device Cleanup | [Invoke-EntraDeviceCleanup.ps1](Features/Invoke-EntraDeviceCleanup.ps1) | Inspects devices from a CSV, an Entra stale-device lookup or typed names, then disables or deletes selected cloud-device records |

`Shared/Invoke-EntraToolboxOperation.ps1` is the GUI's internal background runner for authentication, tenant verification, module loading, and result queuing. Feature scripts do not need to load the GUI or this runner to execute independently.

## Standalone Examples

Supply the required inputs to avoid interactive file selection. All six features support `-PassThru` for structured results and `-SkipGraphConnect` to require an existing session with the necessary scopes. These switches do not grant permissions or make a write operation read-only. Where supported, `-NonInteractive` disables optional console prompts. Authentication may still be interactive without `-SkipGraphConnect`.

Use `Get-Help` on a feature script for its complete parameters and CSV schema. Input paths and explicit output destinations remain caller-supplied; omitted output paths retain their existing defaults, including working-directory-based CSV exports.

### Add Group Members

```powershell
.\Features\Add-MembersToAADGroup\Add-MembersToEntraGroup.ps1 -TargetGroupName 'Pilot Users' -MemberType Users -MemberNames 'user@example.com' -WhatIf -PassThru
```

Use `-MemberFile` instead of `-MemberNames` for file input. Remove `-WhatIf` only when ready to add memberships.

### Create Security Groups

```powershell
.\Features\Create-AADGroups\Create-EntraSecurityGroupsFromCsv.ps1 -CsvPath 'C:\Inputs\groups.csv' -PassThru
```

This command performs writes. It does not offer a WhatIf preview. The input is a semicolon-delimited CSV with `GroupName`, `Description`, and optional `MemberGroup` values. It retains existing group reuse and nested-membership behavior.

### Find Primary PCs

```powershell
.\Features\Get-PrimaryPCFromEntraUsers\Get-PrimaryPCFromEntraUsers.ps1 -UserPrincipalName 'user@example.com' -DeviceMode Latest -OutputPath 'C:\Reports\primary-pcs.csv' -NonInteractive -PassThru
.\Features\Get-PrimaryPCFromEntraUsers\Get-PrimaryPCFromEntraUsers.ps1 -GroupName 'Pilot Users' -DeviceMode All -OutputPath 'C:\Reports\all-pcs.csv' -NonInteractive -PassThru
.\Features\Get-PrimaryPCFromEntraUsers\Get-PrimaryPCFromEntraUsers.ps1 -UserPrincipalName 'user@example.com' -DeviceSource SignInLogs -DeviceMode Most -SignInDays 30 -OutputPath 'C:\Reports\most-used-pcs.csv' -NonInteractive -PassThru
.\Features\Get-PrimaryPCFromEntraUsers\Get-PrimaryPCFromEntraUsers.ps1 -CsvPath 'C:\Inputs\users.csv' -DeviceSource SignInLogs -DeviceMode All -SignInDays 14 -OutputPath 'C:\Reports\logged-on-pcs.csv' -NonInteractive -PassThru
```

Also accepts `-GroupId` and `-CsvPath`. The **User source** / input CSV selects users using `UserPrincipalName`; the separate **Entra import CSV output** contains unique discovered **device object IDs**, retaining the existing bulk-import header. GUI results and **Export results** retain the per-user/device report columns.

The GUI's **Device discovery** selects:

- **Entra owned/registered devices** (default): existing **Latest** active Intune-managed Windows PC or **All** associated Windows devices, including unmanaged devices. Optional `-SignInFallback` still returns one device only when no associated device qualifies; its existing event-counting behavior is unchanged.
- **Sign-in logs only** (`-DeviceSource SignInLogs`): ignores owned/registered relationships and counts only successful **Windows Sign In** application events with a device ID, within **Sign-in lookback days** (`-SignInDays`, default 30; allowed 1-365).
  - **Most logged-on PCs** (`-DeviceMode Most`, default for this source) returns every eligible Windows PC tied for the highest event count.
  - **All logged-on PCs** (`-DeviceMode All`) returns each eligible Windows PC once per user, ordered by count descending, then newest successful sign-in, then device ID. Tied Most results use the same ordering.

Sign-in-only discovery resolves device registration IDs from logs against Entra for Windows validation, names and exportable object IDs. Missing/ambiguous devices and lookup failures are reported as warnings and skipped. The optional **Only devices in Intune** filter (`-OnlyIntuneDevices`) applies before selecting winners and requires `DeviceManagementManagedDevices.Read.All`. Without that filter, sign-in-only discovery can include unmanaged Windows PCs. Query failures and empty results never fall back to device ownership or console prompts.

The report's `Source` includes each device's event count. In sign-in-only mode, `LastActivity` is that user's newest successful Windows Sign In event on the selected PC, rather than the device-wide approximate activity used by Entra discovery. The result columns and unique device-ID output CSV format are unchanged.

Both sign-in-only discovery and optional fallback require `Microsoft.Graph.Reports`, delegated `AuditLog.Read.All`, appropriate access/consent and applicable Entra ID P1/P2 licensing. The signed-in account also needs an active Entra role that can read sign-in logs, such as Reports Reader, Global Reader or Security Reader. Before the first sign-in lookup, the script probes sign-in log access once. If Graph reports a missing role or P1/P2 license, then during the probe or the first lookup, one warning is shown and remaining sign-in lookups are skipped instead of failing for every user. Available history is limited by tenant log retention; increasing the lookback does not extend retention. These are observed Entra Windows Sign In events, **not a complete history of local Windows logons**. Sign-in-only discovery cannot be combined with `-SignInFallback` or `-DeviceMode Latest`; Entra discovery supports `Latest|All`, not `Most`.

### Device Last User(s)

```powershell
.\Features\Get-DevicesLastUser.ps1 -DeviceName 'PC1','PC2' -OutputPath 'C:\Reports\last-users.csv' -PassThru
.\Features\Get-DevicesLastUser.ps1 -GroupNameOrId 'Pilot Windows Devices' -OutputPath 'C:\Reports\last-users.csv' -PassThru
.\Features\Get-DevicesLastUser.ps1 -DeviceName 'PC1' -UserMode Most -OutputPath 'C:\Reports\most-users.csv' -PassThru
.\Features\Get-DevicesLastUser.ps1 -DeviceName 'PC1' -UserMode All -OutputPath 'C:\Reports\all-users.csv' -PassThru
.\Features\Get-DevicesLastUser.ps1 -DeviceName 'PC1' -ExtensionAttributeNumbers 1,5,15 -OutputPath 'C:\Reports\user-attributes.csv' -PassThru
```

The GUI's **User selection** maps to `-UserMode Latest|Most|All`:

- **Latest logged-on user** (default): the usable `usersLoggedOn` entry with the newest `lastLogOnDateTime`.
- **Most logged-on user(s)**: count each `userId` occurrence in `usersLoggedOn` and return every user tied for the highest count. This measures only the returned array, not complete historical sign-in frequency.
- **All logged-on users**: each distinct `userId` once per device, ordered by that user's newest logon, newest first. Most-mode ties use the same ordering.

Results and CSV exports retain the existing columns, with one row per selected device/user. `LastLoggedOnUserId` identifies that row's selected logged-on user, not necessarily the device's latest user in Most/All modes. Blank user IDs are ignored; all modes retain the primary-user fallback and missing-user reporting when no usable logged-on ID exists. A device with one usable entry produces the same result in every mode. GUI run-summary counts reflect result rows, which can exceed the number of devices.

Use the **User extension attribute columns (optional)** dropdown to check any combination of
`extensionAttribute1` through `extensionAttribute15` before running. The dropdown
stays open while you check several attributes and shows the selection when closed.
None are selected by default; **Reset** clears all selections. Selections apply to the
next run, not existing results, and are not saved across launches. The selected
columns are appended in numeric order to the output grid, report CSV, and
**Export results** CSV. With none selected, the existing report schema is unchanged.

The standalone equivalent is `-ExtensionAttributeNumbers 1,5,15` (each number must
be 1-15; duplicates are removed). Values come from the resolved user's Microsoft
Entra `onPremisesExtensionAttributes`, including when the report falls back to
the primary user. Unset attributes and missing users produce blank cells.
Profile/attribute lookup failures retain the row and core identity when available,
with details in `LookupError`, `LookupPhase` and the activity warnings. The
existing `User.Read.All` permission covers this read-only retrieval; this does not
modify user attributes or include other custom/directory extension types.

Use `-CsvPath` instead of device names for CSV input. `-GroupNameOrId` accepts a group object ID or an exact, unique display name. It includes direct device members that also have an Intune managed-device record; it does not expand user members or nested groups. Group input additionally requires the delegated `Device.Read.All`, `Group.Read.All` and `GroupMember.Read.All` permissions. The standalone script's existing no-input behavior scans all managed devices; the GUI continues to require explicit selection of that input mode.

### User Extension Attributes

```powershell
.\Features\ExtensionAttributes\Set-ExtensionAttribute.ps1 -UserPrincipalName 'user@example.com' -AttributeNumber 5 -Value 'Project' -NonInteractive -WhatIf -PassThru
.\Features\ExtensionAttributes\Set-ExtensionAttribute.ps1 -UserPrincipalName 'user@example.com' -AttributeNumber 5 -Clear -NonInteractive -WhatIf -PassThru
```

`-CsvPath` supports multiple users. Remove `-WhatIf` only when ready to set or clear values.

### Device Cleanup

```powershell
$devices = .\Features\Invoke-EntraDeviceCleanup.ps1 -CsvPath 'C:\Inputs\devices.csv' -PassThru
.\Features\Invoke-EntraDeviceCleanup.ps1 -Action Delete -Device $devices[0] -ExpectedTenantId $devices[0].TenantId -AcknowledgeCloudDeletion -WhatIf -PassThru
# Keep Autopilot registration while deleting Intune and Entra records
.\Features\Invoke-EntraDeviceCleanup.ps1 -Action Delete -Device $devices[0] -ExpectedTenantId $devices[0].TenantId -AcknowledgeCloudDeletion -DeleteAutopilotRegistration:$false -WhatIf -PassThru

# Stale Windows devices straight from Entra, or exact device names
$stale = .\Features\Invoke-EntraDeviceCleanup.ps1 -Source Entra -InactivityDays 90 -PassThru
$named = .\Features\Invoke-EntraDeviceCleanup.ps1 -Source Names -DeviceName 'PC-001','PC-002' -PassThru
.\Features\Invoke-EntraDeviceCleanup.ps1 -Action Disable -Device $stale -ExpectedTenantId $stale[0].TenantId -WhatIf -PassThru
```

The first command inspects devices. The second previews deletion for one inspected device, including Autopilot deregistration. The third keeps Autopilot registration. Before a real deletion, verify eligibility, tenant, administrator roles and required deletion scopes. `-SkipGraphConnect` can reuse a suitably consented session. Remove `-WhatIf` only for an intentional deletion; normal confirmation remains supported.

**Delete Autopilot registration** is unchecked by default in the GUI. The CLI parameter `-DeleteAutopilotRegistration` defaults to `$true`, preserving the previous script behavior:

- **Unchecked / `-DeleteAutopilotRegistration:$false`:** delete and confirm linked Intune records, then delete and confirm the remaining Entra object. Autopilot deregistration requests and the ten-second handoff delay are skipped.
- **Checked / CLI default:** delete and confirm Intune records, request deregistration of linked Autopilot registrations, wait ten seconds once after accepted requests, then delete and confirm any remaining Entra object. Autopilot deregistration completion is not verified.

Intune cleanup is unchanged in both modes.

**Warning:** [Windows Autopilot depends on its associated Entra object](https://learn.microsoft.com/en-us/autopilot/registration-overview). Deleting that object while keeping registration can cause future enrollment failures. The confirmation and Activity warn when keeping linked registrations; keeping registration is not a guarantee of successful redeployment. Entra deletion errors are reported without falling back to Autopilot deregistration.

Both modes still read Intune and Autopilot inventory and validate linkage. Real deletion requires `Directory.AccessAsUser.All` and `DeviceManagementManagedDevices.ReadWrite.All`; `DeviceManagementServiceConfig.ReadWrite.All` is additionally required only when deregistration is selected. The inventory read scopes remain required in both modes, and preview uses read permissions only.

Device sources (GUI **Device source** / `-Source`):

- **CSV file** (`Csv`, default): devices listed in a CSV, as before.
- **Entra lookup** (`Entra`): all Windows devices whose latest activity (Entra approximate sign-in or Intune sync, whichever is newer) is older than `-InactivityDays` (default 90), plus devices that have never reported activity. Autopilot registration is not required; the **Autopilot** column shows it.
- **Device names** (`Names`): exact, case-insensitive display names, one per line. Every Entra device with a matching name is listed; unknown names are reported as not found.

**Disable selected** (`-Action Disable`) only sets the Entra device `accountEnabled` to false, then re-reads the device to confirm it. Intune and Autopilot records are untouched, already-disabled devices are skipped, and hybrid or unknown join types are blocked. It honours **Preview only**, asks for tenant confirmation, and needs `Directory.AccessAsUser.All` plus Cloud Device Administrator (or higher); the deletion acknowledgement is not required. Deletion does not require the device to be disabled first. A common flow is to load stale devices, disable them, and delete them after a grace period.

Cleanup retains its verification rules, hybrid-device restrictions, and cloud-deletion acknowledgement. Selecting Autopilot deregistration preserves the former three-stage deletion sequence. It does not wipe devices or delete on-premises AD objects. The GUI retains preview-by-default, selection filtering, acknowledgement, tenant checks, and confirmation safeguards.

## Offline Validation

```powershell
pwsh.exe -NoProfile -STA -File .\Tests\Test-EntraToolbox.ps1
powershell.exe -NoProfile -STA -File .\Tests\Test-EntraToolbox.ps1
```

The suite uses mocked Graph commands, checks CLI results and CSV formats, exercises cleanup safeguards and background-worker behavior, and validates GUI layout/theme. It also copies the package to a temporary path containing spaces, launches it from another working directory, and checks an explicit `-ScriptsRoot` override. Test windows close automatically and temporary data is removed. No live tenant operations are performed.

To construct the interface without showing a window, importing Graph, or probing the connection, use `-ValidateOnly`. The test suite also validates this entry point and mocked startup connection detection. Production Graph workflows require separate testing with an authorized tenant.