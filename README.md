# exo-archiver

Interactive PowerShell script to enable and configure the In-Place Archive on Exchange Online, set a retention policy and immediately trigger archiving — no manual cmdlets needed.

---

## Requirements

- PowerShell 5.1 or later
- `ExchangeOnlineManagement` module installed
- Minimum roles: **Mail Recipients** + **Retention Management**

If the module is missing, the script detects it and shows the install command:

```powershell
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
```

---

## Usage

```powershell
.\Enable-ExArchive_Custom.ps1
```

Fully interactive — no parameters required.

---

## What the script does

```
[1] Checks that the ExchangeOnlineManagement module is installed
[2] Opens Microsoft OAuth2 login (browser popup)
[3] Enables Auto-Expanding Archive at organization level
[4] Asks for the target user (UPN or display name)
[5] Lets you choose the archiving threshold
[6] Shows a full summary and asks for confirmation before proceeding
[7] Executes all required operations
[8] Verifies the final state and disconnects
```

### Operations performed automatically

- Enables the In-Place Archive on the mailbox (if not already active)
- Enables Auto-Expanding Archive on the mailbox
- Creates the `MoveToArchive` Retention Tag with the chosen threshold
- Creates the Retention Policy and links the tag
- Assigns the policy to the mailbox
- Forces an immediate run of the Managed Folder Assistant

> **Note:** If a Retention Tag or Policy already exists in the tenant it is reused, not duplicated. At most 6 fixed policies will exist in the tenant (one per threshold), shared across all mailboxes configured with the same threshold.

---

## Available archiving thresholds

| Choice | Threshold | Days |
|--------|-----------|------|
| `1` | 1 year | 365 |
| `2` | 2 years | 730 |
| `3` | 3 years | 1095 |
| `4` | 4 years | 1460 |
| `5` | 5 years | 1825 |
| `6` | 6 years | 2190 |
| `7` | Custom | `years × 365` |

Option `7` accepts any integer from 1 to 99 years and calculates the days automatically.

---

## License compatibility

| License | In-Place Archive | Base storage | Native auto-expanding | With add-on |
|---|---|---|---|---|
| Exchange Online Plan 1 | ✅ YES | ~50 GB | ❌ NO | ✅ 100 GB + auto-expanding (~1.5 TB) |
| Exchange Online Plan 2 | ✅ YES | 100 GB | ✅ YES | — |
| M365 Business Basic | ✅ YES | ~50 GB | ❌ NO | ✅ 100 GB + auto-expanding |
| M365 Business Standard | ✅ YES | ~50 GB | ❌ NO | ✅ 100 GB + auto-expanding |
| M365 Business Premium | ✅ YES | 100 GB | ✅ YES | — |
| M365 E1 | ✅ YES | ~50 GB | ❌ NO | ✅ 100 GB + auto-expanding |
| M365 E3 | ✅ YES | 100 GB | ✅ YES | — |
| M365 E5 | ✅ YES | 100 GB | ✅ YES | — |
| Office 365 E1 | ✅ YES | ~50 GB | ❌ NO | ✅ 100 GB + auto-expanding |
| Office 365 E3 | ✅ YES | 100 GB | ✅ YES | — |
| Office 365 E5 | ✅ YES | 100 GB | ✅ YES | — |

Licenses without native auto-expanding require the **Exchange Online Archiving** add-on (additional Microsoft license assigned to the account).

---

## Monitoring progress

After running the script, the Managed Folder Assistant works in the background. To monitor progress:

**Primary mailbox item count** — if it drops, archiving is working:
```powershell
Get-MailboxStatistics user@domain.com | fl ItemCount,TotalItemSize
```

**Archive contents — non-empty folders only** (recommended):
```powershell
Get-MailboxFolderStatistics user@domain.com -Archive | where {$_.ItemsInFolder -gt 0} | ft Name,ItemsInFolder,FolderSize
```

**Archive contents — all folders:**
```powershell
Get-MailboxFolderStatistics user@domain.com -Archive | ft Name,ItemsInFolder,FolderSize
```

---

## Execution times

| Scenario | Expected time |
|---|---|
| First run, standard mailbox | 1–6 hours |
| Large mailboxes (50 GB+) | up to 24 hours |
| Subsequent runs | automatic Microsoft cycles |

The process is entirely **server-side** — Outlook or OWA do not need to be open.

---

## Notes

- Archiving is **one-way** — mail moved to the archive does not return to the inbox if the policy is changed
- Changing the threshold on an already configured mailbox only affects future archiving; mail already archived remains untouched
- There is no visible MRM queue — the Managed Folder Assistant runs in silent batches

---

## References

- [Microsoft Docs — Enable archive mailboxes](https://learn.microsoft.com/en-us/purview/enable-archive-mailboxes)
- [Microsoft Docs — Retention tags and retention policies](https://learn.microsoft.com/en-us/exchange/security-and-compliance/messaging-records-management/retention-tags-and-policies)
- [Microsoft Docs — Start-ManagedFolderAssistant](https://learn.microsoft.com/en-us/powershell/module/exchange/start-managedfolderassistant)
