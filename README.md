# Calendar Tools for Exchange (EWS)

This repository contains two PowerShell scripts designed to analyze and maintain calendar items in Microsoft Exchange (Exchange 2016 / Exchange SE / Exchange Online via EWS API).

---

## Scripts Overview

### 1. **get-calendar-meetings-info.ps1**

A reporting utility that connects to Exchange Web Services (EWS) and exports detailed information about calendar meetings.

#### **Features**
- Connects to EWS using impersonation.
- Scans the Calendar folder with paging and filtering.
- Extracts meeting details:
  - Subject
  - Organizer
  - Start/End time
  - Creation time
  - Attachments flag
  - Appointment type (master/occurrence)
- Handles recurring meetings and retrieves master occurrences where necessary.
- Outputs results to a CSV file (`meetings.csv`).

#### **Usage**
```powershell
.\get-calendar-meetings-info.ps1 -Mailbox user@example.com -EwsUrl https://mail.example.com/EWS/Exchange.asmx
```

#### **Output**
A UTF‑8 CSV file containing all detected meetings over the configured date range.

---

### 2. **remove_attachements.ps1**  
*(calendar_cleaner.ps1)*

A universal tool for cleaning calendar attachments, intended for large-scale cleanup and troubleshooting Outlook calendar issues.

#### **Modes**
- **Report mode (default):**  
  Scans meetings and outputs a CSV report without making changes.
- **Clean mode:**  
  Removes attachments from meetings (including recurring events) through the organizer mailbox.

#### **Key Capabilities**
- Loads meeting objects and detects:
  - Whether a meeting contains attachments
  - Whether it is an occurrence or series master
- Removes attachments safely:
  - Avoids modifying items not owned by the organizer
  - Updates the event while preventing re-sending invites (SendToNone)
- Writes detailed processing logs to console
- Outputs a final report (always)

#### **Usage**
**Report only:**
```powershell
.emove_attachements.ps1 -Mailbox user@example.com -EwsUrl https://mail.example.com/EWS/Exchange.asmx -Report
```

**Clean attachments:**
```powershell
.emove_attachements.ps1 -Mailbox user@example.com -EwsUrl https://mail.example.com/EWS/Exchange.asmx
```

#### **Output**
- CSV report with:
  - Meeting UID
  - Subject
  - Start/End times
  - Attachments present before/after cleanup
- Console output showing progress and any detected issues

---

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- EWS Managed API (bundled with scripts via Reflection)
- Exchange Impersonation role:
  ```powershell
  New-ManagementRoleAssignment -Name "AllowImpersonation" -Role ApplicationImpersonation -User svc_ews
  ```

---

## Notes

- Scripts are optimized for environments with many calendar items and recurring-meeting complexities.
- The logic avoids modifying attendee calendars to prevent accidental re‑notifications.
- Intended for system administrators performing calendar auditing or cleanup.

---

## License

Internal / Private Use Only.

