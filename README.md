# Exchange-scripts
Repository of Exchange scripts.

## Repository richtlijnen
- Zie [`AGENTS.md`](AGENTS.md) voor code- en structuurafspraken voor nieuwe of aangepaste scripts.
- Groepeer nieuwe scripts per domein (bijv. `exchange/`, `azure-ad/`) en documenteer ze in een lokale `README.md`.
- Deelbare functies plaats je in modules onder `modules/<ModuleNaam>/<ModuleNaam>.psm1` en documenteer exports.

## Find-GhostRoomMeetings.ps1
Audits room mailbox calendars to find meetings whose organizers are missing, disabled, or external. The script connects to Exchange via remote PowerShell and EWS to gather calendar data, validate organizer status, and produce a report.

### Key parameters
- **ExchangeUri**: Remote PowerShell endpoint for Exchange.
- **Credential**: Credentials with permissions to query mailboxes.
- **EwsAssemblyPath**: Path to the EWS Managed API assembly.
- **MonthsAhead / MonthsBehind**: Date window for scanning meetings.
- **OutputPath**: CSV report path (default: `ghost-meetings-report.csv` in the current directory).
- **ExcelOutputPath**: Optional Excel report path. Provide a path to generate an `.xlsx` export; otherwise only the CSV is created. Requires the `ImportExcel` PowerShell module.
- **OrganizationSmtpSuffix**: Domain suffix used to identify internal organizers.
- **ImpersonationSmtp**: SMTP address used for EWS impersonation and Autodiscover.
- **SendInquiry / NotificationFrom / NotificationTemplate**: Options to email attendees of ghost meetings.

### Exporting to Excel
When **ExcelOutputPath** is provided, the script attempts to load the `ImportExcel` module and writes the report to an `.xlsx` file alongside the CSV. Install the module beforehand with `Install-Module ImportExcel` if it is not already available.
