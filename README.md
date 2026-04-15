# MicrosoftULI

PowerShell script that aggregates user and license data from multiple Microsoft Entra ID tenants, generates a structured Excel report, and uploads it to SharePoint.

---

## Features

* Collects users from multiple Entra ID tenants
* Filters out guest accounts (`#EXT#`)
* Maps license SKUs to human-readable names
* Generates Excel report with:

  * Per-user license breakdown
  * Dynamic totals (filter-friendly)
  * License usage statistics
  * Cost calculations (monthly / yearly)
* Creates separate **Summary** sheet with aggregated metrics
* Uploads generated file to SharePoint via Microsoft Graph
* Logging support

---

## Configuration

Configuration is provided via environment variables.

1. Copy example file:

```bash
cp env.example .env
```

2. Fill in your values inside `.env`.

### Required variables

```bash
CONFIG_JSON={"OutputPath":".","FileName":"your_filename.xlsx","LogPath":"your_log_file.log","SummarySheetName":"Summary","MainSheetName":"Users","SharePoint":{"SiteUrl":"your_site.sharepoint.com","Library":"Documents","Folder":"Your/Folder/"},"Tenants":[{"Name":"Tenant 1","TenantId":"your-tenant-id-1","ClientId":"your-client-id-1","ClientSecret":"your-client-secret-1"},{"Name":"Tenant 2","TenantId":"your-tenant-id-2","ClientId":"your-client-id-2","ClientSecret":"your-client-secret-2"}]}

LICENSE_PRICES_JSON={"License Name 1":0.00,"License Name 2":0.00}
```

* `CONFIG_JSON` — main configuration
* `LICENSE_PRICES_JSON` — used for cost calculations in Excel

---

## Docker Support

Project can be run in Docker:

```bash
docker compose up
```

The container runs the script on a schedule using cron.

---

## Change Execution Schedule

To modify how often the script uploads the file to SharePoint, update the cron expression in the `Dockerfile`:

```dockerfile
RUN echo "0 2 * * * pwsh -File /app/ULI.ps1 >> /app/output/cron.log 2>&1" > /etc/crontabs/root
```

Example above runs the script daily at **02:00**.

---

## Output

* Excel file with:

  * User list
  * License matrix
  * Totals and cost calculations
* Summary sheet with aggregated statistics
* Log file with execution details

---

## Requirements

* PowerShell 7+
* Microsoft Graph PowerShell SDK
* Permissions to (Entra ID app):

  * Read users and licenses
  * Upload files to SharePoint

---