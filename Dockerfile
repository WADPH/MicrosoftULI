FROM mcr.microsoft.com/powershell:7.4-alpine-3.17

WORKDIR /app

# Install dependencies
RUN apk add --no-cache curl bash busybox-suid ca-certificates libgdiplus

# Install Microsoft Graph module for Connect-MgGraph
RUN pwsh -NoProfile -Command "Install-PackageProvider -Name NuGet -Force -Scope AllUsers; Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop; Install-Module -Name Microsoft.Graph -Force -Scope AllUsers -AllowClobber -ErrorAction Stop; Install-Module -Name ImportExcel -Force -Scope AllUsers -AllowClobber -ErrorAction Stop"

# Copy script and configuration
COPY ULI.ps1 .
COPY .env .

# Create output directory
RUN mkdir -p /app/output

# Create cron file - runs daily at 22:00 UTC (2:00 AM local time in UTC+4 timezone)
RUN echo "0 2 * * * pwsh -File /app/ULI.ps1 >> /app/output/cron.log 2>&1" > /etc/crontabs/root

# Run cron in foreground
CMD ["crond", "-f", "-l", "8"]