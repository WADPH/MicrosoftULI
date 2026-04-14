FROM mcr.microsoft.com/powershell:7.4-alpine

WORKDIR /app

# Install cron
RUN apk add --no-cache curl bash busybox-suid

# Copy script
COPY ULI.ps1 .

# Create cron file
RUN echo "0 22 * * * pwsh -File /app/ULI.ps1 >> /app/cron.log 2>&1" > /etc/crontabs/root

# Run cron in foreground
CMD ["crond", "-f", "-l", "8"]