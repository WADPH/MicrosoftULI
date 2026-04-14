# Sync-EntraUsersToSharePoint.ps1
# Объединенный скрипт: получение данных из нескольких Entra ID и загрузка в SharePoint

# === КОНФИГУРАЦИЯ ===
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry
    if ($logFilePath) {
        Add-Content -Path $logFilePath -Value $logEntry -ErrorAction SilentlyContinue
    }
}

# Load environment variables from .env file
if (Test-Path ".env") {
    Get-Content ".env" | ForEach-Object {
        $line = $_.Trim()
        if ($line -and -not ($line -match '^\s*#')) {
            $parts = $line -split '=', 2
            if ($parts.Count -eq 2) {
                $key = $parts[0].Trim()
                $value = $parts[1].Trim()
                [Environment]::SetEnvironmentVariable($key, $value)
            }
        }
    }
    Write-Log "Configuration loaded from: .env" "INFO"
} else {
    Write-Log "Warning: .env file not found. Using environment variables." "WARNING"
}

# Helper function to convert PSCustomObject to Hashtable
function ConvertTo-Hashtable {
    param(
        [Parameter(ValueFromPipeline = $true)]
        $InputObject
    )
    
    if ($InputObject -is [System.Collections.Hashtable]) {
        return $InputObject
    }
    
    $result = @{}
    $InputObject.PSObject.Properties | ForEach-Object {
        if ($_.Value -is [PSCustomObject]) {
            $result[$_.Name] = ConvertTo-Hashtable -InputObject $_.Value
        } elseif ($_.Value -is [Object[]] -and $_.Value.Count -gt 0 -and $_.Value[0] -is [PSCustomObject]) {
            $result[$_.Name] = @($_.Value | ForEach-Object { ConvertTo-Hashtable -InputObject $_ })
        } else {
            $result[$_.Name] = $_.Value
        }
    }
    return $result
}

# Load configuration from environment
$configJson = ConvertFrom-Json $env:CONFIG_JSON
$config = ConvertTo-Hashtable -InputObject $configJson

# Убеждаемся, что директория существует
if (-not (Test-Path $config.OutputPath)) {
    New-Item -ItemType Directory -Path $config.OutputPath -Force
}

# Ensure Microsoft Graph module is available before using Connect-MgGraph
try {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Log "Microsoft.Graph module is not installed. Install it with Install-Module Microsoft.Graph" "ERROR"
        throw "Microsoft.Graph module required"
    }
    Import-Module Microsoft.Graph -ErrorAction Stop
}
catch {
    Write-Host "ERROR: Failed to import Microsoft.Graph module: $($_.Exception.Message)"
    throw
}

# Формируем полные пути
$excelFilePath = Join-Path $config.OutputPath $config.FileName
$logFilePath = Join-Path $config.OutputPath $config.LogPath

# === ФУНКЦИЯ ПРЕОБРАЗОВАНИЯ НАЗВАНИЙ ЛИЦЕНЗИЙ ===
function Convert-LicenseName {
    param([string]$SkuName)

    # Таблица преобразования SKU в человекочитаемые названия
    $licenseTranslation = @{
        # Основные M365 лицензии
        "SPB" = "Microsoft 365 Business Premium"
        "O365_BUSINESS_PREMIUM" = "Microsoft 365 Business Standard"
        "O365_BUSINESS_ESSENTIALS" = "Microsoft 365 Business Basic"

        # Windows лицензии
        "Win10_VDA_E3" = "Windows 10/11 Enterprise E3"
        "WINDOWS_STORE" = "Microsoft Store for Business"

        # Безопасность
        "THREAT_INTELLIGENCE" = "Microsoft Defender for Office 365 (Plan 2)"
        "Microsoft Defender for Office 365 (Plan 2)" = "Microsoft Defender for Office 365 (Plan 2)"
        "AAD_PREMIUM" = "Azure Active Directory Premium P1"
        "RIGHTSMANAGEMENT_ADHOC" = "Azure Information Protection P1"

        # Power Platform
        "FLOW_FREE" = "Microsoft Power Automate Free"
        "POWERAPPS_DEV" = "Microsoft Power Apps for Developer"
        "Microsoft Power Apps for Developer" = "Microsoft Power Apps for Developer"
        "POWERAPPS_VIRAL" = "Power Apps Viral"
        "CCIBOTS_PRIVPREV_VIRAL" = "Microsoft Copilot Studio Viral Trial"
        "Microsoft Copilot Studio Viral Trial" = "Microsoft Copilot Studio Viral Trial"

        # Intune
        "Microsoft_Intune_Plan_2" = "Microsoft Intune Plan 2"
        "Microsoft Intune Plan 2" = "Microsoft Intune Plan 2"

        # Power BI
        "POWER_BI_STANDARD" = "Microsoft Fabric (Free)"
        "POWER_BI_PRO" = "Power BI Pro"
        "Power BI Pro" = "Power BI Pro"

        # Exchange
        "EXCHANGESTANDARD" = "Exchange Online (Plan 1)"
        "Exchange Online (Plan 1)" = "Exchange Online (Plan 1)"

        # OneDrive
        "OneDrive for business (Plan 2)" = "OneDrive for Business (Plan 2)"
        "WACONEDRIVEENTERPRISE" = "OneDrive for Business (Plan 2)"

        # Fabric
        "Microsoft Fabric Free" = "Microsoft Fabric (Free)"

        # Teams
        "Microsoft_Teams_Exploratory_Dept" = "Microsoft Teams Exploratory"
        # Project
        "PROJECT_PLAN3_DEPT" = "Project Plan 3"

        # EMS
        "EMS" = "Enterprise Mobility + Security E3"

        # Другие
        "ENTERPRISEPACK" = "Office 365 E3"
        "ENTERPRISEPREMIUM" = "Office 365 E5"
        "INTUNE_A" = "Microsoft Intune"
        "ATP_ENTERPRISE" = "Microsoft Defender for Office 365"
        "MCOMEETADV" = "Audio Conferencing"
        "MCOEV" = "Microsoft Phone System"
        "STREAM" = "Microsoft Stream"
        "VISIO_CLIENT_SUBSCRIPTION" = "Visio Plan 2"
        "PROJECT_PROFESSIONAL" = "Project Professional"
        "MICROSOFT_BUSINESS_CENTER" = "Microsoft 365 Admin Center"
    }

    # Если есть перевод - возвращаем его, иначе возвращаем оригинал
    if ($licenseTranslation.ContainsKey($SkuName)) {
        return $licenseTranslation[$SkuName]
    } else {
        return $SkuName
    }
}

# === ФУНКЦИЯ ДЛЯ ПОЛУЧЕНИЯ ДАННЫХ ИЗ TENANT ===
function Get-EntraUsersFromTenant {
    param(
        [hashtable]$TenantConfig
    )

    Write-Log "Connectint to the Tenant: $($TenantConfig.Name)"

    try {
        # Аутентификация в Microsoft Graph для тенанта
        $body = @{
            client_id     = $TenantConfig.ClientId
            client_secret = $TenantConfig.ClientSecret
            scope         = "https://graph.microsoft.com/.default"
            grant_type    = "client_credentials"
        }

        Write-Log "Getting token for the Tenant: $($TenantConfig.Name)..."
        $tokenResponse = Invoke-RestMethod `
            -Uri "https://login.microsoftonline.com/$($TenantConfig.TenantId)/oauth2/v2.0/token" `
            -Method Post `
            -Body $body `
            -ContentType "application/x-www-form-urlencoded" `
            -ErrorAction Stop

        $secureToken = ConvertTo-SecureString $tokenResponse.access_token -AsPlainText -Force

        # Временное подключение к Graph для этого тенанта
        Write-Log "Connecting to Microsoft Graph for the Tenant: $($TenantConfig.Name)..."
        Connect-MgGraph -AccessToken $secureToken -NoWelcome | Out-Null

        Write-Log "Successful authentication in the Tenant: $($TenantConfig.Name)" "SUCCESS"

        # Получение данных лицензий
        Write-Log "Getting list of available licenses for the Tenant: $($TenantConfig.Name)..."
        $allLicenses = Get-MgSubscribedSku -All
        $licenseMap = @{}

        foreach ($license in $allLicenses) {
            $licenseMap[$license.SkuId] = $license.SkuPartNumber
        }

        Write-Log "Loaded $($licenseMap.Count) types of licenses from the Tenant: $($TenantConfig.Name)"

        # Получение данных пользователей
        Write-Log "Getting list of users from the Tenant: $($TenantConfig.Name)..."
        $users = Get-MgUser -All `
            -Property "id,displayName,userPrincipalName,accountEnabled,department,jobTitle,assignedLicenses,createdDateTime"

        # Фильтрация гостевых пользователей
        Write-Log "Filtering guest users (excluding users with '#EXT#' in UPN)..."
        $filteredUsers = @()
        $guestUsersCount = 0

        foreach ($user in $users) {
            if ($user.UserPrincipalName -like "*#EXT#*") {
                $guestUsersCount++
                Write-Log "Skipping guest user: $($user.UserPrincipalName)" "DEBUG"
            } else {
                $filteredUsers += $user
            }
        }

        Write-Log "Guest users excluded: $guestUsersCount"
        Write-Log "Users remaining after filtering: $($filteredUsers.Count)"

        Write-Log "Processing $($filteredUsers.Count) users from the Tenant: $($TenantConfig.Name)..."
        $userData = @()
        $tenantLicenseStats = @{}
        $tenantEnabledUsers = 0
        $tenantAssignedLicenses = @()

        foreach ($user in $filteredUsers) {
            # Определение лицензий пользователя
            $userLicenseNames = @()
            if ($user.AssignedLicenses) {
                foreach ($assignedLicense in $user.AssignedLicenses) {
                    if ($licenseMap.ContainsKey($assignedLicense.SkuId)) {
                        $licenseName = $licenseMap[$assignedLicense.SkuId]
                        # Преобразуем название лицензии
                        $friendlyLicenseName = Convert-LicenseName -SkuName $licenseName
                        $userLicenseNames += $friendlyLicenseName

                        # Добавляем в общий список назначенных лицензий
                        if ($friendlyLicenseName -notin $tenantAssignedLicenses) {
                            $tenantAssignedLicenses += $friendlyLicenseName
                        }

                        # Статистика по лицензиям
                        if ($tenantLicenseStats.ContainsKey($friendlyLicenseName)) {
                            $tenantLicenseStats[$friendlyLicenseName]++
                        } else {
                            $tenantLicenseStats[$friendlyLicenseName] = 1
                        }
                    }
                }
            }

            # Статистика по активным пользователям
            if ($user.AccountEnabled -eq $true) {
                $tenantEnabledUsers++
            }

            # Формируем объект пользователя
            $userObject = [PSCustomObject]@{
                DisplayName       = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                AccountEnabled    = if ($user.AccountEnabled) { "Yes" } else { "No" }
                Department        = $user.Department
                JobTitle          = $user.JobTitle
                CreatedDate       = $user.CreatedDateTime.ToString("yyyy-MM-dd")
                TotalLicenses     = $userLicenseNames.Count
                UserLicenses      = $userLicenseNames  # Временное поле для лицензий
            }

            $userData += $userObject
        }

        Write-Log "Data from the Tenant: $($TenantConfig.Name) successfully retrieved" "SUCCESS"

        return @{
            Users = $userData
            LicenseStats = $tenantLicenseStats
            EnabledUsers = $tenantEnabledUsers
            AssignedLicenses = $tenantAssignedLicenses
            TotalUsers = $filteredUsers.Count
            GuestUsersExcluded = $guestUsersCount
        }

    } catch {
        Write-Log "Error occurred while retrieving data from the Tenant: $($TenantConfig.Name): $_" "ERROR"
        Write-Log "Call Stack: $($_.Exception.Message)" "ERROR"

        return $null
    } finally {
        # Отключаемся от Graph в любом случае
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Write-Log "Disconnected from Microsoft Graph for the Tenant: $($TenantConfig.Name)" "INFO"
        } catch {}
    }
}

# === ФУНКЦИЯ НАСТРОЙКИ ШИРИНЫ СТОЛБЦОВ ===
function Set-ExcelColumnWidths {
    param(
        [string]$FilePath,
        [string]$WorksheetName,
        [hashtable]$Widths
    )

    try {
        # Открываем Excel файл
        $excel = Open-ExcelPackage -Path $FilePath

        # Получаем лист
        $worksheet = $excel.Workbook.Worksheets[$WorksheetName]

        if ($worksheet) {
            Write-Log "Setting column widths for the worksheet '$WorksheetName'..."

            # Проходим по всем столбцам листа
            for ($col = 1; $col -le $worksheet.Dimension.Columns; $col++) {
                $header = $worksheet.Cells[1, $col].Value
                if ($header -and $Widths.ContainsKey($header)) {
                    # Устанавливаем ширину столбца
                    $worksheet.Column($col).Width = $Widths[$header]
                    Write-Log "  Width of column '$header': $($Widths[$header]) characters"
                }
            }

            # Сохраняем изменения
            Close-ExcelPackage -ExcelPackage $excel
            Write-Log "Column widths for the worksheet '$WorksheetName' configured" "INFO"
            return $true
        } else {
            Write-Log "Worksheet '$WorksheetName' not found" "WARNING"
            Close-ExcelPackage -ExcelPackage $excel -NoSave
            return $false
        }
    } catch {
        Write-Log "Error occurred while setting column widths for the worksheet '$WorksheetName': $_" "WARNING"
        Write-Log "Call Stack: $($_.Exception.Message)" "WARNING"
        return $false
    }
}

# === ФУНКЦИЯ ДОБАВЛЕНИЯ ИТОГОВ В EXCEL ===
function Add-ExcelTotalsRow {
    param(
        [string]$FilePath,
        [string]$WorksheetName,
        [array]$LicenseColumns
    )

    try {
        Write-Log "Adding totals row for the worksheet '$WorksheetName'..."

        # Открываем Excel файл
        $excel = Open-ExcelPackage -Path $FilePath

        # Получаем лист
        $worksheet = $excel.Workbook.Worksheets[$WorksheetName]

        if ($worksheet) {
            # Определяем последнюю строку данных
            $lastRow = $worksheet.Dimension.Rows
            $firstDataRow = 2  # Заголовок в строке 1

            # Добавляем строку после данных
            $totalsRow = $lastRow + 1

            # Устанавливаем заголовок в первой колонке
            $worksheet.Cells[$totalsRow, 1].Value = "TOTAL LICENSES:"
            $worksheet.Cells[$totalsRow, 1].Style.Font.Bold = $true
            $worksheet.Cells[$totalsRow, 1].Style.Fill.PatternType = "Solid"
            $worksheet.Cells[$totalsRow, 1].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)

            # Находим индексы колонок с лицензиями
            $licenseColumnIndices = @{}
            for ($col = 1; $col -le $worksheet.Dimension.Columns; $col++) {
                $header = $worksheet.Cells[1, $col].Value
                if ($header -in $LicenseColumns) {
                    $licenseColumnIndices[$header] = $col
                }
            }

            # Добавляем формулы SUBTOTAL для подсчета "+" в каждой колонке
            # SUBTOTAL(103, range) - считает только видимые строки (работает с фильтрацией)
            # Мы будем считать количество непустых ячеек в колонке
            foreach ($licenseName in $LicenseColumns) {
                if ($licenseColumnIndices.ContainsKey($licenseName)) {
                    $colIndex = $licenseColumnIndices[$licenseName]

                    # Формула для подсчета непустых ячеек в колонке (только видимые строки)
                    $formula = "SUBTOTAL(103,{0}2:{0}{1})" -f [char](64 + $colIndex), $lastRow

                    $worksheet.Cells[$totalsRow, $colIndex].Formula = $formula
                    $worksheet.Cells[$totalsRow, $colIndex].Style.Font.Bold = $true
                    $worksheet.Cells[$totalsRow, $colIndex].Style.NumberFormat = "0"
                    $worksheet.Cells[$totalsRow, $colIndex].Style.Fill.PatternType = "Solid"
                    $worksheet.Cells[$totalsRow, $colIndex].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)

                    Write-Log "  Formula added for license '$licenseName' in column $colIndex"
                }
            }

            # Добавляем формулу для TotalLicenses
            $totalLicensesCol = $worksheet.Dimension.Columns  # Последняя колонка
            $formulaTotal = "SUBTOTAL(109,{0}2:{0}{1})" -f [char](64 + $totalLicensesCol), $lastRow
            $worksheet.Cells[$totalsRow, $totalLicensesCol].Formula = $formulaTotal
            $worksheet.Cells[$totalsRow, $totalLicensesCol].Style.Font.Bold = $true
            $worksheet.Cells[$totalsRow, $totalLicensesCol].Style.Fill.PatternType = "Solid"
            $worksheet.Cells[$totalsRow, $totalLicensesCol].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)

            # Сохраняем изменения
            Close-ExcelPackage -ExcelPackage $excel
            Write-Log "Totals row successfully added to row $totalsRow" "SUCCESS"
            return $true
        } else {
            Write-Log "Worksheet '$WorksheetName' not found" "WARNING"
            Close-ExcelPackage -ExcelPackage $excel -NoSave
            return $false
        }
    } catch {
        Write-Log "Error occurred while adding totals row for the worksheet '$WorksheetName': $_" "WARNING"
        Write-Log "Call Stack: $($_.Exception.Message)" "WARNING"
        return $false
    }
}

# === ФУНКЦИЯ ЗАГРУЗКИ В SHAREPOINT ===
function Upload-ToSharePoint {
    param(
        [string]$FilePath,
        [hashtable]$SharePointConfig
    )

    try {
        Write-Log "Starting file upload to SharePoint..."

        if (-not (Test-Path $FilePath)) {
            throw "Local file not found: $FilePath"
        }

        # Подключаемся к Graph для загрузки в SharePoint (используем первый тенант для аутентификации)
        $firstTenant = $config.Tenants[0]
        $body = @{
            client_id     = $firstTenant.ClientId
            client_secret = $firstTenant.ClientSecret
            scope         = "https://graph.microsoft.com/.default"
            grant_type    = "client_credentials"
        }

        Write-Log "Authenticating in SharePoint..."
        $tokenResponse = Invoke-RestMethod `
            -Uri "https://login.microsoftonline.com/$($firstTenant.TenantId)/oauth2/v2.0/token" `
            -Method Post `
            -Body $body `
            -ContentType "application/x-www-form-urlencoded" `
            -ErrorAction Stop

        $secureToken = ConvertTo-SecureString $tokenResponse.access_token -AsPlainText -Force
        Connect-MgGraph -AccessToken $secureToken -NoWelcome | Out-Null

        Write-Log "Authentication for SharePoint completed"

        # Получаем сайт SharePoint
        Write-Log "Searching for SharePoint site: $($SharePointConfig.SiteUrl)"
        $site = Get-MgSite -SiteId $SharePointConfig.SiteUrl -ErrorAction Stop

        if (-not $site) {
            throw "SharePoint site not found: $($SharePointConfig.SiteUrl)"
        }
        Write-Log "SharePoint site found: $($site.DisplayName)"

        # Получаем библиотеку документов
        Write-Log "Searching for document library: $($SharePointConfig.Library)"
        $drives = Get-MgSiteDrive -SiteId $site.Id -All -ErrorAction Stop
        $drive = $drives | Where-Object { $_.Name -eq $SharePointConfig.Library }

        if (-not $drive) {
            Write-Log "Available libraries: $($drives.Name -join ', ')"
            throw "Document library '$($SharePointConfig.Library)' not found"
        }
        Write-Log "Document library found: $($drive.Name)"

        # Подготавливаем путь для загрузки
        $fileName = Split-Path $FilePath -Leaf
        $uploadPath = "$($SharePointConfig.Folder)$fileName"

        # Формируем DriveItemId в правильном формате - используем ${} для корректного разбора переменной
        $driveItemId = "root:/${uploadPath}:"
        Write-Log "Path for upload: $driveItemId"

        # Загружаем файл
        Write-Log "Starting file upload to SharePoint..."
        Set-MgDriveItemContent -DriveId $drive.Id -DriveItemId $driveItemId -InFile $FilePath -ErrorAction Stop

        Write-Log "File successfully uploaded to SharePoint: $uploadPath" "SUCCESS"
        return $true

    } catch {
        $errorMsg = "Error occurred while uploading to SharePoint: $_"
        Write-Log $errorMsg "ERROR"
        Write-Log "Details of the error: $($_.Exception.Message)" "ERROR"
        return $false
    } finally {
        # Отключаемся от Graph
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Write-Log "Disconnecting from Microsoft Graph after upload to SharePoint" "INFO"
        } catch {}
    }
}

# === ОСНОВНОЙ ПРОЦЕСС ===
try {
    $startTime = Get-Date
    Write-Log "=" * 60
    Write-Log "RUN MERGED SCRIPT: Entra ID -> Excel -> SharePoint" "INFO"
    Write-Log "=" * 60

    # === ЭТАП 1: ПОЛУЧЕНИЕ ДАННЫХ ИЗ ВСЕХ TENANTS ===
    Write-Log "Stage 1: Retrieving data from all tenants..."

    $allUsers = @()
    $allLicenseStats = @{}
    $allAssignedLicenses = @()
    $totalUsers = 0
    $totalEnabledUsers = 0
    $totalGuestUsersExcluded = 0
    $tenantResults = @()

    foreach ($tenant in $config.Tenants) {
        Write-Log "Processing tenant: $($tenant.Name) (Tenant ID: $($tenant.TenantId))"

        $tenantData = Get-EntraUsersFromTenant -TenantConfig $tenant

        if ($tenantData) {
            $tenantResults += @{
                Name = $tenant.Name
                Data = $tenantData
            }

            # Собираем общую статистику
            $totalUsers += $tenantData.TotalUsers
            $totalEnabledUsers += $tenantData.EnabledUsers
            $totalGuestUsersExcluded += $tenantData.GuestUsersExcluded

            # Объединяем статистику лицензий
            foreach ($licenseKey in $tenantData.LicenseStats.Keys) {
                if ($allLicenseStats.ContainsKey($licenseKey)) {
                    $allLicenseStats[$licenseKey] += $tenantData.LicenseStats[$licenseKey]
                } else {
                    $allLicenseStats[$licenseKey] = $tenantData.LicenseStats[$licenseKey]
                }
            }

            # Объединяем списки лицензий
            foreach ($license in $tenantData.AssignedLicenses) {
                if ($license -notin $allAssignedLicenses) {
                    $allAssignedLicenses += $license
                }
            }

            # Добавляем пользователей
            $allUsers += $tenantData.Users
        }
    }

    if ($allUsers.Count -eq 0) {
        throw "Failed to retrieve data from any tenant"
    }

    Write-Log "Data collected from $($tenantResults.Count) tenants"
    Write-Log "Total users (excluding guests): $totalUsers"
    Write-Log "Excluded guest users: $totalGuestUsersExcluded"
    Write-Log "Total enabled users: $totalEnabledUsers"

    # === ЭТАП 2: ОБЪЕДИНЕНИЕ И ОБРАБОТКА ДАННЫХ ===
    Write-Log "Stage 2: Merging and processing data..."

    # Сортируем лицензии по алфавиту
    $allAssignedLicenses = $allAssignedLicenses | Sort-Object
    Write-Log "Total unique license types found: $($allAssignedLicenses.Count)"

    # Преобразуем данные пользователей для Excel
    $userData = @()

    foreach ($user in $allUsers) {
        # Создаем новый объект с базовыми полями (без SourceTenant и LastSync)
        $userObject = [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            AccountEnabled    = $user.AccountEnabled
            Department        = $user.Department
            JobTitle          = $user.JobTitle
            CreatedDate       = $user.CreatedDate
        }

        # Добавляем столбцы для всех лицензий
        foreach ($licenseName in $allAssignedLicenses) {
            $licenseValue = if ($user.UserLicenses -contains $licenseName) { "+" } else { "" }
            $userObject | Add-Member -MemberType NoteProperty -Name $licenseName -Value $licenseValue
        }

        # Добавляем TotalLicenses ПОСЛЕ всех лицензий
        $userObject | Add-Member -MemberType NoteProperty -Name "TotalLicenses" -Value $user.TotalLicenses

        $userData += $userObject
    }

    # Сортируем пользователей по DisplayName
    $userData = $userData | Sort-Object DisplayName
    Write-Log "Data sorted by DisplayName (total $($userData.Count) records)"


# === ЭТАП 3: СОЗДАНИЕ EXCEL ФАЙЛА С ИТОГОВОЙ СТРОКОЙ ===
Write-Log "Stage 3: Creating Excel file with summary row..."

# --- КЛЮЧЕВОЕ ИЗМЕНЕНИЕ: ОПРЕДЕЛЯЕМ ЦЕНЫ ЛИЦЕНЗИЙ ---
$licensePricesJson = ConvertFrom-Json $env:LICENSE_PRICES_JSON
$licensePrices = ConvertTo-Hashtable -InputObject $licensePricesJson
Write-Log "Loaded $($licensePrices.Count) license prices for calculation"

# Функция для конвертации номера столбца в букву Excel
function Convert-ColumnNumberToLetter {
    param([int]$ColumnNumber)

    $dividend = $ColumnNumber
    $columnLetter = ''

    while ($dividend -gt 0) {
        $modulo = ($dividend - 1) % 26
        $columnLetter = [char](65 + $modulo) + $columnLetter
        $dividend = [Math]::Floor(($dividend - $modulo) / 26)
    }

    return $columnLetter
}

# Создаем временный файл
$tempFilePath = Join-Path $config.OutputPath "temp_$($config.FileName)"

# Сначала создаем файл без таблицы, чтобы потом её создать с правильным диапазоном
$excelParams = @{
    Path          = $tempFilePath
    WorksheetName = $config.MainSheetName
    FreezeTopRow  = $true
    BoldTopRow    = $true
    AutoFilter    = $true
    PassThru      = $true
}

$excelPackage = $userData | Export-Excel @excelParams
$worksheet = $excelPackage.Workbook.Worksheets[$config.MainSheetName]

Write-Log "Main worksheet created"

# Определяем количество строк и столбцов
$rowCount = $worksheet.Dimension.Rows
$colCount = $worksheet.Dimension.Columns

Write-Log "Size of table: $rowCount rows, $colCount columns"

# СОЗДАЕМ ТАБЛИЦУ ТОЛЬКО ДЛЯ ДАННЫХ ПОЛЬЗОВАТЕЛЕЙ (без итоговой строки)
$lastColumnLetter = Convert-ColumnNumberToLetter -ColumnNumber $colCount
$tableRange = "A1:$($lastColumnLetter)$rowCount"
$table = $worksheet.Tables.Add($tableRange, "EntraUsers")
$table.TableStyle = [OfficeOpenXml.Table.TableStyles]::Medium6
$table.ShowFilter = $true

# --- ПРИМЕНЯЕМ НАСТРОЙКИ ШИРИНЫ СТОЛБЦОВ И ФОРМАТИРОВАНИЕ ---
Write-Log "Configuring column widths and formatting..."

# Подготавливаем настройки ширины столбцов
$updatedWidths = @{
    DisplayName       = 30
    UserPrincipalName = 35
    AccountEnabled    = 15
    Department        = 25
    JobTitle          = 25
    CreatedDate       = 12
    TotalLicenses     = 12
}

# Добавляем ширину для столбцов лицензий
foreach ($licenseName in $allAssignedLicenses) {
    $licenseLength = $licenseName.Length
    if ($licenseLength -le 25) {
        $width = 30
    } elseif ($licenseLength -le 35) {
        $width = 40
    } else {
        $width = 50
    }
    $updatedWidths[$licenseName] = $width
}

# Применяем ширину столбцов
for ($col = 1; $col -le $colCount; $col++) {
    $header = $worksheet.Cells[1, $col].Value
    if ($updatedWidths.ContainsKey($header)) {
        $worksheet.Column($col).Width = $updatedWidths[$header]
    }
}

# Центрируем ячейки с лицензиями
$basicColumns = @("DisplayName", "UserPrincipalName", "AccountEnabled", "Department", "JobTitle", "CreatedDate", "TotalLicenses")

# Сначала определим столбцы лицензий
$licenseColumns = @()
for ($col = 1; $col -le $colCount; $col++) {
    $header = $worksheet.Cells[1, $col].Value
    if ($header -ne $null -and $header -notin $basicColumns) {
        $licenseColumns += $col

        # Центрируем содержимое в столбцах с лицензиями
        for ($row = 2; $row -le $rowCount; $row++) {
            $worksheet.Cells[$row, $col].Style.HorizontalAlignment = "Center"
        }
    }
}

Write-Log "Found $($licenseColumns.Count) columns with licenses"

# --- СОЗДАЕМ СКРЫТЫЙ ЛИСТ ДЛЯ ВЫЧИСЛЕНИЙ ---
Write-Log "Creating hidden sheet for calculations..."

$calcSheet = $excelPackage.Workbook.Worksheets.Add("_Calculations")
$calcSheet.Hidden = "Hidden"  # Скрываем лист

# Копируем заголовки с основного листа
for ($col = 1; $col -le $colCount; $col++) {
    $header = $worksheet.Cells[1, $col].Value
    $calcSheet.Cells[1, $col].Value = $header
}

# Заполняем скрытый лист формулами подсчета
if ($licenseColumns.Count -gt 0) {
    Write-Log "Filling hidden sheet with formulas..."

    foreach ($col in $licenseColumns) {
        $colLetter = Convert-ColumnNumberToLetter -ColumnNumber $col

        # В столбцах лицензий создаем формулы подсчета видимых "+"
        for ($row = 2; $row -le $rowCount; $row++) {
            # Формула: если строка видима (SUBTOTAL(103)=1) и содержит "+", то 1, иначе 0
            $formula = "=IF(SUBTOTAL(103,'$($config.MainSheetName)'!${colLetter}$row), IF('$($config.MainSheetName)'!${colLetter}$row=""+"",1,0), 0)"
            $calcSheet.Cells[$row, $col].Formula = $formula
        }

        Write-Log "  Added formulas for column $colLetter"
    }
} else {
    Write-Log "No columns with licenses found for processing" "WARNING"
}

# --- ДОБАВЛЕНИЕ ИТОГОВОЙ СТРОКИ ПОД ТАБЛИЦЕЙ ---
Write-Log "Adding summary row under the table..."

# Итоговая строка будет ПОД таблицей (через 2 строки для отступа)
$summaryRow = $rowCount + 2

Write-Log "Adding totals to row $summaryRow (under the table)..."

# Заполняем итоговую строку (КОЛИЧЕСТВО)
$worksheet.Cells[$summaryRow, 1].Value = "Licenses Count (Filter Friendly):"
$worksheet.Cells[$summaryRow, 1].Style.Font.Bold = $true
$worksheet.Cells[$summaryRow, 1].Style.Fill.PatternType = "Solid"
$worksheet.Cells[$summaryRow, 1].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(198, 239, 206))

# Для столбцов с лицензиями добавляем ПРОСТУЮ формулу суммирования
if ($licenseColumns.Count -gt 0) {
    foreach ($col in $licenseColumns) {
        $colLetter = Convert-ColumnNumberToLetter -ColumnNumber $col

        # ПРОСТАЯ и стабильная формула: сумма значений из скрытого листа
        $formula = "=SUM('_Calculations'!${colLetter}2:${colLetter}$rowCount)"

        $worksheet.Cells[$summaryRow, $col].Formula = $formula
        $worksheet.Cells[$summaryRow, $col].Style.Font.Bold = $true
        $worksheet.Cells[$summaryRow, $col].Style.HorizontalAlignment = "Center"
        $worksheet.Cells[$summaryRow, $col].Style.Fill.PatternType = "Solid"
        $worksheet.Cells[$summaryRow, $col].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 235, 156))
        $worksheet.Cells[$summaryRow, $col].Style.Numberformat.Format = 0

        Write-Log "  Added summary for column $colLetter"
    }
}

# Для TotalLicenses используем простую сумму
$totalLicensesCol = 0
for ($col = 1; $col -le $colCount; $col++) {
    if ($worksheet.Cells[1, $col].Value -eq "TotalLicenses") {
        $totalLicensesCol = $col
        break
    }
}

if ($totalLicensesCol -gt 0) {
    $colLetter = Convert-ColumnNumberToLetter -ColumnNumber $totalLicensesCol
    # Прямая формула без сложных вычислений
    $formula = "=SUBTOTAL(109,${colLetter}2:${colLetter}$rowCount)"
    $worksheet.Cells[$summaryRow, $totalLicensesCol].Formula = $formula
    $worksheet.Cells[$summaryRow, $totalLicensesCol].Style.Font.Bold = $true
    $worksheet.Cells[$summaryRow, $totalLicensesCol].Style.HorizontalAlignment = "Center"
    $worksheet.Cells[$summaryRow, $totalLicensesCol].Style.Fill.PatternType = "Solid"
    $worksheet.Cells[$summaryRow, $totalLicensesCol].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(198, 239, 206))
    $worksheet.Cells[$summaryRow, $totalLicensesCol].Style.Numberformat.Format = 0

    Write-Log "  Added summary for TotalLicenses (column $colLetter)"
}

# --- КЛЮЧЕВОЕ ИЗМЕНЕНИЕ: ДОБАВЛЯЕМ РАСЧЕТ СТОИМОСТИ ---
Write-Log "Adding rows with license cost calculations..."

# Создаем справочник: название лицензии -> номер столбца
$licenseColumnMap = @{}
for ($col = 1; $col -le $colCount; $col++) {
    $header = $worksheet.Cells[1, $col].Value
    if ($header -ne $null) {
        $licenseColumnMap[$header] = $col
    }
}

# 1. Строка "Price per License"
$pricePerLicenseRow = $summaryRow + 2
$worksheet.Cells[$pricePerLicenseRow, 1].Value = "Price per License ($):"
$worksheet.Cells[$pricePerLicenseRow, 1].Style.Font.Bold = $true
$worksheet.Cells[$pricePerLicenseRow, 1].Style.Fill.PatternType = "Solid"
$worksheet.Cells[$pricePerLicenseRow, 1].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(234, 241, 255)) # Светло-синий

# 2. Строка "Total Cost per License (Monthly)"
$totalCostRow = $pricePerLicenseRow + 1
$worksheet.Cells[$totalCostRow, 1].Value = "Total Cost per License (Monthly $):"
$worksheet.Cells[$totalCostRow, 1].Style.Font.Bold = $true
$worksheet.Cells[$totalCostRow, 1].Style.Fill.PatternType = "Solid"
$worksheet.Cells[$totalCostRow, 1].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 242, 204)) # Светло-оранжевый

# 3. Строка "Total Cost (Yearly)"
$yearlyCostRow = $totalCostRow + 1
$worksheet.Cells[$yearlyCostRow, 1].Value = "Total Cost per License (Yearly $):"
$worksheet.Cells[$yearlyCostRow, 1].Style.Font.Bold = $true
$worksheet.Cells[$yearlyCostRow, 1].Style.Fill.PatternType = "Solid"
$worksheet.Cells[$yearlyCostRow, 1].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(226, 239, 218)) # Светло-зеленый

# Добавляем новые столбцы для общих итогов по деньгам (справа от столбца TotalLicenses)
$totalMonthlyCol = $totalLicensesCol + 1
$totalYearlyCol = $totalLicensesCol + 2

# Заголовки для новых столбцов с общими суммами
$worksheet.Cells[1, $totalMonthlyCol].Value = "Total Monthly Cost ($)"
$worksheet.Cells[1, $totalMonthlyCol].Style.Font.Bold = $true
$worksheet.Cells[1, $totalMonthlyCol].Style.HorizontalAlignment = "Center"
$worksheet.Cells[1, $totalMonthlyCol].Style.Fill.PatternType = "Solid"
$worksheet.Cells[1, $totalMonthlyCol].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 242, 204))

$worksheet.Cells[1, $totalYearlyCol].Value = "Total Yearly Cost ($)"
$worksheet.Cells[1, $totalYearlyCol].Style.Font.Bold = $true
$worksheet.Cells[1, $totalYearlyCol].Style.HorizontalAlignment = "Center"
$worksheet.Cells[1, $totalYearlyCol].Style.Fill.PatternType = "Solid"
$worksheet.Cells[1, $totalYearlyCol].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(226, 239, 218))

# Заполняем строки с ценами и расчётами для каждой лицензии
foreach ($licenseName in $licensePrices.Keys) {
    if ($licenseColumnMap.ContainsKey($licenseName)) {
        $col = $licenseColumnMap[$licenseName]
        $price = $licensePrices[$licenseName]
        $colLetter = Convert-ColumnNumberToLetter -ColumnNumber $col
        $countCellAddress = "${colLetter}${summaryRow}" # Адрес ячейки с количеством

        # Цена за одну лицензию (просто число)
        $worksheet.Cells[$pricePerLicenseRow, $col].Value = $price
        $worksheet.Cells[$pricePerLicenseRow, $col].Style.Numberformat.Format = "0.00"
        $worksheet.Cells[$pricePerLicenseRow, $col].Style.HorizontalAlignment = "Center"
        $worksheet.Cells[$pricePerLicenseRow, $col].Style.Font.Bold = $true

        # Общая стоимость по лицензии в месяц (цена × количество)
        $monthlyFormula = "=${colLetter}${pricePerLicenseRow} * ${countCellAddress}"
        $worksheet.Cells[$totalCostRow, $col].Formula = $monthlyFormula
        $worksheet.Cells[$totalCostRow, $col].Style.Numberformat.Format = "0.00"
        $worksheet.Cells[$totalCostRow, $col].Style.HorizontalAlignment = "Center"
        $worksheet.Cells[$totalCostRow, $col].Style.Font.Bold = $true

        # Общая стоимость по лицензии в год (месячная × 12)
        $yearlyFormula = "=${colLetter}${totalCostRow} * 12"
        $worksheet.Cells[$yearlyCostRow, $col].Formula = $yearlyFormula
        $worksheet.Cells[$yearlyCostRow, $col].Style.Numberformat.Format = "0.00"
        $worksheet.Cells[$yearlyCostRow, $col].Style.HorizontalAlignment = "Center"
        $worksheet.Cells[$yearlyCostRow, $col].Style.Font.Bold = $true

        Write-Log "  Added cost calculation for license '$licenseName'"
    } else {
        Write-Log "  Warning: Price defined for '$licenseName', but column not found in report" "WARNING"
    }
}

# Общие итоги в новых столбцах (справа)
# Итоговая месячная стоимость (сумма всех лицензий)
$totalMonthlyLetter = Convert-ColumnNumberToLetter -ColumnNumber $totalMonthlyCol
$monthlySumFormula = "=SUM("
foreach ($licenseName in $licensePrices.Keys) {
    if ($licenseColumnMap.ContainsKey($licenseName)) {
        $col = $licenseColumnMap[$licenseName]
        $colLetter = Convert-ColumnNumberToLetter -ColumnNumber $col
        $monthlySumFormula += "${colLetter}${totalCostRow},"
    }
}
$monthlySumFormula = $monthlySumFormula.TrimEnd(',') + ")"
$worksheet.Cells[$totalCostRow, $totalMonthlyCol].Formula = $monthlySumFormula
$worksheet.Cells[$totalCostRow, $totalMonthlyCol].Style.Numberformat.Format = "0.00"
$worksheet.Cells[$totalCostRow, $totalMonthlyCol].Style.Font.Bold = $true
$worksheet.Cells[$totalCostRow, $totalMonthlyCol].Style.HorizontalAlignment = "Center"

# Итоговая годовая стоимость (сумма всех лицензий × 12, или месячная × 12)
$totalYearlyLetter = Convert-ColumnNumberToLetter -ColumnNumber $totalYearlyCol
$yearlyFormula = "=${totalMonthlyLetter}${totalCostRow} * 12"
$worksheet.Cells[$yearlyCostRow, $totalYearlyCol].Formula = $yearlyFormula
$worksheet.Cells[$yearlyCostRow, $totalYearlyCol].Style.Numberformat.Format = "0.00"
$worksheet.Cells[$yearlyCostRow, $totalYearlyCol].Style.Font.Bold = $true
$worksheet.Cells[$yearlyCostRow, $totalYearlyCol].Style.HorizontalAlignment = "Center"

# Также продублируем итоги в строке с количеством лицензий для наглядности
$worksheet.Cells[$summaryRow, $totalMonthlyCol].Value = "Monthly Cost"
$worksheet.Cells[$summaryRow, $totalMonthlyCol].Style.Font.Bold = $true
$worksheet.Cells[$summaryRow, $totalMonthlyCol].Style.HorizontalAlignment = "Center"

$worksheet.Cells[$summaryRow, $totalYearlyCol].Value = "Yearly Cost"
$worksheet.Cells[$summaryRow, $totalYearlyCol].Style.Font.Bold = $true
$worksheet.Cells[$summaryRow, $totalYearlyCol].Style.HorizontalAlignment = "Center"

# Обновляем примечание с учетом новых строк
$explanationRow = $yearlyCostRow + 1
$worksheet.Cells[$explanationRow, 1].Value = "Note: Costs are calculated based on prices per user. Monthly cost = Price × Count. Yearly cost = Monthly × 12."
$worksheet.Cells[$explanationRow, 1].Style.Font.Italic = $true
$worksheet.Cells[$explanationRow, 1].Style.Font.Size = 9
$worksheet.Cells[$explanationRow, 1].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)

# Сохраняем изменения
Close-ExcelPackage -ExcelPackage $excelPackage
Write-Log "Summary row and cost calculations successfully added under the table" "SUCCESS"

# =====================================================================================================
    # === ЭТАП 5: СОЗДАНИЕ ЛИСТА SUMMARY ===
    Write-Log "Adding Summary sheet..."

    $summaryData = @()

    # Общая статистика
    $summaryData += [PSCustomObject]@{
        Metric      = "Total Users (All Tenants)"
        Value       = $totalUsers
        Description = "Total user accounts across all tenants (excluding guest users)"
    }

    $summaryData += [PSCustomObject]@{
        Metric      = "Active Users (All Tenants)"
        Value       = $totalEnabledUsers
        Description = "Accounts with AccountEnabled = true across all tenants"
    }

    $summaryData += [PSCustomObject]@{
        Metric      = "Guest Users Excluded"
        Value       = $totalGuestUsersExcluded
        Description = "Guest users with '#EXT#' in UPN excluded from report"
    }

    # Статистика по тенантам
    foreach ($tenantResult in $tenantResults) {
        $tenantData = $tenantResult.Data
        $summaryData += [PSCustomObject]@{
            Metric      = "Users from $($tenantResult.Name)"
            Value       = $tenantData.TotalUsers
            Description = "User accounts in $($tenantResult.Name) (excluding guests)"
        }

        $summaryData += [PSCustomObject]@{
            Metric      = "Active Users in $($tenantResult.Name)"
            Value       = $tenantData.EnabledUsers
            Description = "Active accounts in $($tenantResult.Name)"
        }

        $summaryData += [PSCustomObject]@{
            Metric      = "Guest Users in $($tenantResult.Name)"
            Value       = $tenantData.GuestUsersExcluded
            Description = "Guest users excluded in $($tenantResult.Name)"
        }
    }

    $usersWithLicenses = ($userData | Where-Object { $_.TotalLicenses -gt 0 }).Count
    $summaryData += [PSCustomObject]@{
        Metric      = "Users with Licenses"
        Value       = $usersWithLicenses
        Description = "Users with at least one assigned license"
    }

    $totalAssignedLicenses = ($userData | Measure-Object -Property TotalLicenses -Sum).Sum
    $summaryData += [PSCustomObject]@{
        Metric      = "Total Assigned Licenses"
        Value       = $totalAssignedLicenses
        Description = "Total number of assigned licenses (duplicates counted)"
    }

    $summaryData += [PSCustomObject]@{
        Metric      = "Unique License Types"
        Value       = $allAssignedLicenses.Count
        Description = "Different types of licenses assigned to users"
    }

    # Add license type statistics
    foreach ($licenseType in $allLicenseStats.Keys | Sort-Object) {
        $summaryData += [PSCustomObject]@{
            Metric      = "License: $licenseType"
            Value       = $allLicenseStats[$licenseType]
            Description = "Number of assignments across all tenants"
        }
    }

    # Настройки ширины столбцов для Summary
    $summaryWidths = @{
        Metric      = 60  # Увеличим для длинных названий лицензий
        Value       = 15
        Description = 70
    }

    # Export Summary sheet
    $summaryParams = @{
        Path          = $tempFilePath
        WorksheetName = $config.SummarySheetName
        FreezeTopRow  = $true
        BoldTopRow    = $true
        AutoFilter    = $true
        Append        = $true
    }

    $summaryData | Export-Excel @summaryParams
    Write-Log "Summary sheet created"

    # Configure column widths for Summary sheet
    $widthResult2 = Set-ExcelColumnWidths -FilePath $tempFilePath -WorksheetName $config.SummarySheetName -Widths $summaryWidths

    # Переименовываем временный файл в окончательный
    if (Test-Path $excelFilePath) {
        Remove-Item $excelFilePath -Force
    }
    Rename-Item -Path $tempFilePath -NewName $config.FileName
    Write-Log "File renamed: $excelFilePath"

    # Проверка размера файла
    $fileSize = (Get-Item $excelFilePath).Length / 1MB
    Write-Log "File created: $excelFilePath ($([math]::Round($fileSize, 2)) MB)" "SUCCESS"

    # Итог настройки ширины столбцов
    if ($widthResult1 -and $widthResult2) {
        Write-Log "Width of columns successfully configured for both sheets" "SUCCESS"
    } else {
        Write-Log "There were issues with configuring column widths (but the file was created)" "WARNING"
    }

    # === ЭТАП 6: ЗАГРУЗКА В SHAREPOINT ===
    Write-Log "Stage 6: Uploading file to SharePoint..."

    $uploadResult = Upload-ToSharePoint -FilePath $excelFilePath -SharePointConfig $config.SharePoint

    if ($uploadResult) {
        Write-Log "File uploaded to SharePoint successfully" "SUCCESS"
    } else {
        Write-Log "File upload to SharePoint failed" "WARNING"
    }

    # === ЭТАП 7: ИТОГОВЫЙ ОТЧЕТ ===
    $duration = New-TimeSpan -Start $startTime -End (Get-Date)
    Write-Log "=" * 60
    Write-Log "Merged PROCESS COMPLETED" "INFO"
    Write-Log "Total execution time: $([math]::Round($duration.TotalMinutes, 2)) minutes"
    Write-Log "Processed tenants: $($config.Tenants.Count)"
    Write-Log "Total users (excluding guests): $totalUsers"
    Write-Log "Active users: $totalEnabledUsers"
    Write-Log "Excluded guest users: $totalGuestUsersExcluded"
    Write-Log "Unique license types: $($allAssignedLicenses.Count)"
    Write-Log "File saved locally: $excelFilePath"
    Write-Log "Uploaded to SharePoint: $(if ($uploadResult) { 'Yes' } else { 'No' })"
    Write-Log "A row 'TOTAL LICENSES' has been added to the Excel file with dynamic calculation" "SUCCESS"
    Write-Log "SUBTOTAL formulas will be recalculated when filters are applied" "INFO"
    Write-Log "=" * 60

} catch {
    $errorMsg = "Critical error in merged script: $_"
    Write-Log $errorMsg "ERROR"
    Write-Log "Call stack: $($_.ScriptStackTrace)" "ERROR"
    throw $_
} finally {
    # Удаляем временный файл, если он остался
    if ($tempFilePath -and (Test-Path $tempFilePath)) {
        Remove-Item $tempFilePath -Force -ErrorAction SilentlyContinue
    }

    # Отключение от Graph (на всякий случай)
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Log "Disconnecting from Microsoft Graph" "INFO"
    } catch {}
}