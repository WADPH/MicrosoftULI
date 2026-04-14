# Docker запуск скрипта MicrosoftULI

## Структура

- **Dockerfile** — образ PowerShell 7.4 на Alpine Linux с запуском cron
- **docker-compose.yml** — окружение с volumes и подгрузкой `.env`
- **.env** — единственный файл конфигурации (однострочный JSON)
- **output/** — директория для выходных файлов (Excel, логи)

## Подготовка

### 1. Создать .env файл

Скопируйте `.env.example` и заполните реальными значениями:

```bash
cp .env.example .env
# Отредактируйте .env с вашими параметрами:
# - OutputPath: "." (будет /app в контейнере)
# - Реальные Tenant ID, Client ID, Secret
# - Реальные цены лицензий
```

### 2. Создать директорию для выходных файлов

```bash
mkdir -p output
```

## Запуск

### Локально (разработка)

```bash
# Убедитесь что редактировали .env с вашими параметрами
.\ULI.ps1
```

### В Docker (production)

```bash
docker compose up --build

# В фоне
docker compose up -d --build

# Просмотр логов
docker compose logs -f

# Остановка
docker compose down
```

## Как работает

1. **Локально** (пример: `.\ULI.ps1`)
   - Скрипт ищет `.env`
   - Загружает конфигурацию
   - Сохраняет файлы в текущую директорию

2. **В Docker** (пример: `docker compose up`)
   - docker-compose загружает переменные из `.env`
   - Контейнер копирует `.env` в `/app`
   - Скрипт внутри контейнера ищет `.env`
   - cron запускает скрипт ежедневно в 22:00 UTC
   - Выходные файлы сохраняются в volume `/app/output`

## Остановка

```bash
docker compose down
```

## Переменные окружения

### .env (многострочный — для локальной разработки)

```jsonc
CONFIG_JSON={
  "OutputPath": ".",           // Текущая директория скрипта
  "FileName": "MicrosoftULI.xlsx",
  "LogPath": "ULI.log",
  "SummarySheetName": "Summary",
  "MainSheetName": "Users",
  "SharePoint": {
    "SiteUrl": "eigcom.sharepoint.com",
    "Library": "Documents",
    "Folder": "IT Team/Licenses/"
  },
  "Tenants": [
    {
      "Name": "...",
      "TenantId": "...",
      "ClientId": "...",
      "ClientSecret": "..."
    }
  ]
}

LICENSE_PRICES_JSON={
  "Microsoft 365 Business Basic": 17.25,
  "Microsoft 365 Business Standard": 15.00,
  ...
}
```

### .env (однострочный JSON)

```
CONFIG_JSON={"OutputPath":".","FileName":"MicrosoftULI.xlsx","LogPath":"ULI.log","SummarySheetName":"Summary","MainSheetName":"Users","SharePoint":{"SiteUrl":"eigcom.sharepoint.com","Library":"Documents","Folder":"IT Team/Licenses/"},"Tenants":[{"Name":"...","TenantId":"...","ClientId":"...","ClientSecret":"..."}]}
LICENSE_PRICES_JSON={"Microsoft 365 Business Basic":17.25,"Microsoft 365 Business Standard":15.00,...}
```

**💡 Совет:** редактируйте `.env` как один файл; значения должны быть на одной строке.

## Изменение расписания

Отредактируйте строку в **Dockerfile**:

```dockerfile
# Текущее расписание: 22:00 каждый день
RUN echo "0 22 * * * pwsh -File /app/ULI.ps1 >> /app/output/cron.log 2>&1" > /etc/crontabs/root

# Примеры:
# 0 9 * * *     - 09:00 каждый день
# 0 */6 * * *   - каждые 6 часов
# 0 0 * * 0     - каждый воскресенье в 00:00
```

После изменения пересоберите контейнер:

```bash
docker compose up --build
```

## Troubleshooting

### Ошибка: ".env file not found"

```bash
ls -la .env    # Проверьте, что файл существует
docker compose exec microsoft-uli ls -la /app   # Проверьте в контейнере
```

### Логи не обновляются

```bash
docker compose logs -f microsoft-uli   # Смотрите логи контейнера
tail -f output/cron.log               # Или напрямую
```

### Файлы Excel не создаются

```bash
# Проверьте, что директория output монтирована
docker inspect MicrosoftULI | grep -A 10 Mounts

# Проверьте логи ошибок
cat output/ULI.log
```

## Production развёртывание

1. Используйте переменные окружения вместо .env:
```bash
docker run -e CONFIG_JSON='...' -e LICENSE_PRICES_JSON='...' ...
```

2. Или используйте secrets в Docker Swarm/Kubernetes

3. Регулярно извлекайте файлы из output:
```bash
docker cp MicrosoftULI:/app/output ./backup-$(date +%Y%m%d)
```
