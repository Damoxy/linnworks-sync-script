## 🔧 Setup

### 1. Google Apps Script Environment

1. Open [Google Apps Script](https://script.google.com/)
2. Create a new project and paste the contents of `script.gs`
3. Set up **JDBC** access and enable the `UrlFetchApp` service

### 2. Fill in Configuration

Replace the placeholder values in the script:

| Placeholder           | Description                          |
|------------------------|--------------------------------------|
| `<YOUR_APP_ID>`        | Linnworks application ID             |
| `<YOUR_APP_SECRET>`    | Linnworks application secret         |
| `<YOUR_AUTH_TOKEN>`    | Linnworks auth token                 |
| `<DB_HOST>`            | Your SQL Server hostname             |
| `<DB_NAME>`            | Your database name                   |
| `<DB_USERNAME>`        | DB username                          |
| `<DB_PASSWORD>`        | DB password                          |
| `<SCHEMA>.<TABLE_NAME>`| SQL target table                     |

---
