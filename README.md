This Google Apps Script automates the process of syncing  **linnworks data** from the **Linnworks API** to a **Microsoft SQL Server** database. It ensures only new records are inserted, preventing duplicates and maintaining an up-to-date database mirror of Linnworks data

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
## ⚙️ Functionality

- Authenticates with Linnworks API using application credentials.
- Fetches purchase orders in a **paginated** fashion from a specific date range (`2020-01-01` to now).
- Filters out records that already exist in the target SQL Server table (`pkPurchaseID` used as the unique key).
- Handles **duplicate detection** inside the API data and logs them.
- Converts API timestamps safely to **SQL-compatible `TIMESTAMP`** format.
- Inserts new data into the SQL Server table in **batch mode** using JDBC.
- Supports **rollback** if the insertion fails to ensure data consistency.
- Logs all key actions for debugging and transparency.
