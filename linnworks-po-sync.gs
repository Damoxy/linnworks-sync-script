/**
 * Main entry point: Fetches Linnworks purchase orders and inserts new ones into MS SQL.
 */
function main_stocklevel() {
  const allData = getLinnworksPurchaseOrdersData();
  if (allData.length === 0) {
    Logger.log('No Linnworks data found to insert.');
    return;
  }
  Logger.log(`Fetched ${allData.length} purchase orders from Linnworks.`);
  pushLinnworksDataToMSSQL(allData);
}

/**
 * Authenticates with the Linnworks API and retrieves an access token.
 * @returns {string|null} Access token if successful, otherwise null.
 */
function getLinnworksAccessToken() {
  const url = "https://api.linnworks.net/api/Auth/AuthorizeByApplication";

  const payload = {
    ApplicationId: "<YOUR_APP_ID>",
    ApplicationSecret: "<YOUR_APP_SECRET>",
    Token: "<YOUR_AUTH_TOKEN>"
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    const token = data.Token || data.AccessToken;

    if (!token) {
      Logger.log("Access token not found in response: " + response.getContentText());
      return null;
    }

    Logger.log("Fetched Linnworks access token successfully.");
    return token;

  } catch (e) {
    Logger.log("Error fetching Linnworks access token: " + e);
    return null;
  }
}

/**
 * Retrieves paginated purchase order data from Linnworks API.
 * @returns {Array<Object>} Array of purchase order records.
 */
function getLinnworksPurchaseOrdersData() {
  const accessToken = getLinnworksAccessToken();
  if (!accessToken) {
    Logger.log("Failed to retrieve access token.");
    return [];
  }

  const payload = {
    searchParameters: {
      DateFrom: '2020-01-01T00:00:00',
      DateTo: new Date().toISOString()
    },
    entriesPerPage: 100,
    pageNumber: 1
  };

  const headers = {
    'Authorization': accessToken,
    'Accept': 'application/json',
    'Content-Type': 'application/json'
  };

  const options = {
    method: 'POST',
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let allData = [];
  let currentPage = 1;
  let totalPages = 1;

  do {
    payload.pageNumber = currentPage;
    options.payload = JSON.stringify(payload);

    const response = UrlFetchApp.fetch('https://eu-ext.linnworks.net/api/PurchaseOrder/Search_PurchaseOrders2', options);
    const json = JSON.parse(response.getContentText());

    if (json.Result && json.Result.length > 0) {
      allData = allData.concat(json.Result);
    }

    totalPages = json.TotalPages || 1;
    currentPage++;
  } while (currentPage <= totalPages);

  return allData;
}

/**
 * Converts a date string to a JDBC timestamp, or null if invalid.
 * @param {string|null|undefined} value - Date string.
 * @returns {Jdbc.Timestamp|null}
 */
function safeDate(value) {
  if (!value) return null;
  const jsDate = new Date(value);
  if (isNaN(jsDate.getTime())) return null;
  return Jdbc.newTimestamp(jsDate.getTime());
}

/**
 * Inserts new Linnworks purchase orders into an MS SQL Server table using JDBC.
 * @param {Array<Object>} allData - Array of purchase order data.
 */
function pushLinnworksDataToMSSQL(allData) {
  const instanceUrl = 'jdbc:sqlserver://<DB_HOST>:1433;databaseName=<DB_NAME>';
  const user = '<DB_USERNAME>';
  const password = '<DB_PASSWORD>';

  let conn;
  try {
    conn = Jdbc.getConnection(instanceUrl, user, password);
    conn.setAutoCommit(false);

    // Step 1: Get existing primary keys
    const existingIds = new Set();
    const rs = conn.createStatement().executeQuery('SELECT pkPurchaseID FROM <SCHEMA>.<TABLE_NAME>');
    while (rs.next()) {
      const id = rs.getString('pkPurchaseID');
      if (id) existingIds.add(id.trim().toLowerCase());
    }
    rs.close();
    Logger.log(`Fetched ${existingIds.size} existing pkPurchaseIDs from database.`);

    // Step 2: Check for duplicates in fetched data
    const seen = new Set();
    const duplicates = [];
    allData.forEach(row => {
      if (!row.pkPurchaseID) return;
      const id = row.pkPurchaseID.trim().toLowerCase();
      if (seen.has(id)) duplicates.push(id);
      else seen.add(id);
    });
    if (duplicates.length > 0) {
      Logger.log(`Duplicate pkPurchaseIDs found in API data: ${JSON.stringify(duplicates)}`);
    } else {
      Logger.log('No duplicate pkPurchaseIDs found in API data.');
    }

    // Step 3: Filter for new records
    const newRecords = allData.filter(row => {
      if (!row.pkPurchaseID) return false;
      return !existingIds.has(row.pkPurchaseID.trim().toLowerCase());
    });

    Logger.log(`Total records from API: ${allData.length}`);
    Logger.log(`New records to insert: ${newRecords.length}`);

    if (newRecords.length === 0) {
      Logger.log('No new records to insert. All purchase orders already exist.');
      return;
    }

    // Step 4: Prepare and execute insert
    const sql = `INSERT INTO <SCHEMA>.<TABLE_NAME>
      (
        pkPurchaseID, fkSupplierId, fkLocationId, ExternalInvoiceNumber, Status, Currency, SupplierReferenceNumber,
        Locked, LineCount, DeliveredLinesCount, UnitAmountTaxIncludedType, DateOfPurchase, DateOfDelivery, QuotedDeliveryDate,
        PostagePaid, TotalCost, taxPaid, ShippingTaxRate, ConversionRate, ConvertedShippingCost,
        ConvertedShippingTax, ConvertedOtherCost, ConvertedOtherTax, ConvertedGrandTotal
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;

    const stmt = conn.prepareStatement(sql);

    for (let i = 0; i < newRecords.length; i++) {
      const row = newRecords[i];

      stmt.setString(1, row.pkPurchaseID || null);
      stmt.setString(2, row.fkSupplierId || null);
      stmt.setString(3, row.fkLocationId || null);
      stmt.setString(4, row.ExternalInvoiceNumber || null);
      stmt.setString(5, row.Status || null);
      stmt.setString(6, row.Currency || null);
      stmt.setString(7, row.SupplierReferenceNumber || null);
      stmt.setInt(8, row.Locked ? 1 : 0);
      stmt.setInt(9, row.LineCount || 0);
      stmt.setInt(10, row.DeliveredLinesCount || 0);
      stmt.setString(11, row.UnitAmountTaxIncludedType != null ? String(row.UnitAmountTaxIncludedType) : '');
      stmt.setTimestamp(12, safeDate(row.DateOfPurchase));
      stmt.setTimestamp(13, safeDate(row.DateOfDelivery));
      stmt.setTimestamp(14, safeDate(row.QuotedDeliveryDate));
      stmt.setDouble(15, row.PostagePaid != null ? row.PostagePaid : 0);
      stmt.setDouble(16, row.TotalCost != null ? row.TotalCost : 0);
      stmt.setDouble(17, row.taxPaid != null ? row.taxPaid : 0);
      stmt.setDouble(18, row.ShippingTaxRate != null ? row.ShippingTaxRate : 0);
      stmt.setDouble(19, row.ConversionRate != null ? row.ConversionRate : 0);
      stmt.setDouble(20, row.ConvertedShippingCost != null ? row.ConvertedShippingCost : 0);
      stmt.setDouble(21, row.ConvertedShippingTax != null ? row.ConvertedShippingTax : 0);
      stmt.setDouble(22, row.ConvertedOtherCost != null ? row.ConvertedOtherCost : 0);
      stmt.setDouble(23, row.ConvertedOtherTax != null ? row.ConvertedOtherTax : 0);
      stmt.setDouble(24, row.ConvertedGrandTotal != null ? row.ConvertedGrandTotal : 0);

      stmt.addBatch();

      // Optional: batch insert every 100 records
      if ((i + 1) % 100 === 0) {
        stmt.executeBatch();
      }
    }

    stmt.executeBatch();
    conn.commit();
    stmt.close();

    Logger.log(`Inserted ${newRecords.length} new purchase orders successfully.`);

  } catch (e) {
    Logger.log('Error during MSSQL insert: ' + e);
    if (conn) {
      try {
        conn.rollback();
      } catch (rollbackErr) {
        Logger.log('Rollback error: ' + rollbackErr);
      }
    }
  } finally {
    if (conn) conn.close();
  }
}
