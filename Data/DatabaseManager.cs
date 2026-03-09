using Microsoft.Data.Sqlite;
using Dapper;
using Cafe24ShipmentManager.Models;

namespace Cafe24ShipmentManager.Data;

public class DatabaseManager
{
    private readonly string _connectionString;

    public DatabaseManager(string connectionString)
    {
        _connectionString = connectionString;
        InitializeDatabase();
    }

    private SqliteConnection GetConnection()
    {
        var conn = new SqliteConnection(_connectionString);
        conn.Open();
        return conn;
    }

    private void InitializeDatabase()
    {
        using var conn = GetConnection();
        conn.Execute(@"
            CREATE TABLE IF NOT EXISTS shipment_source_rows (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                SourceRowKey TEXT NOT NULL UNIQUE,
                VendorName TEXT NOT NULL,
                TrackingNumber TEXT NOT NULL,
                RecipientPhone TEXT NOT NULL,
                RecipientName TEXT DEFAULT '',
                ProductCode TEXT DEFAULT '',
                OrderDate TEXT DEFAULT '',
                ShippingCompany TEXT DEFAULT '',
                RawData TEXT DEFAULT '',
                ProcessStatus TEXT DEFAULT 'pending',
                ImportedAt TEXT NOT NULL,
                MatchedOrderId TEXT
            );

            CREATE TABLE IF NOT EXISTS cafe24_orders_cache (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                OrderId TEXT NOT NULL,
                OrderItemCode TEXT DEFAULT '',
                RecipientPhone TEXT DEFAULT '',
                RecipientName TEXT DEFAULT '',
                RecipientCellPhone TEXT DEFAULT '',
                OrderStatus TEXT DEFAULT '',
                ProductName TEXT DEFAULT '',
                OrderAmount REAL DEFAULT 0,
                Quantity INTEGER DEFAULT 0,
                OrderDate TEXT DEFAULT '',
                ShippingCode TEXT DEFAULT '',
                RawJson TEXT DEFAULT '',
                CachedAt TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS match_results (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                SourceRowId INTEGER NOT NULL,
                Cafe24OrderCacheId INTEGER DEFAULT 0,
                Cafe24OrderId TEXT DEFAULT '',
                Cafe24OrderItemCode TEXT DEFAULT '',
                Confidence TEXT DEFAULT 'none',
                MatchStatus TEXT DEFAULT 'pending',
                ChosenByUser INTEGER DEFAULT 0,
                CreatedAt TEXT NOT NULL,
                FOREIGN KEY (SourceRowId) REFERENCES shipment_source_rows(Id)
            );

            CREATE TABLE IF NOT EXISTS push_log (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                MatchResultId INTEGER NOT NULL,
                Cafe24OrderId TEXT DEFAULT '',
                RequestBody TEXT DEFAULT '',
                ResponseBody TEXT DEFAULT '',
                HttpStatusCode INTEGER DEFAULT 0,
                Result TEXT DEFAULT '',
                ErrorMessage TEXT DEFAULT '',
                PushedAt TEXT NOT NULL,
                FOREIGN KEY (MatchResultId) REFERENCES match_results(Id)
            );

            CREATE INDEX IF NOT EXISTS idx_source_vendor ON shipment_source_rows(VendorName);
            CREATE INDEX IF NOT EXISTS idx_source_phone ON shipment_source_rows(RecipientPhone);
            CREATE INDEX IF NOT EXISTS idx_cache_phone ON cafe24_orders_cache(RecipientCellPhone);

            CREATE TABLE IF NOT EXISTS stock_inventory_cache (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                ProductCode TEXT DEFAULT '',
                OrderCode TEXT DEFAULT '',
                Supplier TEXT DEFAULT '',
                ImportCostRaw TEXT DEFAULT '',
                SupplyPriceRaw TEXT DEFAULT '',
                RetailPriceRaw TEXT DEFAULT '',
                InboundRaw TEXT DEFAULT '',
                SoldRaw TEXT DEFAULT '',
                TwoMonthRaw TEXT DEFAULT '',
                OneMonthRaw TEXT DEFAULT '',
                ThisMonthRaw TEXT DEFAULT '',
                StockRaw TEXT DEFAULT '',
                BuyLink TEXT DEFAULT '',
                OptionName TEXT DEFAULT '',
                ImportedAt TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_match_source ON match_results(SourceRowId);
            CREATE INDEX IF NOT EXISTS idx_stock_supplier ON stock_inventory_cache(Supplier);
            CREATE INDEX IF NOT EXISTS idx_stock_product ON stock_inventory_cache(ProductCode);

            CREATE TABLE IF NOT EXISTS stock_order_headers (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                BaseCodeA TEXT DEFAULT '',
                SiteUrl TEXT DEFAULT '',
                TotalQty INTEGER DEFAULT 0,
                TotalAmountYuan REAL DEFAULT 0,
                ItemCount INTEGER DEFAULT 0,
                OrderedAt TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS stock_order_lines (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                HeaderId INTEGER NOT NULL,
                ProductCode TEXT DEFAULT '',
                ImportDetail TEXT DEFAULT '',
                OptionText TEXT DEFAULT '',
                OrderQty INTEGER DEFAULT 0,
                UnitYuan REAL DEFAULT 0,
                AmountYuan REAL DEFAULT 0,
                FOREIGN KEY (HeaderId) REFERENCES stock_order_headers(Id)
            );

            CREATE INDEX IF NOT EXISTS idx_soh_ordered ON stock_order_headers(OrderedAt);
            CREATE INDEX IF NOT EXISTS idx_sol_header ON stock_order_lines(HeaderId);
            CREATE INDEX IF NOT EXISTS idx_sol_product ON stock_order_lines(ProductCode);
        ");
    }

    // ── ShipmentSourceRow ──
    public long InsertSourceRow(ShipmentSourceRow row)
    {
        using var conn = GetConnection();
        return conn.ExecuteScalar<long>(@"
            INSERT OR IGNORE INTO shipment_source_rows
                (SourceRowKey, VendorName, TrackingNumber, RecipientPhone, RecipientName, ProductCode, OrderDate, ShippingCompany, RawData, ProcessStatus, ImportedAt)
            VALUES
                (@SourceRowKey, @VendorName, @TrackingNumber, @RecipientPhone, @RecipientName, @ProductCode, @OrderDate, @ShippingCompany, @RawData, @ProcessStatus, @ImportedAt);
            SELECT last_insert_rowid();", row);
    }

    /// <summary>
    /// 필터된 행만 단일 트랜잭션으로 DB 저장 (INSERT OR IGNORE + ID 할당)
    /// </summary>
    public void BulkInsertSourceRows(List<ShipmentSourceRow> rows)
    {
        using var conn = GetConnection();
        using var tx = conn.BeginTransaction();

        foreach (var row in rows)
        {
            var id = conn.ExecuteScalar<long>(@"
                INSERT OR IGNORE INTO shipment_source_rows
                    (SourceRowKey, VendorName, TrackingNumber, RecipientPhone, RecipientName, ProductCode, OrderDate, ShippingCompany, RawData, ProcessStatus, ImportedAt)
                VALUES
                    (@SourceRowKey, @VendorName, @TrackingNumber, @RecipientPhone, @RecipientName, @ProductCode, @OrderDate, @ShippingCompany, @RawData, @ProcessStatus, @ImportedAt);
                SELECT last_insert_rowid();", row, transaction: tx);

            if (id > 0)
                row.Id = id;
            else
            {
                var existingId = conn.ExecuteScalar<long?>(
                    "SELECT Id FROM shipment_source_rows WHERE SourceRowKey = @SourceRowKey",
                    new { row.SourceRowKey }, transaction: tx);
                if (existingId.HasValue) row.Id = existingId.Value;
            }
        }

        tx.Commit();
    }

    public List<ShipmentSourceRow> GetSourceRowsByVendor(string vendor)
    {
        using var conn = GetConnection();
        return conn.Query<ShipmentSourceRow>(
            "SELECT * FROM shipment_source_rows WHERE VendorName = @vendor ORDER BY Id",
            new { vendor }).ToList();
    }

    public void UpdateSourceRowStatus(long id, string status, string? matchedOrderId = null)
    {
        using var conn = GetConnection();
        conn.Execute(
            "UPDATE shipment_source_rows SET ProcessStatus = @status, MatchedOrderId = @matchedOrderId WHERE Id = @id",
            new { id, status, matchedOrderId });
    }

    public List<string> GetDistinctVendors()
    {
        using var conn = GetConnection();
        return conn.Query<string>("SELECT DISTINCT VendorName FROM shipment_source_rows ORDER BY VendorName").ToList();
    }

    // ── Cafe24 Order Cache ──
    public void ClearOrderCache()
    {
        using var conn = GetConnection();
        conn.Execute("DELETE FROM cafe24_orders_cache");
    }

    public void InsertOrderCache(Cafe24Order order)
    {
        using var conn = GetConnection();
        conn.Execute(@"
            INSERT INTO cafe24_orders_cache
                (OrderId, OrderItemCode, RecipientPhone, RecipientName, RecipientCellPhone, OrderStatus, ProductName, OrderAmount, Quantity, OrderDate, ShippingCode, RawJson, CachedAt)
            VALUES
                (@OrderId, @OrderItemCode, @RecipientPhone, @RecipientName, @RecipientCellPhone, @OrderStatus, @ProductName, @OrderAmount, @Quantity, @OrderDate, @ShippingCode, @RawJson, @CachedAt)", order);
    }

    public List<Cafe24Order> GetCachedOrdersByPhone(string phone)
    {
        using var conn = GetConnection();
        return conn.Query<Cafe24Order>(
            "SELECT * FROM cafe24_orders_cache WHERE RecipientCellPhone = @phone OR RecipientPhone = @phone",
            new { phone }).ToList();
    }

    public List<Cafe24Order> GetAllCachedOrders()
    {
        using var conn = GetConnection();
        return conn.Query<Cafe24Order>("SELECT * FROM cafe24_orders_cache").ToList();
    }

    // ── Match Results ──
    public long InsertMatchResult(MatchResult mr)
    {
        using var conn = GetConnection();
        return conn.ExecuteScalar<long>(@"
            INSERT INTO match_results
                (SourceRowId, Cafe24OrderCacheId, Cafe24OrderId, Cafe24OrderItemCode, Confidence, MatchStatus, ChosenByUser, CreatedAt)
            VALUES
                (@SourceRowId, @Cafe24OrderCacheId, @Cafe24OrderId, @Cafe24OrderItemCode, @Confidence, @MatchStatus, @ChosenByUser, @CreatedAt);
            SELECT last_insert_rowid();", mr);
    }

    public List<MatchResult> GetMatchResultsBySourceIds(List<long> sourceIds)
    {
        if (sourceIds.Count == 0) return new();
        using var conn = GetConnection();
        var ids = string.Join(",", sourceIds);
        return conn.Query<MatchResult>(
            $"SELECT * FROM match_results WHERE SourceRowId IN ({ids}) ORDER BY SourceRowId, Confidence DESC").ToList();
    }

    public void UpdateMatchStatus(long id, string status, bool chosenByUser)
    {
        using var conn = GetConnection();
        conn.Execute(
            "UPDATE match_results SET MatchStatus = @status, ChosenByUser = @chosenByUser WHERE Id = @id",
            new { id, status, chosenByUser });
    }

    public void DeleteMatchResultsBySourceIds(List<long> sourceIds)
    {
        if (sourceIds.Count == 0) return;
        using var conn = GetConnection();
        var ids = string.Join(",", sourceIds);
        conn.Execute($"DELETE FROM match_results WHERE SourceRowId IN ({ids})");
    }


    // ── Stock Inventory Cache ──
    public void ReplaceStockInventoryCache(IEnumerable<object> rows)
    {
        using var conn = GetConnection();
        using var tx = conn.BeginTransaction();

        conn.Execute("DELETE FROM stock_inventory_cache", transaction: tx);

        conn.Execute(@"
            INSERT INTO stock_inventory_cache
                (ProductCode, OrderCode, Supplier, ImportCostRaw, SupplyPriceRaw, RetailPriceRaw,
                 InboundRaw, SoldRaw, TwoMonthRaw, OneMonthRaw, ThisMonthRaw, StockRaw,
                 BuyLink, OptionName, ImportedAt)
            VALUES
                (@ProductCode, @OrderCode, @Supplier, @ImportCostRaw, @SupplyPriceRaw, @RetailPriceRaw,
                 @InboundRaw, @SoldRaw, @TwoMonthRaw, @OneMonthRaw, @ThisMonthRaw, @StockRaw,
                 @BuyLink, @OptionName, @ImportedAt)", rows, transaction: tx);

        tx.Commit();
    }
    // ── Push Log ──
    public void InsertPushLog(PushLog log)
    {
        using var conn = GetConnection();
        conn.Execute(@"
            INSERT INTO push_log
                (MatchResultId, Cafe24OrderId, RequestBody, ResponseBody, HttpStatusCode, Result, ErrorMessage, PushedAt)
            VALUES
                (@MatchResultId, @Cafe24OrderId, @RequestBody, @ResponseBody, @HttpStatusCode, @Result, @ErrorMessage, @PushedAt)", log);
    }

    public List<PushLog> GetPushLogs(int limit = 500)
    {
        using var conn = GetConnection();
        return conn.Query<PushLog>(
            "SELECT * FROM push_log ORDER BY Id DESC LIMIT @limit", new { limit }).ToList();
    }

    // ── Stock Order History ──
    public long InsertStockOrder(StockOrderHeader header, List<StockOrderLine> lines)
    {
        using var conn = GetConnection();
        using var tx = conn.BeginTransaction();

        var headerId = conn.ExecuteScalar<long>(@"
            INSERT INTO stock_order_headers (BaseCodeA, SiteUrl, TotalQty, TotalAmountYuan, ItemCount, OrderedAt)
            VALUES (@BaseCodeA, @SiteUrl, @TotalQty, @TotalAmountYuan, @ItemCount, @OrderedAt);
            SELECT last_insert_rowid();", header, transaction: tx);

        foreach (var line in lines)
        {
            line.HeaderId = headerId;
            conn.Execute(@"
                INSERT INTO stock_order_lines (HeaderId, ProductCode, ImportDetail, OptionText, OrderQty, UnitYuan, AmountYuan)
                VALUES (@HeaderId, @ProductCode, @ImportDetail, @OptionText, @OrderQty, @UnitYuan, @AmountYuan)", line, transaction: tx);
        }

        tx.Commit();
        return headerId;
    }

    public List<StockOrderHeader> SearchStockOrderHeaders(string? dateFrom, string? dateTo, string? productCodeFilter)
    {
        using var conn = GetConnection();
        var sql = "SELECT DISTINCT h.* FROM stock_order_headers h";
        var where = new List<string>();
        var param = new DynamicParameters();

        if (!string.IsNullOrWhiteSpace(productCodeFilter))
        {
            sql += " INNER JOIN stock_order_lines l ON l.HeaderId = h.Id";
            where.Add("l.ProductCode LIKE @code");
            param.Add("code", $"%{productCodeFilter.Trim()}%");
        }
        if (!string.IsNullOrWhiteSpace(dateFrom))
        {
            where.Add("h.OrderedAt >= @dateFrom");
            param.Add("dateFrom", dateFrom);
        }
        if (!string.IsNullOrWhiteSpace(dateTo))
        {
            where.Add("h.OrderedAt <= @dateTo");
            param.Add("dateTo", dateTo + " 23:59:59");
        }

        if (where.Count > 0)
            sql += " WHERE " + string.Join(" AND ", where);
        sql += " ORDER BY h.OrderedAt DESC";

        return conn.Query<StockOrderHeader>(sql, param).ToList();
    }

    public List<StockOrderLine> GetStockOrderLines(long headerId)
    {
        using var conn = GetConnection();
        return conn.Query<StockOrderLine>(
            "SELECT * FROM stock_order_lines WHERE HeaderId = @headerId ORDER BY Id",
            new { headerId }).ToList();
    }

    public List<TopOrderedProduct> GetTopOrderedProducts(string? dateFrom, string? dateTo, int topN = 20)
    {
        using var conn = GetConnection();
        var where = new List<string>();
        var param = new DynamicParameters();
        param.Add("topN", topN);

        if (!string.IsNullOrWhiteSpace(dateFrom))
        {
            where.Add("h.OrderedAt >= @dateFrom");
            param.Add("dateFrom", dateFrom);
        }
        if (!string.IsNullOrWhiteSpace(dateTo))
        {
            where.Add("h.OrderedAt <= @dateTo");
            param.Add("dateTo", dateTo + " 23:59:59");
        }

        var whereClause = where.Count > 0 ? "WHERE " + string.Join(" AND ", where) : "";

        return conn.Query<TopOrderedProduct>($@"
            SELECT l.ProductCode, SUM(l.OrderQty) AS TotalQty, COUNT(DISTINCT l.HeaderId) AS OrderCount
            FROM stock_order_lines l
            INNER JOIN stock_order_headers h ON h.Id = l.HeaderId
            {whereClause}
            GROUP BY l.ProductCode
            ORDER BY TotalQty DESC
            LIMIT @topN", param).ToList();
    }

    public List<OptionMonthlyRecord> GetOptionMonthlyBreakdown(string baseCodeA)
    {
        using var conn = GetConnection();
        return conn.Query<OptionMonthlyRecord>(@"
            SELECT l.ProductCode, l.OptionText, SUBSTR(h.OrderedAt, 1, 7) AS Month, SUM(l.OrderQty) AS TotalQty
            FROM stock_order_lines l
            INNER JOIN stock_order_headers h ON h.Id = l.HeaderId
            WHERE h.BaseCodeA = @baseCodeA
            GROUP BY l.ProductCode, l.OptionText, SUBSTR(h.OrderedAt, 1, 7)
            ORDER BY l.ProductCode, Month", new { baseCodeA }).ToList();
    }

    public List<MonthlyOrderTrend> GetMonthlyOrderTrend(string? productCode, string? dateFrom, string? dateTo)
    {
        using var conn = GetConnection();
        var where = new List<string>();
        var param = new DynamicParameters();

        if (!string.IsNullOrWhiteSpace(productCode))
        {
            where.Add("l.ProductCode LIKE @code");
            param.Add("code", $"%{productCode.Trim()}%");
        }
        if (!string.IsNullOrWhiteSpace(dateFrom))
        {
            where.Add("h.OrderedAt >= @dateFrom");
            param.Add("dateFrom", dateFrom);
        }
        if (!string.IsNullOrWhiteSpace(dateTo))
        {
            where.Add("h.OrderedAt <= @dateTo");
            param.Add("dateTo", dateTo + " 23:59:59");
        }

        var whereClause = where.Count > 0 ? "WHERE " + string.Join(" AND ", where) : "";

        return conn.Query<MonthlyOrderTrend>($@"
            SELECT SUBSTR(h.OrderedAt, 1, 7) AS Month, SUM(l.OrderQty) AS TotalQty, SUM(l.AmountYuan) AS TotalAmountYuan
            FROM stock_order_lines l
            INNER JOIN stock_order_headers h ON h.Id = l.HeaderId
            {whereClause}
            GROUP BY SUBSTR(h.OrderedAt, 1, 7)
            ORDER BY Month", param).ToList();
    }
}


