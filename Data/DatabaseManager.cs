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
            CREATE INDEX IF NOT EXISTS idx_match_source ON match_results(SourceRowId);
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
}
