namespace Cafe24ShipmentManager.Models;

public class MatchResult
{
    public long Id { get; set; }
    public long SourceRowId { get; set; }
    public long Cafe24OrderCacheId { get; set; }
    public string Cafe24OrderId { get; set; } = "";
    public string Cafe24OrderItemCode { get; set; } = "";
    public string Confidence { get; set; } = "none"; // exact / probable / candidate / none
    public string MatchStatus { get; set; } = "pending"; // pending / confirmed / rejected / pushed / push_failed
    public bool ChosenByUser { get; set; }
    public string CreatedAt { get; set; } = "";

    // Display helpers (not DB columns)
    public string SourcePhone { get; set; } = "";
    public string SourceName { get; set; } = "";
    public string SourceTracking { get; set; } = "";
    public string OrderPhone { get; set; } = "";
    public string OrderName { get; set; } = "";
    public string OrderProduct { get; set; } = "";
}
