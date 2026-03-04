namespace Cafe24ShipmentManager.Models;

public class ShipmentSourceRow
{
    public long Id { get; set; }
    public string SourceRowKey { get; set; } = "";  // dedupe key
    public string VendorName { get; set; } = "";     // C열: 발주사명
    public string TrackingNumber { get; set; } = ""; // L열: 송장번호
    public string RecipientPhone { get; set; } = ""; // 정규화된 휴대폰
    public string RecipientName { get; set; } = "";  // F열: 수령인명
    public string ProductCode { get; set; } = "";    // B열: 상품코드
    public string OrderDate { get; set; } = "";      // D열: 발주일
    public string ShippingCompany { get; set; } = "";
    public string RawData { get; set; } = "";        // JSON: 원본 행 전체
    public string ProcessStatus { get; set; } = "pending"; // pending/matched/pushed/failed
    public string ImportedAt { get; set; } = "";
    public string? MatchedOrderId { get; set; }
}
