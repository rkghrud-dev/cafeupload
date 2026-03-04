namespace Cafe24ShipmentManager.Models;

public class Cafe24Order
{
    public long Id { get; set; }
    public string OrderId { get; set; } = "";
    public string OrderItemCode { get; set; } = "";
    public string RecipientPhone { get; set; } = "";
    public string RecipientName { get; set; } = "";
    public string RecipientCellPhone { get; set; } = "";
    public string OrderStatus { get; set; } = "";
    public string ProductName { get; set; } = "";
    public decimal OrderAmount { get; set; }
    public int Quantity { get; set; }
    public string OrderDate { get; set; } = "";
    public string ShippingCode { get; set; } = "";
    public string RawJson { get; set; } = "";
    public string CachedAt { get; set; } = "";
}
