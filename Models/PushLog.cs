namespace Cafe24ShipmentManager.Models;

public class PushLog
{
    public long Id { get; set; }
    public long MatchResultId { get; set; }
    public string Cafe24OrderId { get; set; } = "";
    public string RequestBody { get; set; } = "";
    public string ResponseBody { get; set; } = "";
    public int HttpStatusCode { get; set; }
    public string Result { get; set; } = ""; // success / fail
    public string ErrorMessage { get; set; } = "";
    public string PushedAt { get; set; } = "";
}
