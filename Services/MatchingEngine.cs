using Cafe24ShipmentManager.Data;
using Cafe24ShipmentManager.Models;

namespace Cafe24ShipmentManager.Services;

public class MatchingEngine
{
    private readonly DatabaseManager _db;
    private readonly AppLogger _log;

    public MatchingEngine(DatabaseManager db, AppLogger logger)
    {
        _db = db;
        _log = logger;
    }

    /// <summary>
    /// 출고정보 행 목록과 캐시된 Cafe24 주문을 매칭
    /// </summary>
    public List<MatchResult> ExecuteMatching(List<ShipmentSourceRow> sourceRows)
    {
        var results = new List<MatchResult>();
        var allOrders = _db.GetAllCachedOrders();
        var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        // 휴대폰 → 주문 인덱스
        var phoneIndex = new Dictionary<string, List<Cafe24Order>>();
        foreach (var o in allOrders)
        {
            var phone = o.RecipientCellPhone;
            if (!string.IsNullOrEmpty(phone))
            {
                if (!phoneIndex.ContainsKey(phone))
                    phoneIndex[phone] = new();
                phoneIndex[phone].Add(o);
            }
            // 일반 전화번호도 인덱싱
            if (!string.IsNullOrEmpty(o.RecipientPhone) && o.RecipientPhone != phone)
            {
                if (!phoneIndex.ContainsKey(o.RecipientPhone))
                    phoneIndex[o.RecipientPhone] = new();
                phoneIndex[o.RecipientPhone].Add(o);
            }
        }

        int exactCount = 0, candidateCount = 0, noMatchCount = 0;

        foreach (var src in sourceRows)
        {
            if (string.IsNullOrEmpty(src.RecipientPhone))
            {
                // 전화번호 없으면 매칭 불가
                results.Add(new MatchResult
                {
                    SourceRowId = src.Id,
                    Confidence = "none",
                    MatchStatus = "unmatched",
                    CreatedAt = now,
                    SourcePhone = src.RecipientPhone,
                    SourceName = src.RecipientName,
                    SourceTracking = src.TrackingNumber
                });
                noMatchCount++;
                continue;
            }

            var candidates = phoneIndex.GetValueOrDefault(src.RecipientPhone, new());

            if (candidates.Count == 0)
            {
                results.Add(new MatchResult
                {
                    SourceRowId = src.Id,
                    Confidence = "none",
                    MatchStatus = "unmatched",
                    CreatedAt = now,
                    SourcePhone = src.RecipientPhone,
                    SourceName = src.RecipientName,
                    SourceTracking = src.TrackingNumber
                });
                noMatchCount++;
            }
            else if (candidates.Count == 1)
            {
                var c = candidates[0];
                var confidence = DetermineConfidence(src, c);
                results.Add(new MatchResult
                {
                    SourceRowId = src.Id,
                    Cafe24OrderCacheId = c.Id,
                    Cafe24OrderId = c.OrderId,
                    Cafe24OrderItemCode = c.OrderItemCode,
                    Confidence = confidence,
                    MatchStatus = confidence == "exact" ? "auto_confirmed" : "pending",
                    ChosenByUser = false,
                    CreatedAt = now,
                    SourcePhone = src.RecipientPhone,
                    SourceName = src.RecipientName,
                    SourceTracking = src.TrackingNumber,
                    OrderPhone = c.RecipientCellPhone,
                    OrderName = c.RecipientName,
                    OrderProduct = c.ProductName
                });
                exactCount++;
            }
            else
            {
                // 다중 후보: 이름까지 매칭 시도
                var nameMatches = candidates
                    .Where(c => !string.IsNullOrEmpty(src.RecipientName) &&
                                c.RecipientName.Contains(src.RecipientName))
                    .ToList();

                if (nameMatches.Count == 1)
                {
                    var c = nameMatches[0];
                    results.Add(new MatchResult
                    {
                        SourceRowId = src.Id,
                        Cafe24OrderCacheId = c.Id,
                        Cafe24OrderId = c.OrderId,
                        Cafe24OrderItemCode = c.OrderItemCode,
                        Confidence = "exact",
                        MatchStatus = "auto_confirmed",
                        ChosenByUser = false,
                        CreatedAt = now,
                        SourcePhone = src.RecipientPhone,
                        SourceName = src.RecipientName,
                        SourceTracking = src.TrackingNumber,
                        OrderPhone = c.RecipientCellPhone,
                        OrderName = c.RecipientName,
                        OrderProduct = c.ProductName
                    });
                    exactCount++;
                }
                else
                {
                    // 후보 여러 개 → 사용자 선택 필요
                    foreach (var c in candidates)
                    {
                        results.Add(new MatchResult
                        {
                            SourceRowId = src.Id,
                            Cafe24OrderCacheId = c.Id,
                            Cafe24OrderId = c.OrderId,
                            Cafe24OrderItemCode = c.OrderItemCode,
                            Confidence = "candidate",
                            MatchStatus = "pending",
                            ChosenByUser = false,
                            CreatedAt = now,
                            SourcePhone = src.RecipientPhone,
                            SourceName = src.RecipientName,
                            SourceTracking = src.TrackingNumber,
                            OrderPhone = c.RecipientCellPhone,
                            OrderName = c.RecipientName,
                            OrderProduct = c.ProductName
                        });
                    }
                    candidateCount++;
                }
            }
        }

        _log.Info($"매칭 완료: 확정 {exactCount}, 후보선택필요 {candidateCount}, 미매칭 {noMatchCount}");
        return results;
    }

    private string DetermineConfidence(ShipmentSourceRow src, Cafe24Order order)
    {
        // 전화번호 일치(기본)
        bool phoneMatch = src.RecipientPhone == order.RecipientCellPhone ||
                          src.RecipientPhone == order.RecipientPhone;
        if (!phoneMatch) return "none";

        // 이름도 일치하면 exact
        if (!string.IsNullOrEmpty(src.RecipientName) &&
            order.RecipientName.Contains(src.RecipientName))
            return "exact";

        // 전화번호만 일치
        return "probable";
    }
}
