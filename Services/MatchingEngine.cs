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

    /// <summary>
    /// 역방향 매칭: Cafe24 주문 기준으로 스프레드시트에서 송장 검색
    /// </summary>
    public List<MatchResult> ExecuteReverseMatching(List<Cafe24Order> orders, List<ShipmentSourceRow> sheetRows)
    {
        var results = new List<MatchResult>();
        var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        // 스프레드시트 전화번호 → 행 인덱스
        var phoneIndex = new Dictionary<string, List<ShipmentSourceRow>>();
        foreach (var row in sheetRows)
        {
            var phone = row.RecipientPhone;
            if (string.IsNullOrEmpty(phone)) continue;
            if (!phoneIndex.ContainsKey(phone))
                phoneIndex[phone] = new();
            phoneIndex[phone].Add(row);
        }

        int exactCount = 0, noTrackingCount = 0, candidateCount = 0, noMatchCount = 0;

        foreach (var order in orders)
        {
            var phone = order.RecipientCellPhone;
            var phone2 = order.RecipientPhone;

            // 전화번호로 스프레드시트 행 검색
            var candidates = new List<ShipmentSourceRow>();
            if (!string.IsNullOrEmpty(phone) && phoneIndex.TryGetValue(phone, out var list1))
                candidates.AddRange(list1);
            if (!string.IsNullOrEmpty(phone2) && phone2 != phone && phoneIndex.TryGetValue(phone2, out var list2))
                candidates.AddRange(list2);

            // 중복 제거
            candidates = candidates.DistinctBy(r => r.Id > 0 ? (object)r.Id : r.SourceRowKey).ToList();

            if (candidates.Count == 0)
            {
                // 스프레드시트에 매칭 없음
                results.Add(new MatchResult
                {
                    SourceRowId = 0,
                    Cafe24OrderCacheId = order.Id,
                    Cafe24OrderId = order.OrderId,
                    Cafe24OrderItemCode = order.OrderItemCode,
                    Confidence = "none",
                    MatchStatus = "unmatched",
                    CreatedAt = now,
                    SourcePhone = "",
                    SourceName = "",
                    SourceTracking = "",
                    OrderPhone = order.RecipientCellPhone,
                    OrderName = order.RecipientName,
                    OrderProduct = order.ProductName
                });
                noMatchCount++;
                continue;
            }

            // 송장번호가 있는 후보만 필터
            var withTracking = candidates.Where(r => !string.IsNullOrEmpty(r.TrackingNumber)).ToList();
            var withoutTracking = candidates.Where(r => string.IsNullOrEmpty(r.TrackingNumber)).ToList();

            if (withTracking.Count == 0)
            {
                // 매칭은 됐지만 송장번호가 없음 (재고없음 등)
                var firstMatch = candidates[0];
                results.Add(new MatchResult
                {
                    SourceRowId = firstMatch.Id,
                    Cafe24OrderCacheId = order.Id,
                    Cafe24OrderId = order.OrderId,
                    Cafe24OrderItemCode = order.OrderItemCode,
                    Confidence = "no_tracking",
                    MatchStatus = "no_tracking",
                    CreatedAt = now,
                    SourcePhone = firstMatch.RecipientPhone,
                    SourceName = firstMatch.RecipientName,
                    SourceTracking = "",
                    OrderPhone = order.RecipientCellPhone,
                    OrderName = order.RecipientName,
                    OrderProduct = order.ProductName
                });
                noTrackingCount++;
                continue;
            }

            if (withTracking.Count == 1)
            {
                var src = withTracking[0];
                var confidence = DetermineReverseConfidence(order, src);
                results.Add(new MatchResult
                {
                    SourceRowId = src.Id,
                    Cafe24OrderCacheId = order.Id,
                    Cafe24OrderId = order.OrderId,
                    Cafe24OrderItemCode = order.OrderItemCode,
                    Confidence = confidence,
                    MatchStatus = confidence == "exact" ? "auto_confirmed" : "pending",
                    ChosenByUser = false,
                    CreatedAt = now,
                    SourcePhone = src.RecipientPhone,
                    SourceName = src.RecipientName,
                    SourceTracking = src.TrackingNumber,
                    OrderPhone = order.RecipientCellPhone,
                    OrderName = order.RecipientName,
                    OrderProduct = order.ProductName
                });
                exactCount++;
            }
            else
            {
                // 다중 후보: 이름 매칭 시도
                var nameMatches = withTracking
                    .Where(r => !string.IsNullOrEmpty(order.RecipientName) &&
                                !string.IsNullOrEmpty(r.RecipientName) &&
                                order.RecipientName.Contains(r.RecipientName))
                    .ToList();

                if (nameMatches.Count == 1)
                {
                    var src = nameMatches[0];
                    results.Add(new MatchResult
                    {
                        SourceRowId = src.Id,
                        Cafe24OrderCacheId = order.Id,
                        Cafe24OrderId = order.OrderId,
                        Cafe24OrderItemCode = order.OrderItemCode,
                        Confidence = "exact",
                        MatchStatus = "auto_confirmed",
                        ChosenByUser = false,
                        CreatedAt = now,
                        SourcePhone = src.RecipientPhone,
                        SourceName = src.RecipientName,
                        SourceTracking = src.TrackingNumber,
                        OrderPhone = order.RecipientCellPhone,
                        OrderName = order.RecipientName,
                        OrderProduct = order.ProductName
                    });
                    exactCount++;
                }
                else
                {
                    // 후보 여러 개 → 사용자 선택 필요
                    foreach (var src in withTracking)
                    {
                        results.Add(new MatchResult
                        {
                            SourceRowId = src.Id,
                            Cafe24OrderCacheId = order.Id,
                            Cafe24OrderId = order.OrderId,
                            Cafe24OrderItemCode = order.OrderItemCode,
                            Confidence = "candidate",
                            MatchStatus = "pending",
                            ChosenByUser = false,
                            CreatedAt = now,
                            SourcePhone = src.RecipientPhone,
                            SourceName = src.RecipientName,
                            SourceTracking = src.TrackingNumber,
                            OrderPhone = order.RecipientCellPhone,
                            OrderName = order.RecipientName,
                            OrderProduct = order.ProductName
                        });
                    }
                    candidateCount++;
                }
            }
        }

        _log.Info($"역매칭 완료: 확정 {exactCount}, 후보선택필요 {candidateCount}, 송장없음 {noTrackingCount}, 미매칭 {noMatchCount}");
        return results;
    }

    private string DetermineReverseConfidence(Cafe24Order order, ShipmentSourceRow src)
    {
        // 전화번호 일치(기본)
        bool phoneMatch = src.RecipientPhone == order.RecipientCellPhone ||
                          src.RecipientPhone == order.RecipientPhone;
        if (!phoneMatch) return "none";

        // 이름도 일치하면 exact
        if (!string.IsNullOrEmpty(src.RecipientName) &&
            !string.IsNullOrEmpty(order.RecipientName) &&
            order.RecipientName.Contains(src.RecipientName))
            return "exact";

        return "probable";
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
