using System.Net;
using System.Net.Http.Headers;
using System.Text;
using Cafe24ShipmentManager.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Cafe24ShipmentManager.Services;

public class Cafe24Config
{
    public string MallId { get; set; } = "";
    public string AccessToken { get; set; } = "";
    public string ClientId { get; set; } = "";
    public string ClientSecret { get; set; } = "";
    public string RefreshToken { get; set; } = "";
    public string ApiVersion { get; set; } = "2023-03-01";
    public string DefaultShippingCompanyCode { get; set; } = "0019"; // CJ대한통운
    public int OrderFetchDays { get; set; } = 14;
    public string? ConfigFilePath { get; set; }
}

public class Cafe24ApiClient
{
    private readonly HttpClient _http;
    private readonly Cafe24Config _config;
    private readonly AppLogger _log;
    private const int MaxRetries = 3;

    // 택배사 코드 매핑 (Cafe24 기준)
    public static readonly Dictionary<string, string> ShippingCompanyCodes = new()
    {
        { "CJ대한통운", "0019" },
        { "롯데택배", "0016" },
        { "한진택배", "0018" },
        { "로젠택배", "0020" },
        { "우체국택배", "0001" },
        { "경동택배", "0023" },
        { "대신택배", "0022" },
        { "일양로지스", "0011" },
        { "합동택배", "0032" },
        { "건영택배", "0034" },
        { "천일택배", "0017" },
        { "CU편의점택배", "0049" },
        { "GS편의점택배", "0050" },
        { "EMS", "0005" },
        { "DHL", "0004" },
        { "FedEx", "0012" },
        { "UPS", "0015" },
    };

    public Cafe24ApiClient(Cafe24Config config, AppLogger logger)
    {
        _config = config;
        _log = logger;
        _http = new HttpClient
        {
            BaseAddress = new Uri($"https://{config.MallId}.cafe24api.com/api/v2/"),
            Timeout = TimeSpan.FromSeconds(30)
        };
        _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", config.AccessToken);
        _http.DefaultRequestHeaders.Add("X-Cafe24-Api-Version", config.ApiVersion);
        _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
    }

    /// <summary>
    /// 날짜 범위로 주문 목록 조회 (페이징 처리)
    /// </summary>
    /// <summary>
    /// 날짜 범위 + 주문상태로 주문 목록 조회 (페이징 처리)
    /// orderStatus: null이면 전체, "N20"이면 배송준비중 등
    /// </summary>
    public async Task<List<Cafe24Order>> FetchRecentOrders(DateTime startDt, DateTime endDt, IProgress<string>? progress = null, string? orderStatus = null)
    {
        var orders = new List<Cafe24Order>();
        var startDate = startDt.ToString("yyyy-MM-dd");
        var endDate = endDt.ToString("yyyy-MM-dd");
        int offset = 0;
        const int limit = 100;
        int page = 0;

        while (true)
        {
            page++;
            var statusLabel = orderStatus != null ? $", 상태={orderStatus}" : "";
            progress?.Report($"Cafe24 주문 조회 중... (페이지 {page}{statusLabel})");

            var url = $"admin/orders?start_date={startDate}&end_date={endDate}&limit={limit}&offset={offset}";
            if (!string.IsNullOrEmpty(orderStatus))
                url += $"&order_status={orderStatus}";
            var response = await ExecuteWithRetry(() => _http.GetAsync(url));

            if (response == null || !response.IsSuccessStatusCode)
            {
                var body = response != null ? await response.Content.ReadAsStringAsync() : "no response";
                _log.Error($"주문 조회 실패: {response?.StatusCode} - {body}");
                break;
            }

            var json = await response.Content.ReadAsStringAsync();
            var jObj = JObject.Parse(json);
            var ordersArray = jObj["orders"] as JArray;

            if (ordersArray == null || ordersArray.Count == 0)
                break;

            foreach (var o in ordersArray)
            {
                var orderId = o["order_id"]?.ToString() ?? "";

                // items 안의 각 주문 아이템을 개별 행으로
                var items = o["items"] as JArray;
                if (items != null)
                {
                    foreach (var item in items)
                    {
                        orders.Add(ParseOrder(o, item, orderId));
                    }
                }
                else
                {
                    orders.Add(ParseOrder(o, null, orderId));
                }
            }

            if (ordersArray.Count < limit)
                break;

            offset += limit;
        }

        _log.Info($"Cafe24 주문 {orders.Count}건 조회 완료 ({startDate} ~ {endDate})");
        return orders;
    }

    private Cafe24Order ParseOrder(JToken order, JToken? item, string orderId)
    {
        // 수령인 정보는 receiver 필드에 있을 수 있음
        var receiverName = order["receiver_name"]?.ToString()
                          ?? order["buyer_name"]?.ToString() ?? "";
        var receiverPhone = order["receiver_cellphone"]?.ToString()
                           ?? order["receiver_phone"]?.ToString() ?? "";
        var receiverPhone2 = order["receiver_phone"]?.ToString() ?? "";

        return new Cafe24Order
        {
            OrderId = orderId,
            OrderItemCode = item?["order_item_code"]?.ToString() ?? "",
            RecipientName = receiverName,
            RecipientCellPhone = PhoneNormalizer.Normalize(receiverPhone),
            RecipientPhone = PhoneNormalizer.Normalize(receiverPhone2),
            OrderStatus = item?["order_status"]?.ToString() ?? order["order_status"]?.ToString() ?? "",
            ProductName = item?["product_name"]?.ToString() ?? "",
            OrderAmount = decimal.TryParse(item?["product_price"]?.ToString(), out var amt) ? amt : 0,
            Quantity = int.TryParse(item?["quantity"]?.ToString(), out var qty) ? qty : 0,
            OrderDate = order["order_date"]?.ToString() ?? "",
            ShippingCode = item?["shipping_code"]?.ToString() ?? "",
            RawJson = order.ToString(),
            CachedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        };
    }

    /// <summary>
    /// 주문에 송장번호 입력 + 배송중 처리
    /// </summary>
    public async Task<(bool success, string responseBody, int statusCode)> PushTrackingNumber(
        string orderId, string orderItemCode, string trackingNumber, string shippingCompanyCode)
    {
        // Cafe24 배송 정보 업데이트 API
        // PUT /api/v2/admin/orders/{order_id}/items/{order_item_code}/shipments
        // 먼저 shipment 조회 후 업데이트, 없으면 생성

        try
        {
            // 1) 기존 shipment 조회
            var shipmentsUrl = $"admin/orders/{orderId}/shipments";
            var getResp = await ExecuteWithRetry(() => _http.GetAsync(shipmentsUrl));
            var getBody = getResp != null ? await getResp.Content.ReadAsStringAsync() : "";

            _log.Info($"Shipment 조회: {orderId} → {getResp?.StatusCode}");

            // 2) 송장 정보 업데이트 (POST로 생성 또는 PUT으로 업데이트)
            var payload = new
            {
                shop_no = 1,
                request = new
                {
                    tracking_no = trackingNumber,
                    shipping_company_code = shippingCompanyCode,
                    status = "shipping"
                }
            };

            var jsonPayload = JsonConvert.SerializeObject(payload);
            var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

            // shipping_code가 있으면 PUT, 없으면 POST
            HttpResponseMessage? resp;
            if (!string.IsNullOrEmpty(orderItemCode))
            {
                // Cafe24 fulfillments API
                var fulfillUrl = $"admin/orders/{orderId}/items/{orderItemCode}/shipments";
                resp = await ExecuteWithRetry(() =>
                {
                    var req = new HttpRequestMessage(HttpMethod.Post, fulfillUrl)
                    {
                        Content = new StringContent(jsonPayload, Encoding.UTF8, "application/json")
                    };
                    return _http.SendAsync(req);
                });
            }
            else
            {
                resp = await ExecuteWithRetry(() =>
                    _http.PostAsync(shipmentsUrl, new StringContent(jsonPayload, Encoding.UTF8, "application/json")));
            }

            var respBody = resp != null ? await resp.Content.ReadAsStringAsync() : "no response";
            var statusCode = (int)(resp?.StatusCode ?? 0);
            var success = resp?.IsSuccessStatusCode ?? false;

            if (success)
                _log.Info($"송장 반영 성공: {orderId} → {trackingNumber}");
            else
                _log.Error($"송장 반영 실패: {orderId} → {statusCode}: {respBody}");

            return (success, respBody, statusCode);
        }
        catch (Exception ex)
        {
            _log.Error($"송장 반영 예외: {orderId}", ex);
            return (false, ex.Message, 0);
        }
    }

    private async Task<HttpResponseMessage?> ExecuteWithRetry(Func<Task<HttpResponseMessage>> action)
    {
        for (int i = 0; i < MaxRetries; i++)
        {
            try
            {
                var resp = await action();

                // 401 → 토큰 갱신 후 재시도
                if (resp.StatusCode == HttpStatusCode.Unauthorized && i == 0)
                {
                    _log.Warn("Access Token 만료 감지, 자동 갱신 시도...");
                    if (await RefreshAccessTokenAsync())
                    {
                        resp = await action();
                        return resp;
                    }
                    return resp;
                }

                // Rate limit 처리
                if (resp.StatusCode == HttpStatusCode.TooManyRequests)
                {
                    var retryAfter = resp.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(2);
                    _log.Warn($"Rate limited. {retryAfter.TotalSeconds}초 후 재시도 ({i + 1}/{MaxRetries})");
                    await Task.Delay(retryAfter);
                    continue;
                }

                return resp;
            }
            catch (Exception ex)
            {
                _log.Warn($"API 호출 실패 (시도 {i + 1}/{MaxRetries}): {ex.Message}");
                if (i < MaxRetries - 1)
                    await Task.Delay(TimeSpan.FromSeconds(1 * (i + 1)));
            }
        }
        return null;
    }

    /// <summary>
    /// Refresh Token으로 Access Token 자동 갱신 + appsettings.json 저장
    /// </summary>
    private async Task<bool> RefreshAccessTokenAsync()
    {
        if (string.IsNullOrEmpty(_config.ClientId) || string.IsNullOrEmpty(_config.ClientSecret) ||
            string.IsNullOrEmpty(_config.RefreshToken))
        {
            _log.Error("토큰 갱신 불가: ClientId, ClientSecret, RefreshToken 설정을 확인하세요.");
            return false;
        }

        try
        {
            var tokenUrl = $"https://{_config.MallId}.cafe24api.com/api/v2/oauth/token";
            var authBytes = Encoding.ASCII.GetBytes($"{_config.ClientId}:{_config.ClientSecret}");
            var authHeader = Convert.ToBase64String(authBytes);

            using var tokenHttp = new HttpClient();
            var request = new HttpRequestMessage(HttpMethod.Post, tokenUrl);
            request.Headers.Authorization = new AuthenticationHeaderValue("Basic", authHeader);
            request.Content = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                { "grant_type", "refresh_token" },
                { "refresh_token", _config.RefreshToken }
            });

            var resp = await tokenHttp.SendAsync(request);
            var body = await resp.Content.ReadAsStringAsync();

            if (!resp.IsSuccessStatusCode)
            {
                _log.Error($"토큰 갱신 실패 ({resp.StatusCode}): {body}");
                return false;
            }

            var json = JObject.Parse(body);
            var newAccessToken = json["access_token"]?.ToString();
            var newRefreshToken = json["refresh_token"]?.ToString();

            if (string.IsNullOrEmpty(newAccessToken))
            {
                _log.Error($"토큰 갱신 응답에 access_token 없음: {body}");
                return false;
            }

            // 메모리 갱신
            _config.AccessToken = newAccessToken;
            if (!string.IsNullOrEmpty(newRefreshToken))
                _config.RefreshToken = newRefreshToken;

            // HttpClient 헤더 갱신
            _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", newAccessToken);

            _log.Info("Access Token 자동 갱신 성공");

            // appsettings.json 저장
            SaveTokensToConfig();

            return true;
        }
        catch (Exception ex)
        {
            _log.Error("토큰 갱신 예외", ex);
            return false;
        }
    }

    private void SaveTokensToConfig()
    {
        try
        {
            var configPath = _config.ConfigFilePath;
            if (string.IsNullOrEmpty(configPath) || !File.Exists(configPath)) return;

            var text = File.ReadAllText(configPath);
            var json = JObject.Parse(text);
            var cafe24 = json["Cafe24"] as JObject;
            if (cafe24 == null) return;

            cafe24["AccessToken"] = _config.AccessToken;
            cafe24["RefreshToken"] = _config.RefreshToken;

            File.WriteAllText(configPath, json.ToString(Formatting.Indented));
            _log.Info("appsettings.json 토큰 저장 완료");
        }
        catch (Exception ex)
        {
            _log.Warn($"appsettings.json 토큰 저장 실패: {ex.Message}");
        }
    }
}
