using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using Cafe24ShipmentManager.Models;
using Newtonsoft.Json;

namespace Cafe24ShipmentManager.Services;

public class GoogleSheetsReader
{
    private readonly SheetsService _service;
    private readonly AppLogger _log;

    private static readonly string[] PhoneKeywords = {
        "휴대폰", "핸드폰", "수령인휴대폰", "수취인휴대폰",
        "수령인연락처", "수취인연락처", "연락처", "수취인전화",
        "수령인전화", "HP", "Phone", "CellPhone", "전화번호",
        "수령인HP", "수취인HP", "휴대전화"
    };

    private static readonly string[] ShippingCompanyKeywords = {
        "택배사", "배송사", "운송사", "택배", "배송업체", "운송업체"
    };

    private GoogleSheetsReader(SheetsService service, AppLogger logger)
    {
        _service = service;
        _log = logger;
    }

    /// <summary>
    /// OAuth2 인증으로 GoogleSheetsReader 생성 (브라우저 로그인)
    /// </summary>
    public static async Task<GoogleSheetsReader> CreateAsync(string credentialJsonPath, AppLogger logger)
    {
        UserCredential credential;
        using (var stream = new FileStream(credentialJsonPath, FileMode.Open, FileAccess.Read))
        {
            // 토큰은 프로그램 폴더의 token 폴더에 저장 (재로그인 불필요)
            var tokenPath = Path.Combine(
                Path.GetDirectoryName(credentialJsonPath) ?? ".", "token");

            credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.FromStream(stream).Secrets,
                new[] { SheetsService.Scope.SpreadsheetsReadonly },
                "user",
                CancellationToken.None,
                new FileDataStore(tokenPath, true));
        }

        var service = new SheetsService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = "Cafe24ShipmentManager"
        });

        logger.Info("Google OAuth2 인증 완료 (브라우저 로그인)");
        return new GoogleSheetsReader(service, logger);
    }

    public List<(string title, int sheetId)> GetSheetList(string spreadsheetId)
    {
        var spreadsheet = _service.Spreadsheets.Get(spreadsheetId).Execute();
        return spreadsheet.Sheets
            .Select(s => (s.Properties.Title, s.Properties.SheetId ?? 0))
            .ToList();
    }

    /// <summary>
    /// C열(발주사명)만 빠르게 가져와서 유니크 목록 반환
    /// </summary>
    public List<string> FetchVendorList(string spreadsheetId, string sheetName)
    {
        var range = $"'{sheetName}'!C:C";
        var request = _service.Spreadsheets.Values.Get(spreadsheetId, range);
        request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMATTEDVALUE;
        var response = request.Execute();
        var values = response.Values;
        if (values == null || values.Count < 2) return new();

        var vendors = new HashSet<string>();
        for (int i = 1; i < values.Count; i++) // 1행(헤더) 스킵
        {
            var v = values[i].Count > 0 ? values[i][0]?.ToString()?.Trim() ?? "" : "";
            if (!string.IsNullOrEmpty(v)) vendors.Add(v);
        }

        _log.Info($"발주사 목록 로드: {vendors.Count}개 ('{sheetName}')");
        return vendors.OrderBy(v => v).ToList();
    }

    /// <summary>
    /// 특정 발주사 + 날짜 범위로 필터링해서 읽기
    /// </summary>
    public ExcelReadResult ReadSheetFiltered(string spreadsheetId, string sheetName,
        HashSet<string> selectedVendors, DateTime? startDate, DateTime? endDate)
    {
        var full = ReadSheet(spreadsheetId, sheetName);

        // 발주사 필터
        if (selectedVendors.Count > 0)
            full.Rows = full.Rows.Where(r => selectedVendors.Contains(r.VendorName)).ToList();

        // D열 발주일 기준 날짜 필터
        if (startDate.HasValue || endDate.HasValue)
        {
            var start = startDate?.Date ?? DateTime.MinValue;
            var end = endDate?.Date ?? DateTime.MaxValue;

            full.Rows = full.Rows.Where(r =>
            {
                if (string.IsNullOrEmpty(r.OrderDate)) return false;
                if (DateTime.TryParse(r.OrderDate, out var dt))
                    return dt.Date >= start && dt.Date <= end;
                return false;
            }).ToList();
        }

        full.Vendors = full.Rows.Select(r => r.VendorName).Distinct().OrderBy(v => v).ToList();
        _log.Info($"필터 적용: {full.Rows.Count}행 (발주사 {selectedVendors.Count}개, 기간 {startDate:yyyy-MM-dd}~{endDate:yyyy-MM-dd})");
        return full;
    }

    public ExcelReadResult ReadSheet(string spreadsheetId, string sheetName, int? phoneColumnOverride = null)
    {
        var result = new ExcelReadResult();

        var range = $"'{sheetName}'";
        var request = _service.Spreadsheets.Values.Get(spreadsheetId, range);
        request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMATTEDVALUE;

        var response = request.Execute();
        var values = response.Values;

        if (values == null || values.Count < 2)
        {
            _log.Warn($"시트 '{sheetName}'에 데이터가 없거나 부족합니다.");
            return result;
        }

        // 헤더 (1행)
        var headerRow = values[0];
        for (int i = 0; i < headerRow.Count; i++)
            result.Headers.Add(headerRow[i]?.ToString()?.Trim() ?? "");

        // 휴대폰 컬럼 찾기
        if (phoneColumnOverride.HasValue)
        {
            result.PhoneColumnIndex = phoneColumnOverride.Value;
            result.PhoneColumnName = result.Headers.Count > phoneColumnOverride.Value
                ? result.Headers[phoneColumnOverride.Value] : $"Column {phoneColumnOverride.Value + 1}";
        }
        else
        {
            for (int i = 0; i < result.Headers.Count; i++)
            {
                var h = result.Headers[i].Replace(" ", "");
                if (PhoneKeywords.Any(k => h.Contains(k, StringComparison.OrdinalIgnoreCase)))
                {
                    result.PhoneColumnIndex = i;
                    result.PhoneColumnName = result.Headers[i];
                    break;
                }
            }

            // 자동 탐색 실패 시 G열(index 6) 폴백
            if (result.PhoneColumnIndex < 0 && result.Headers.Count > 6)
            {
                result.PhoneColumnIndex = 6;
                result.PhoneColumnName = result.Headers[6] + " (G열 기본값)";
                _log.Info($"휴대폰 컬럼 자동탐색 실패 → G열 사용: {result.Headers[6]}");
            }
        }

        // 택배사 컬럼 찾기
        for (int i = 0; i < result.Headers.Count; i++)
        {
            var h = result.Headers[i].Replace(" ", "");
            if (ShippingCompanyKeywords.Any(k => h.Contains(k, StringComparison.OrdinalIgnoreCase)))
            {
                result.ShippingCompanyColumnIndex = i;
                break;
            }
        }

        var vendorSet = new HashSet<string>();

        for (int row = 1; row < values.Count; row++)
        {
            var cells = values[row];
            string GetCell(int col) => col < cells.Count ? cells[col]?.ToString()?.Trim() ?? "" : "";

            var vendor = GetCell(2); // C열
            if (string.IsNullOrEmpty(vendor)) continue;

            var tracking = GetCell(11);        // L열: 송장번호
            var productCode = GetCell(1);       // B열: 상품코드
            var orderDate = GetCell(3);          // D열: 발주일
            var recipientName = GetCell(5);     // F열: 수령인명
            var phone = result.PhoneColumnIndex >= 0 ? GetCell(result.PhoneColumnIndex) : GetCell(6); // G열: 수령인 휴대폰
            var shippingCompany = result.ShippingCompanyColumnIndex >= 0
                ? GetCell(result.ShippingCompanyColumnIndex) : "";

            var rawDict = new Dictionary<string, string>();
            for (int col = 0; col < result.Headers.Count; col++)
                rawDict[result.Headers[col]] = GetCell(col);

            var normalizedPhone = PhoneNormalizer.Normalize(phone);
            var sourceKey = $"{vendor}|{normalizedPhone}|{tracking}";

            result.Rows.Add(new ShipmentSourceRow
            {
                SourceRowKey = sourceKey,
                VendorName = vendor,
                TrackingNumber = tracking,
                RecipientPhone = normalizedPhone,
                RecipientName = recipientName,
                ProductCode = productCode,
                OrderDate = orderDate,
                ShippingCompany = shippingCompany,
                RawData = JsonConvert.SerializeObject(rawDict),
                ProcessStatus = "pending",
                ImportedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            });

            vendorSet.Add(vendor);
        }

        result.Vendors = vendorSet.OrderBy(v => v).ToList();
        _log.Info($"구글시트 '{sheetName}' 읽기 완료: {result.Rows.Count}행, 발주사 {result.Vendors.Count}개");
        return result;
    }
}
