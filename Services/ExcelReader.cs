using ClosedXML.Excel;
using Cafe24ShipmentManager.Models;
using Newtonsoft.Json;

namespace Cafe24ShipmentManager.Services;

public class ExcelReadResult
{
    public List<ShipmentSourceRow> Rows { get; set; } = new();
    public List<string> Vendors { get; set; } = new();
    public int PhoneColumnIndex { get; set; } = -1;
    public string PhoneColumnName { get; set; } = "";
    public List<string> Headers { get; set; } = new();
    public int ShippingCompanyColumnIndex { get; set; } = -1;
}

public class ExcelReader
{
    // 수령인 휴대폰 컬럼 자동 탐색 키워드
    private static readonly string[] PhoneKeywords = {
        "휴대폰", "핸드폰", "수령인휴대폰", "수취인휴대폰",
        "수령인연락처", "수취인연락처", "연락처", "수취인전화",
        "수령인전화", "HP", "Phone", "CellPhone", "전화번호",
        "수령인HP", "수취인HP", "휴대전화"
    };

    private static readonly string[] ShippingCompanyKeywords = {
        "택배사", "배송사", "운송사", "택배", "배송업체", "운송업체"
    };

    public ExcelReadResult ReadFile(string filePath, int? phoneColumnOverride = null)
    {
        var result = new ExcelReadResult();

        using var workbook = new XLWorkbook(filePath);
        var ws = workbook.Worksheets.First();
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
        var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;

        if (lastRow < 2) return result;

        // 헤더 읽기 (1행)
        for (int col = 1; col <= lastCol; col++)
        {
            var header = ws.Cell(1, col).GetString().Trim();
            result.Headers.Add(header);
        }

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

        // 데이터 읽기 (2행부터)
        for (int row = 2; row <= lastRow; row++)
        {
            // C열(3번째) = 발주사명
            var vendor = ws.Cell(row, 3).GetString().Trim();
            if (string.IsNullOrEmpty(vendor)) continue;

            // L열(12번째) = 송장번호
            var tracking = ws.Cell(row, 12).GetString().Trim();

            // 수령인 휴대폰
            var phone = "";
            if (result.PhoneColumnIndex >= 0)
            {
                phone = ws.Cell(row, result.PhoneColumnIndex + 1).GetString().Trim();
            }

            // 택배사
            var shippingCompany = "";
            if (result.ShippingCompanyColumnIndex >= 0)
            {
                shippingCompany = ws.Cell(row, result.ShippingCompanyColumnIndex + 1).GetString().Trim();
            }

            // 수령인명 탐색 (헤더에서 "수령인", "수취인", "받는분" 키워드)
            var recipientName = "";
            for (int i = 0; i < result.Headers.Count; i++)
            {
                var h = result.Headers[i].Replace(" ", "");
                if (h.Contains("수령인명") || h.Contains("수취인명") || h.Contains("받는분") ||
                    (h.Contains("수령인") && h.Contains("이름")) ||
                    (h.Contains("수취인") && h.Contains("이름")))
                {
                    recipientName = ws.Cell(row, i + 1).GetString().Trim();
                    break;
                }
            }

            // 원본 행 JSON
            var rawDict = new Dictionary<string, string>();
            for (int col = 1; col <= lastCol; col++)
            {
                var hdr = col <= result.Headers.Count ? result.Headers[col - 1] : $"Col{col}";
                rawDict[hdr] = ws.Cell(row, col).GetString();
            }

            var normalizedPhone = PhoneNormalizer.Normalize(phone);
            var sourceKey = $"{vendor}|{normalizedPhone}|{tracking}";

            var sourceRow = new ShipmentSourceRow
            {
                SourceRowKey = sourceKey,
                VendorName = vendor,
                TrackingNumber = tracking,
                RecipientPhone = normalizedPhone,
                RecipientName = recipientName,
                ShippingCompany = shippingCompany,
                RawData = JsonConvert.SerializeObject(rawDict),
                ProcessStatus = "pending",
                ImportedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            };

            result.Rows.Add(sourceRow);
            vendorSet.Add(vendor);
        }

        result.Vendors = vendorSet.OrderBy(v => v).ToList();
        return result;
    }
}
