using System.Text.RegularExpressions;

namespace Cafe24ShipmentManager.Services;

public static partial class PhoneNormalizer
{
    /// <summary>
    /// 하이픈/공백/괄호 제거 후 숫자만 남기고, 010으로 시작하는 11자리로 정규화
    /// </summary>
    public static string Normalize(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";

        var digits = DigitsOnly().Replace(raw, "");

        // 82-10... → 010...
        if (digits.StartsWith("82") && digits.Length >= 11)
            digits = "0" + digits[2..];

        // +82 처리
        if (digits.StartsWith("8210"))
            digits = "0" + digits[2..];

        // 10xxxxxxxx → 010xxxxxxxx
        if (digits.StartsWith("10") && digits.Length == 10)
            digits = "0" + digits;

        return digits;
    }

    [GeneratedRegex(@"[^\d]")]
    private static partial Regex DigitsOnly();
}
