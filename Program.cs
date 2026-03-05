using Cafe24ShipmentManager;
using Cafe24ShipmentManager.Data;
using Cafe24ShipmentManager.Services;
using Newtonsoft.Json.Linq;

namespace Cafe24ShipmentManager;

static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();

        // 전역 예외 처리
        Application.ThreadException += (_, e) =>
        {
            MessageBox.Show($"예기치 않은 오류:\n{e.Exception.Message}\n\n{e.Exception.StackTrace}",
                "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
        };
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            var ex = e.ExceptionObject as Exception;
            MessageBox.Show($"치명적 오류:\n{ex?.Message}\n\n{ex?.StackTrace}",
                "치명적 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
        };
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

        // 설정 로드
        var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
        JObject config;
        try
        {
            config = JObject.Parse(File.ReadAllText(configPath));
        }
        catch
        {
            config = new JObject();
        }

        // Google Sheets 설정
        var gsSection = config["GoogleSheets"];
        var credentialPath = gsSection?["CredentialPath"]?.ToString() ?? "";
        var spreadsheetId = gsSection?["SpreadsheetId"]?.ToString() ?? "";
        var defaultSheetName = gsSection?["DefaultSheetName"]?.ToString() ?? "";

        // Cafe24 설정
        var cafe24Section = config["Cafe24"];
        var cafe24Config = new Cafe24Config
        {
            MallId = cafe24Section?["MallId"]?.ToString() ?? "YOUR_MALL_ID",
            AccessToken = cafe24Section?["AccessToken"]?.ToString() ?? "YOUR_ACCESS_TOKEN",
            ClientId = cafe24Section?["ClientId"]?.ToString() ?? "",
            ClientSecret = cafe24Section?["ClientSecret"]?.ToString() ?? "",
            RefreshToken = cafe24Section?["RefreshToken"]?.ToString() ?? "",
            ApiVersion = cafe24Section?["ApiVersion"]?.ToString() ?? "2023-03-01",
            DefaultShippingCompanyCode = cafe24Section?["DefaultShippingCompanyCode"]?.ToString() ?? "0019",
            OrderFetchDays = cafe24Section?["OrderFetchDays"]?.Value<int>() ?? 14,
            ConfigFilePath = configPath
        };

        var dbConnStr = config["Database"]?["ConnectionString"]?.ToString() ?? "Data Source=app.db";
        var logDir = config["Logging"]?["LogDirectory"]?.ToString() ?? "logs";

        var logger = new AppLogger(logDir);
        var db = new DatabaseManager(dbConnStr);

        Application.Run(new MainForm(db, logger, cafe24Config, credentialPath, spreadsheetId, defaultSheetName));
    }
}
