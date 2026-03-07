using System.Globalization;
using Newtonsoft.Json;
using Cafe24ShipmentManager.Services;

namespace Cafe24ShipmentManager;

public partial class MainForm
{
    private const string FavoriteVendorPrefix = "★ ";
    private const string EnhancedStockSpreadsheetId = "1HWR8zdvx0DYbl4ac9hmGuaaIA47nMO0v1CtO99PyC6w";
    private const int EnhancedStockDefaultSheetGid = 2073400281;

    private CheckedListBox? _clbStockVendors;
    private Button? _btnStockSelectAll;
    private Button? _btnStockDeselectAll;
    private Button? _btnStockFavorite;
    private Button? _btnStockApplyFilter;
    private Button? _btnStockColumns;
    private TextBox? _txtStockYuanRate;
    private TextBox? _txtStockSearch;

    private readonly HashSet<string> _favoriteShipmentVendorsEx = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _favoriteStockVendorsEx = new(StringComparer.OrdinalIgnoreCase);
    private HashSet<string> _visibleStockColumnsEx = new(new[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "M", "N", "O" });
    private DateTime _stockAlertMutedDateEx = DateTime.MinValue;
    private List<StockRowEx> _stockRowsEx = new();
    private bool _enhancedStockUiBuilt;

    private void InitEnhancedState()
    {
        LoadEnhancedState();
        EnsureShipmentFavoriteButton();
    }

    private List<string> ApplyShipmentVendorOrder(IEnumerable<string> vendors)
    {
        return vendors
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(v => _favoriteShipmentVendorsEx.Contains(v) ? 0 : 1)
            .ThenBy(v => v)
            .ToList();
    }

    private void EnsureShipmentFavoriteButton()
    {
        if (tabShipment == null) return;
        var leftPanel = tabShipment.Controls.OfType<Panel>().FirstOrDefault(p => p.Dock == DockStyle.Left && p.Width > 100);
        if (leftPanel == null) return;

        var vendorBtnPanel = leftPanel.Controls.OfType<Panel>().FirstOrDefault(p => p.Dock == DockStyle.Bottom);
        if (vendorBtnPanel == null) return;
        if (vendorBtnPanel.Controls.OfType<Button>().Any(b => b.Name == "btnShipmentFavoriteEx")) return;

        var btn = new Button
        {
            Name = "btnShipmentFavoriteEx",
            Text = "★ 즐겨",
            Width = 64,
            Height = 26,
            Location = new Point(148, 0)
        };
        btn.Click += (_, _) => ToggleShipmentFavoritesEx();
        vendorBtnPanel.Controls.Add(btn);
    }

    private void ToggleShipmentFavoritesEx()
    {
        var selected = clbVendors.SelectedItems.Cast<object>()
            .Select(x => NormalizeVendorLabel(x.ToString() ?? ""))
            .Where(x => x.Length > 0)
            .ToList();
        if (selected.Count == 0)
        {
            MessageBox.Show("즐겨찾기할 발주사를 선택하세요.", "알림");
            return;
        }

        foreach (var v in selected)
        {
            if (!_favoriteShipmentVendorsEx.Add(v)) _favoriteShipmentVendorsEx.Remove(v);
        }

        var checkedSet = clbVendors.CheckedItems.Cast<object>()
            .Select(x => NormalizeVendorLabel(x.ToString() ?? ""))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var all = clbVendors.Items.Cast<object>()
            .Select(x => NormalizeVendorLabel(x.ToString() ?? ""))
            .Where(x => x.Length > 0)
            .ToList();

        clbVendors.Items.Clear();
        foreach (var v in ApplyShipmentVendorOrder(all))
            clbVendors.Items.Add(v, checkedSet.Contains(NormalizeVendorLabel(v)));

        SaveEnhancedState();
    }

    private async Task InitEnhancedStockAsync()
    {
        if (_sheetsReader == null) return;
        BuildEnhancedStockUi();

        var sheets = await Task.Run(() => _sheetsReader.GetSheetList(EnhancedStockSpreadsheetId));
        cboStockSheet.Items.Clear();
        foreach (var (title, sheetId) in sheets)
            cboStockSheet.Items.Add(new StockSheetItemEx(title, sheetId));

        var idx = sheets.FindIndex(s => s.sheetId == EnhancedStockDefaultSheetGid);
        cboStockSheet.SelectedIndex = idx >= 0 ? idx : (cboStockSheet.Items.Count > 0 ? 0 : -1);

        if (cboStockSheet.SelectedIndex >= 0)
            await ReloadEnhancedStockDataAsync(showAlert: true);
    }

    private void BuildEnhancedStockUi()
    {
        if (_enhancedStockUiBuilt) return;
        _enhancedStockUiBuilt = true;

        tabStock.Controls.Clear();

        var left = new Panel { Dock = DockStyle.Left, Width = 220, Padding = new Padding(4) };
        var lblVendor = new Label { Text = "공급사(발주사)", Dock = DockStyle.Top, Height = 20 };
        _clbStockVendors = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, Font = new Font("맑은 고딕", 9f) };

        var vendorBtns = new Panel { Dock = DockStyle.Bottom, Height = 28 };
        _btnStockSelectAll = new Button { Text = "전체선택", Width = 70, Height = 26, Location = new Point(0, 0) };
        _btnStockDeselectAll = new Button { Text = "전체해제", Width = 70, Height = 26, Location = new Point(74, 0) };
        _btnStockFavorite = new Button { Text = "★ 즐겨", Width = 64, Height = 26, Location = new Point(148, 0) };
        vendorBtns.Controls.AddRange(new Control[] { _btnStockSelectAll, _btnStockDeselectAll, _btnStockFavorite });

        left.Controls.Add(_clbStockVendors);
        left.Controls.Add(vendorBtns);
        left.Controls.Add(lblVendor);

        var top = new Panel { Dock = DockStyle.Top, Height = 42 };
        var lblSheet = new Label { Text = "재고 시트:", Location = new Point(8, 12), AutoSize = true };
        cboStockSheet = new ComboBox { Location = new Point(70, 8), Width = 260, DropDownStyle = ComboBoxStyle.DropDownList };
        btnStockLoad = new Button { Text = "새로고침", Location = new Point(336, 6), Width = 78, Height = 30 };
        _btnStockApplyFilter = new Button { Text = "조회", Location = new Point(418, 6), Width = 58, Height = 30 };

        var lblRate = new Label { Text = "위안환율:", Location = new Point(486, 12), AutoSize = true };
        _txtStockYuanRate = new TextBox { Location = new Point(548, 8), Width = 56, Text = "320" };

        var lblSearch = new Label { Text = "코드검색:", Location = new Point(612, 12), AutoSize = true };
        _txtStockSearch = new TextBox { Location = new Point(672, 8), Width = 120 };

        _btnStockColumns = new Button { Text = "컬럼선택", Location = new Point(798, 6), Width = 84, Height = 30 };

        top.Controls.AddRange(new Control[]
        {
            lblSheet, cboStockSheet, btnStockLoad, _btnStockApplyFilter,
            lblRate, _txtStockYuanRate, lblSearch, _txtStockSearch, _btnStockColumns
        });

        dgvStock = CreateGridView();
        dgvStock.ReadOnly = true;

        var right = new Panel { Dock = DockStyle.Fill };
        right.Controls.Add(dgvStock);
        right.Controls.Add(top);

        tabStock.Controls.Add(right);
        tabStock.Controls.Add(new Panel { Dock = DockStyle.Left, Width = 1, BackColor = Color.LightGray });
        tabStock.Controls.Add(left);

        _btnStockSelectAll.Click += (_, _) => { for (int i = 0; i < _clbStockVendors.Items.Count; i++) _clbStockVendors.SetItemChecked(i, true); };
        _btnStockDeselectAll.Click += (_, _) => { for (int i = 0; i < _clbStockVendors.Items.Count; i++) _clbStockVendors.SetItemChecked(i, false); };
        _btnStockFavorite.Click += (_, _) => ToggleStockFavoritesEx();
        _btnStockApplyFilter.Click += (_, _) => ApplyEnhancedStockFilterAndRender();
        _btnStockColumns.Click += (_, _) => OpenEnhancedStockColumnsDialog();
        btnStockLoad.Click += async (_, _) => await ReloadEnhancedStockDataAsync(showAlert: false);
        _txtStockSearch.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) { ApplyEnhancedStockFilterAndRender(); e.SuppressKeyPress = true; } };
        _txtStockYuanRate.Leave += (_, _) => { SaveEnhancedState(); ApplyEnhancedStockFilterAndRender(); };

        if (!string.IsNullOrWhiteSpace(_loadedYuanRateEx)) _txtStockYuanRate.Text = _loadedYuanRateEx;
    }

    private async Task ReloadEnhancedStockDataAsync(bool showAlert)
    {
        if (_sheetsReader == null) return;
        if (cboStockSheet.SelectedItem is not StockSheetItemEx sheet) return;

        btnStockLoad.Enabled = false;
        btnStockLoad.Text = "로딩...";
        try
        {
            var raw = await Task.Run(() => _sheetsReader.ReadRawSheet(EnhancedStockSpreadsheetId, sheet.Title, 30000));
            _stockRowsEx = ParseStockRowsEx(raw);

            var vendors = _stockRowsEx.Select(r => r.Supplier).Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().ToList();
            _clbStockVendors!.Items.Clear();
            foreach (var v in vendors.OrderBy(v => _favoriteStockVendorsEx.Contains(v) ? 0 : 1).ThenBy(v => v))
                _clbStockVendors.Items.Add(FormatVendorLabel(v, _favoriteStockVendorsEx.Contains(v)));

            ApplyEnhancedStockFilterAndRender();
            if (showAlert) ShowLowStockAlertEx();
        }
        catch (Exception ex)
        {
            _log.Error("재고 데이터 로드 실패", ex);
            MessageBox.Show($"재고 데이터 로드 오류:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnStockLoad.Enabled = true;
            btnStockLoad.Text = "새로고침";
        }
    }

    private void ApplyEnhancedStockFilterAndRender()
    {
        if (_clbStockVendors == null || _txtStockSearch == null) return;

        var selected = _clbStockVendors.CheckedItems.Cast<object>()
            .Select(x => NormalizeVendorLabel(x.ToString() ?? ""))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var keyword = _txtStockSearch.Text.Trim();

        var rows = _stockRowsEx.AsEnumerable();
        if (selected.Count > 0) rows = rows.Where(r => selected.Contains(r.Supplier));
        if (!string.IsNullOrWhiteSpace(keyword))
            rows = rows.Where(r => r.ProductCode.Contains(keyword, StringComparison.OrdinalIgnoreCase)
                                || r.OrderCode.Contains(keyword, StringComparison.OrdinalIgnoreCase)
                                || r.OptionName.Contains(keyword, StringComparison.OrdinalIgnoreCase));

        RenderEnhancedStock(rows.ToList());
    }

    private void RenderEnhancedStock(List<StockRowEx> rows)
    {
        dgvStock.Columns.Clear();
        dgvStock.Rows.Clear();

        var columns = new[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "M", "N", "O" }
            .Where(c => _visibleStockColumnsEx.Contains(c)).ToList();
        if (columns.Count == 0) columns = new List<string> { "A", "B", "C", "M" };

        foreach (var c in columns)
        {
            dgvStock.Columns.Add(c, c switch
            {
                "A" => "상품코드", "B" => "주문코드", "C" => "공급사", "D" => "수입가", "E" => "공급가",
                "F" => "소매가", "G" => "입고", "H" => "판매", "I" => "2달전", "J" => "1달전", "K" => "이달",
                "M" => "재고", "N" => "구매링크", "O" => "옵션명", _ => c
            });
        }

        var rate = ParseRateEx(_txtStockYuanRate?.Text);
        foreach (var r in rows)
        {
            var vals = new List<object>();
            foreach (var c in columns)
            {
                vals.Add(c switch
                {
                    "A" => r.ProductCode,
                    "B" => r.OrderCode,
                    "C" => r.Supplier,
                    "D" => FormatMoneyWithYuanEx(r.ImportCost, rate),
                    "E" => FormatMoneyWithYuanEx(r.SupplyPrice, rate),
                    "F" => FormatMoneyWithYuanEx(r.RetailPrice, rate),
                    "G" => r.InboundRaw,
                    "H" => r.SoldRaw,
                    "I" => r.TwoMonthRaw,
                    "J" => r.OneMonthRaw,
                    "K" => r.ThisMonthRaw,
                    "M" => r.StockRaw,
                    "N" => r.BuyLink,
                    "O" => r.OptionName,
                    _ => ""
                });
            }
            var idx = dgvStock.Rows.Add(vals.ToArray());

            if (columns.Contains("M"))
            {
                var cell = dgvStock.Rows[idx].Cells["M"];
                cell.Style.BackColor = Color.FromArgb(255, 249, 196);
                if (r.Inbound > 0 && r.Stock >= 0 && r.Stock < r.Inbound * 0.3m)
                {
                    cell.Style.BackColor = Color.FromArgb(255, 205, 210);
                    cell.Style.ForeColor = Color.DarkRed;
                    cell.Style.Font = new Font(dgvStock.Font, FontStyle.Bold);
                    cell.Value = $"{r.StockRaw} (30%↓)";
                }
            }
        }

        _log.Info($"재고 필터 결과: {rows.Count}행");
    }

    private void ShowLowStockAlertEx()
    {
        if (_stockAlertMutedDateEx.Date == DateTime.Today) return;

        // 알림 대상은 재고관리 즐겨찾기 발주사(공급사)만
        if (_favoriteStockVendorsEx.Count == 0) return;

        var lows = _stockRowsEx
            .Where(r => _favoriteStockVendorsEx.Contains(r.Supplier))
            .Where(r => r.Inbound > 0 && r.Stock >= 0 && r.Stock < r.Inbound * 0.3m)
            .Take(200)
            .ToList();
        if (lows.Count == 0) return;

        using var dlg = new Form { Text = "저재고 알림 (즐겨찾기 기준)", Size = new Size(760, 500), StartPosition = FormStartPosition.CenterParent };
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        };
        grid.Columns.Add("A", "상품코드");
        grid.Columns.Add("C", "공급사");
        grid.Columns.Add("G", "입고");
        grid.Columns.Add("M", "재고");
        grid.Columns.Add("O", "옵션명");
        foreach (var r in lows) grid.Rows.Add(r.ProductCode, r.Supplier, r.InboundRaw, r.StockRaw, r.OptionName);

        var bottom = new Panel { Dock = DockStyle.Bottom, Height = 42 };
        var chkMute = new CheckBox { Text = "오늘은 알람 안 울림", AutoSize = true, Location = new Point(8, 12) };
        var btnGo = new Button { Text = "해당코드로 이동", Width = 120, Height = 28, Location = new Point(480, 7) };
        var btnClose = new Button { Text = "닫기", Width = 80, Height = 28, Location = new Point(610, 7), DialogResult = DialogResult.OK };

        btnGo.Click += (_, _) =>
        {
            if (grid.CurrentRow == null) return;
            var code = grid.CurrentRow.Cells[0]?.Value?.ToString() ?? "";
            if (!string.IsNullOrWhiteSpace(code))
            {
                _txtStockSearch!.Text = code;
                ApplyEnhancedStockFilterAndRender();
                tabMain.SelectedTab = tabStock;
            }
        };

        dlg.FormClosing += (_, _) =>
        {
            if (chkMute.Checked)
            {
                _stockAlertMutedDateEx = DateTime.Today;
                SaveEnhancedState();
            }
        };

        bottom.Controls.AddRange(new Control[] { chkMute, btnGo, btnClose });
        dlg.Controls.Add(grid);
        dlg.Controls.Add(bottom);
        dlg.ShowDialog(this);
    }

    private void OpenEnhancedStockColumnsDialog()
    {
        var all = new[]
        {
            ("A", "상품코드"), ("B", "주문코드"), ("C", "공급사"), ("D", "수입가"), ("E", "공급가"),
            ("F", "소매가"), ("G", "입고"), ("H", "판매"), ("I", "2달전"), ("J", "1달전"), ("K", "이달"),
            ("M", "재고"), ("N", "구매링크"), ("O", "옵션명")
        };

        using var dlg = new StockColumnSelectDialogEx(all, _visibleStockColumnsEx);
        if (dlg.ShowDialog(this) != DialogResult.OK) return;
        _visibleStockColumnsEx = dlg.SelectedColumns;
        SaveEnhancedState();
        ApplyEnhancedStockFilterAndRender();
    }

    private void ToggleStockFavoritesEx()
    {
        if (_clbStockVendors == null) return;
        var selected = _clbStockVendors.SelectedItems.Cast<object>()
            .Select(x => NormalizeVendorLabel(x.ToString() ?? ""))
            .Where(x => x.Length > 0)
            .ToList();
        if (selected.Count == 0)
        {
            MessageBox.Show("즐겨찾기할 공급사를 선택하세요.", "알림");
            return;
        }

        foreach (var v in selected)
            if (!_favoriteStockVendorsEx.Add(v)) _favoriteStockVendorsEx.Remove(v);

        var checkedSet = _clbStockVendors.CheckedItems.Cast<object>()
            .Select(x => NormalizeVendorLabel(x.ToString() ?? ""))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var all = _clbStockVendors.Items.Cast<object>()
            .Select(x => NormalizeVendorLabel(x.ToString() ?? ""))
            .Where(x => x.Length > 0)
            .ToList();

        _clbStockVendors.Items.Clear();
        foreach (var v in all.OrderBy(v => _favoriteStockVendorsEx.Contains(v) ? 0 : 1).ThenBy(v => v))
            _clbStockVendors.Items.Add(FormatVendorLabel(v, _favoriteStockVendorsEx.Contains(v)), checkedSet.Contains(v));

        SaveEnhancedState();
    }


    private static string NormalizeVendorLabel(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return "";
        return text.StartsWith(FavoriteVendorPrefix, StringComparison.Ordinal)
            ? text[FavoriteVendorPrefix.Length..].Trim()
            : text.Trim();
    }

    private static string FormatVendorLabel(string vendor, bool isFavorite)
    {
        var clean = NormalizeVendorLabel(vendor);
        return isFavorite ? FavoriteVendorPrefix + clean : clean;
    }
    private List<StockRowEx> ParseStockRowsEx(RawSheetData raw)
    {
        string Get(List<string> arr, int idx) => idx < arr.Count ? arr[idx] : "";
        var rows = new List<StockRowEx>();
        foreach (var cells in raw.Rows)
        {
            var r = new StockRowEx
            {
                ProductCode = Get(cells, 0),
                OrderCode = Get(cells, 1),
                Supplier = Get(cells, 2),
                ImportCostRaw = Get(cells, 3),
                SupplyPriceRaw = Get(cells, 4),
                RetailPriceRaw = Get(cells, 5),
                InboundRaw = Get(cells, 6),
                SoldRaw = Get(cells, 7),
                TwoMonthRaw = Get(cells, 8),
                OneMonthRaw = Get(cells, 9),
                ThisMonthRaw = Get(cells, 10),
                StockRaw = Get(cells, 12),
                BuyLink = Get(cells, 13),
                OptionName = Get(cells, 14)
            };

            r.ImportCost = ParseDecimalEx(r.ImportCostRaw);
            r.SupplyPrice = ParseDecimalEx(r.SupplyPriceRaw);
            r.RetailPrice = ParseDecimalEx(r.RetailPriceRaw);
            r.Inbound = ParseDecimalEx(r.InboundRaw);
            r.Stock = ParseDecimalEx(r.StockRaw);

            if (string.IsNullOrWhiteSpace(r.ProductCode) && string.IsNullOrWhiteSpace(r.OrderCode) && string.IsNullOrWhiteSpace(r.Supplier))
                continue;

            rows.Add(r);
        }
        return rows;
    }

    private static decimal ParseDecimalEx(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return -1;
        var s = raw.Replace(",", "").Replace("원", "").Trim();
        if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) return v;
        if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out v)) return v;
        return -1;
    }

    private static decimal ParseRateEx(string? raw)
    {
        var r = ParseDecimalEx(raw ?? "");
        return r > 0 ? r : 320m;
    }

    private static string FormatMoneyWithYuanEx(decimal krw, decimal rate)
    {
        if (krw < 0) return "";
        var yuan = rate > 0 ? krw / rate : 0;
        return $"{krw:0} / {yuan:0.##}위안";
    }

    private string _loadedYuanRateEx = "";

    private string GetEnhancedStatePath()
    {
        var dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "ui_state.json");
    }

    private void LoadEnhancedState()
    {
        try
        {
            var path = GetEnhancedStatePath();
            if (!File.Exists(path)) return;

            var st = JsonConvert.DeserializeObject<EnhancedUiState>(File.ReadAllText(path));
            if (st == null) return;

            _favoriteShipmentVendorsEx.Clear();
            foreach (var x in st.FavoriteShipmentVendors ?? new List<string>()) _favoriteShipmentVendorsEx.Add(x);
            _favoriteStockVendorsEx.Clear();
            foreach (var x in st.FavoriteStockVendors ?? new List<string>()) _favoriteStockVendorsEx.Add(x);

            _visibleStockColumnsEx = new HashSet<string>(st.VisibleStockColumns ?? new List<string> { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "M", "N", "O" });
            _loadedYuanRateEx = st.YuanRate ?? "320";

            if (DateTime.TryParse(st.StockAlertMutedDate, out var dt)) _stockAlertMutedDateEx = dt.Date;
        }
        catch { }
    }

    private void SaveEnhancedState()
    {
        try
        {
            var st = new EnhancedUiState
            {
                FavoriteShipmentVendors = _favoriteShipmentVendorsEx.OrderBy(x => x).ToList(),
                FavoriteStockVendors = _favoriteStockVendorsEx.OrderBy(x => x).ToList(),
                VisibleStockColumns = _visibleStockColumnsEx.ToList(),
                StockAlertMutedDate = _stockAlertMutedDateEx == DateTime.MinValue ? "" : _stockAlertMutedDateEx.ToString("yyyy-MM-dd"),
                YuanRate = _txtStockYuanRate?.Text ?? _loadedYuanRateEx
            };
            File.WriteAllText(GetEnhancedStatePath(), JsonConvert.SerializeObject(st, Formatting.Indented));
        }
        catch { }
    }

    private sealed class StockSheetItemEx
    {
        public string Title { get; }
        public int SheetId { get; }
        public StockSheetItemEx(string title, int sheetId) { Title = title; SheetId = sheetId; }
        public override string ToString() => $"{Title} (gid:{SheetId})";
    }

    private sealed class StockRowEx
    {
        public string ProductCode { get; set; } = "";
        public string OrderCode { get; set; } = "";
        public string Supplier { get; set; } = "";
        public string ImportCostRaw { get; set; } = "";
        public string SupplyPriceRaw { get; set; } = "";
        public string RetailPriceRaw { get; set; } = "";
        public string InboundRaw { get; set; } = "";
        public string SoldRaw { get; set; } = "";
        public string TwoMonthRaw { get; set; } = "";
        public string OneMonthRaw { get; set; } = "";
        public string ThisMonthRaw { get; set; } = "";
        public string StockRaw { get; set; } = "";
        public string BuyLink { get; set; } = "";
        public string OptionName { get; set; } = "";

        public decimal ImportCost { get; set; }
        public decimal SupplyPrice { get; set; }
        public decimal RetailPrice { get; set; }
        public decimal Inbound { get; set; }
        public decimal Stock { get; set; }
    }

    private sealed class EnhancedUiState
    {
        public List<string>? FavoriteShipmentVendors { get; set; }
        public List<string>? FavoriteStockVendors { get; set; }
        public List<string>? VisibleStockColumns { get; set; }
        public string? StockAlertMutedDate { get; set; }
        public string? YuanRate { get; set; }
    }
}

public class StockColumnSelectDialogEx : Form
{
    public HashSet<string> SelectedColumns { get; private set; } = new();
    private readonly CheckedListBox _clb;
    private readonly (string Key, string Label)[] _all;

    public StockColumnSelectDialogEx((string Key, string Label)[] all, HashSet<string> selected)
    {
        _all = all;
        Text = "재고 컬럼 선택";
        Size = new Size(340, 460);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        _clb = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true };
        foreach (var c in _all)
        {
            var idx = _clb.Items.Add($"[{c.Key}] {c.Label}");
            if (selected.Contains(c.Key)) _clb.SetItemChecked(idx, true);
        }

        var panel = new Panel { Dock = DockStyle.Bottom, Height = 42 };
        var btnOk = new Button { Text = "적용", Width = 72, Height = 28, Location = new Point(160, 7), DialogResult = DialogResult.OK };
        var btnCancel = new Button { Text = "취소", Width = 72, Height = 28, Location = new Point(240, 7), DialogResult = DialogResult.Cancel };
        btnOk.Click += (_, _) =>
        {
            SelectedColumns = new HashSet<string>();
            foreach (int i in _clb.CheckedIndices)
                SelectedColumns.Add(_all[i].Key);
        };

        panel.Controls.AddRange(new Control[] { btnOk, btnCancel });
        Controls.Add(_clb);
        Controls.Add(panel);
        AcceptButton = btnOk;
        CancelButton = btnCancel;
    }
}








