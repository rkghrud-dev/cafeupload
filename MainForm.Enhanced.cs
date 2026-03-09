using System.Globalization;
using Newtonsoft.Json;
using Cafe24ShipmentManager.Data;
using Cafe24ShipmentManager.Models;
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
        var ordered = vendors
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(v => _favoriteShipmentVendorsEx.Contains(v) ? 0 : 1)
            .ThenBy(v => v)
            .ToList();

        return ordered
            .Select(v => FormatVendorLabel(v, _favoriteShipmentVendorsEx.Contains(v)))
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

        SyncStockVendorsFromShipment();

        // 즐겨찾기 공급사가 있으면 시작 시 자동 조회/표시 + 저재고 알림
        if (_favoriteStockVendorsEx.Count > 0)
        {
            await ReloadEnhancedStockDataAsync(showAlert: true, autoRenderFavorites: true);
            return;
        }

        // 초기에는 자동조회하지 않음
        _stockRowsEx.Clear();
        dgvStock.Columns.Clear();
        dgvStock.Rows.Clear();
        _log.Info("재고관리: 조회 버튼을 눌러야 데이터를 표시합니다.");
    }

    private void SyncStockVendorsFromShipment()
    {
        if (_clbStockVendors == null) return;

        var vendors = clbVendors.Items.Cast<object>()
            .Select(x => NormalizeVendorLabel(x?.ToString() ?? ""))
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var ordered = vendors
            .OrderBy(v => _favoriteStockVendorsEx.Contains(v) ? 0 : 1)
            .ThenBy(v => v)
            .ToList();

        var checkedSet = _clbStockVendors.CheckedItems.Cast<object>()
            .Select(x => NormalizeVendorLabel(x?.ToString() ?? ""))
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var hasCheckedState = checkedSet.Count > 0;

        _clbStockVendors.Items.Clear();
        foreach (var v in ordered)
            _clbStockVendors.Items.Add(FormatVendorLabel(v, _favoriteStockVendorsEx.Contains(v)), hasCheckedState ? checkedSet.Contains(v) : true);
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
        var btnStockHistory = new Button { Text = "📋 주문이력", Location = new Point(888, 6), Width = 100, Height = 30 };
        btnStockHistory.Click += (_, _) =>
        {
            using var dlg = new StockOrderHistoryDialog(_db);
            dlg.ShowDialog(this);
        };

        top.Controls.AddRange(new Control[]
        {
            lblSheet, cboStockSheet, btnStockLoad, _btnStockApplyFilter,
            lblRate, _txtStockYuanRate, lblSearch, _txtStockSearch, _btnStockColumns, btnStockHistory
        });

        dgvStock = CreateGridView();
        dgvStock.ReadOnly = true;
        dgvStock.CellDoubleClick += (_, e) => OpenStockOrderDialogFromGrid(e.RowIndex);

        var right = new Panel { Dock = DockStyle.Fill };
        right.Controls.Add(dgvStock);
        right.Controls.Add(top);

        tabStock.Controls.Add(right);
        tabStock.Controls.Add(new Panel { Dock = DockStyle.Left, Width = 1, BackColor = Color.LightGray });
        tabStock.Controls.Add(left);

        _btnStockSelectAll.Click += (_, _) => { for (int i = 0; i < _clbStockVendors.Items.Count; i++) _clbStockVendors.SetItemChecked(i, true); };
        _btnStockDeselectAll.Click += (_, _) => { for (int i = 0; i < _clbStockVendors.Items.Count; i++) _clbStockVendors.SetItemChecked(i, false); };
        _btnStockFavorite.Click += (_, _) => ToggleStockFavoritesEx();
        _btnStockApplyFilter.Click += async (_, _) => await QuerySelectedStockAsync();
        _btnStockColumns.Click += (_, _) => OpenEnhancedStockColumnsDialog();
        btnStockLoad.Click += async (_, _) => await ReloadEnhancedStockDataAsync(showAlert: true, autoRenderFavorites: _favoriteStockVendorsEx.Count > 0);
        _txtStockSearch.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) { ApplyEnhancedStockFilterAndRender(); e.SuppressKeyPress = true; } };
        _txtStockYuanRate.Leave += (_, _) => { SaveEnhancedState(); ApplyEnhancedStockFilterAndRender(); };

        if (!string.IsNullOrWhiteSpace(_loadedYuanRateEx)) _txtStockYuanRate.Text = _loadedYuanRateEx;
    }

    private async Task ReloadEnhancedStockDataAsync(bool showAlert, bool autoRenderFavorites = false)
    {
        if (_sheetsReader == null) return;
        if (cboStockSheet.SelectedItem is not StockSheetItemEx sheet) return;

        btnStockLoad.Enabled = false;
        btnStockLoad.Text = "로딩...";
        try
        {
            var raw = await Task.Run(() => _sheetsReader.ReadRawSheet(EnhancedStockSpreadsheetId, sheet.Title, 30000));
            _stockRowsEx = ParseStockRowsEx(raw);
            _db.ReplaceStockInventoryCache(_stockRowsEx.Select(r => new
            {
                ProductCode = r.ProductCode,
                OrderCode = r.OrderCode,
                Supplier = r.Supplier,
                ImportCostRaw = r.ImportCostRaw,
                SupplyPriceRaw = r.SupplyPriceRaw,
                RetailPriceRaw = r.RetailPriceRaw,
                InboundRaw = r.InboundRaw,
                SoldRaw = r.SoldRaw,
                TwoMonthRaw = r.TwoMonthRaw,
                OneMonthRaw = r.OneMonthRaw,
                ThisMonthRaw = r.ThisMonthRaw,
                StockRaw = r.StockRaw,
                BuyLink = r.BuyLink,
                OptionName = r.OptionName,
                ImportedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            }));

            // 공급사 목록은 출고 송장 발주사와 동일하게 유지
            SyncStockVendorsFromShipment();

            if (autoRenderFavorites && _favoriteStockVendorsEx.Count > 0)
            {
                ApplyFavoriteSelectionToStockVendors();
                ApplyEnhancedStockFilterAndRender();
            }
            else
            {
                // 조회 버튼으로만 화면 출력
                dgvStock.Columns.Clear();
                dgvStock.Rows.Clear();
            }

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


    private void ApplyFavoriteSelectionToStockVendors()
    {
        if (_clbStockVendors == null) return;
        for (int i = 0; i < _clbStockVendors.Items.Count; i++)
        {
            var vendor = NormalizeVendorLabel(_clbStockVendors.Items[i]?.ToString() ?? "");
            _clbStockVendors.SetItemChecked(i, _favoriteStockVendorsEx.Contains(vendor));
        }
    }
    private async Task QuerySelectedStockAsync()
    {
        if (_clbStockVendors == null) return;

        // 데이터 없으면 먼저 로드(새로고침과 동일)
        if (_stockRowsEx.Count == 0)
            await ReloadEnhancedStockDataAsync(showAlert: false);

        var selectedCount = _clbStockVendors.CheckedItems.Count;
        if (selectedCount == 0)
        {
            dgvStock.Columns.Clear();
            dgvStock.Rows.Clear();
            MessageBox.Show("발주사(공급사)를 1개 이상 선택 후 조회하세요.", "알림");
            return;
        }

        ApplyEnhancedStockFilterAndRender();
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
            dgvStock.Rows[idx].Tag = r;

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
    private void OpenStockOrderDialogFromGrid(int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= dgvStock.Rows.Count) return;
        var src = dgvStock.Rows[rowIndex].Tag as StockRowEx;
        if (src == null) return;

        var baseCode = GetBaseVariantCodeA(src.ProductCode);
        var stem = GetVariantStem(src.ProductCode);

        var candidates = _stockRowsEx
            .Where(r => string.Equals(GetVariantStem(r.ProductCode), stem, StringComparison.OrdinalIgnoreCase))
            .OrderBy(r => r.ProductCode)
            .ToList();
        if (candidates.Count == 0) candidates.Add(src);

        var baseRow = candidates.FirstOrDefault(r => string.Equals(r.ProductCode, baseCode, StringComparison.OrdinalIgnoreCase))
            ?? candidates.First();

        var yuanRate = ParseRateEx(_txtStockYuanRate?.Text);

        var lines = candidates.Select(r => new StockOrderLineEx
        {
            ProductCode = r.ProductCode,
            ImportDetail = BuildDefaultImportDetail(baseRow.OrderCode, r.ProductCode, r.OptionName),
            OptionText = r.OptionName,
            CurrentStock = r.StockRaw,
            OrderQty = 0,
            UnitYuan = ParseDefaultUnitYuan(r.ImportCostRaw, yuanRate)
        }).ToList();

        using var dlg = new StockOrderDialogEx(baseCode, baseRow.BuyLink, lines, _db);
        dlg.ShowDialog(this);
    }

    private static string GetBaseVariantCodeA(string code)
    {
        if (string.IsNullOrWhiteSpace(code)) return "";
        return System.Text.RegularExpressions.Regex.Replace(code.Trim(), "[A-Za-z]$", "A");
    }

    private static string GetVariantStem(string code)
    {
        if (string.IsNullOrWhiteSpace(code)) return "";
        return System.Text.RegularExpressions.Regex.Replace(code.Trim(), "[A-Za-z]$", "");
    }

    private static decimal ParseDefaultUnitYuan(string importCostRaw, decimal yuanRate)
    {
        if (yuanRate <= 0) return 0m;
        var s = (importCostRaw ?? "").Replace(",", "").Replace("원", "").Trim();
        if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var krw) && krw > 0)
            return Math.Round(krw / yuanRate, 2);
        if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out krw) && krw > 0)
            return Math.Round(krw / yuanRate, 2);
        return 0m;
    }

    private static string BuildDefaultImportDetail(string baseName, string productCode, string optionText)
    {
        var name = (baseName ?? "").Trim();
        if (!string.IsNullOrWhiteSpace(name))
            return $"{name} {productCode}".Trim();
        return $"{productCode}".Trim();
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
            SelectionMode = DataGridViewSelectionMode.CellSelect,
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

        if (_stockRowsEx.Count > 0 && _favoriteStockVendorsEx.Count > 0)
        {
            ApplyFavoriteSelectionToStockVendors();
            ApplyEnhancedStockFilterAndRender();
            ShowLowStockAlertEx();
        }
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


public sealed class StockOrderLineEx
{
    public string ProductCode { get; set; } = "";
    public string ImportDetail { get; set; } = "";
    public string OptionText { get; set; } = "";
    public string CurrentStock { get; set; } = "";
    public int OrderQty { get; set; }
    public decimal UnitYuan { get; set; }
}

public class StockOrderDialogEx : Form
{
    private readonly List<StockOrderLineEx> _lines;
    private readonly TextBox _txtSiteUrl;
    private readonly DataGridView _grid;
    private readonly Label _lblTotal;
    private readonly TabControl _tabPaste;
    private readonly DatabaseManager? _db;
    private readonly string _baseCodeA;

    public StockOrderDialogEx(string baseCodeA, string siteUrl, List<StockOrderLineEx> lines, DatabaseManager? db = null)
    {
        _db = db;
        _baseCodeA = baseCodeA;
        _lines = lines ?? new List<StockOrderLineEx>();

        Text = $"상품 주문 - {baseCodeA} (입력모드)";
        Size = new Size(1120, 760);
        StartPosition = FormStartPosition.CenterParent;

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 5,
            Padding = new Padding(10)
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 55f));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 45f));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var top = new Panel { Dock = DockStyle.Top, Height = 34 };
        var lblBase = new Label { Text = $"기준코드(A): {baseCodeA}", AutoSize = true, Location = new Point(4, 9) };
        var lblUrl = new Label { Text = "사이트주소:", AutoSize = true, Location = new Point(330, 9) };
        _txtSiteUrl = new TextBox { Text = siteUrl ?? "", Location = new Point(400, 5), Width = 690 };
        top.Controls.AddRange(new Control[] { lblBase, lblUrl, _txtSiteUrl });

        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.CellSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            RowHeadersVisible = false,
            EditMode = DataGridViewEditMode.EditOnEnter,
            ReadOnly = false
        };

        _grid.Columns.Add("Code", "주문코드");
        var colPaste = new DataGridViewButtonColumn
        {
            Name = "GoPaste",
            HeaderText = "입력",
            Text = "입력하기",
            UseColumnTextForButtonValue = true
        };
        _grid.Columns.Add(colPaste);
        _grid.Columns.Add("Detail", "수입상세");
        _grid.Columns.Add("Option", "옵션");
        _grid.Columns.Add("Stock", "현재재고");
        _grid.Columns.Add("Qty", "발주수량");
        _grid.Columns.Add("Unit", "단가(元)");
        _grid.Columns.Add("Amt", "금액(元)");

        _grid.Columns[0].ReadOnly = true;   // 주문코드
        _grid.Columns[1].ReadOnly = true;   // 입력하기 버튼
        _grid.Columns[2].ReadOnly = true;   // 수입상세
        _grid.Columns[3].ReadOnly = true;   // 옵션
        _grid.Columns[4].ReadOnly = true;   // 현재재고
        _grid.Columns[5].ReadOnly = false;  // 발주수량 (유일한 입력칸)
        _grid.Columns[6].ReadOnly = true;   // 단가
        _grid.Columns[7].ReadOnly = true;   // 금액
        _grid.Columns[0].FillWeight = 95;
        _grid.Columns[1].FillWeight = 65;
        _grid.Columns[2].FillWeight = 185;
        _grid.Columns[3].FillWeight = 165;
        _grid.Columns[4].FillWeight = 70;
        _grid.Columns[5].FillWeight = 70;
        _grid.Columns[6].FillWeight = 70;
        _grid.Columns[7].FillWeight = 70;

        // 발주수량 칸 스타일: 입력 가능한 칸으로 시각적 구분
        _grid.Columns[5].DefaultCellStyle.BackColor = Color.LightYellow;
        _grid.Columns[5].DefaultCellStyle.Font = new Font(_grid.Font, FontStyle.Bold);
        _grid.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

        foreach (var line in _lines)
        {
            var idxRow = _grid.Rows.Add(
                line.ProductCode,
                "입력하기",
                line.ImportDetail,
                line.OptionText,
                line.CurrentStock,
                "",
                line.UnitYuan.ToString("0.##"),
                "");
            _grid.Rows[idxRow].Tag = line;
        }

        var info = new Panel { Dock = DockStyle.Top, Height = 34 };
        _lblTotal = new Label { AutoSize = true, Location = new Point(4, 8) };
        info.Controls.Add(_lblTotal);

        _tabPaste = new TabControl { Dock = DockStyle.Fill };

        var btnPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft };
        var btnClose = new Button { Text = "닫기", Width = 90, Height = 32, DialogResult = DialogResult.OK };
        var btnReset = new Button { Text = "초기화", Width = 90, Height = 32 };
        var btnCopy = new Button { Text = "복사", Width = 90, Height = 32 };
        var btnRefresh = new Button { Text = "갱신", Width = 90, Height = 32 };
        var btnSave = new Button { Text = "💾 저장", Width = 90, Height = 32 };
        btnCopy.Click += (_, _) =>
        {
            var active = _tabPaste.SelectedTab?.Controls.OfType<TextBox>().FirstOrDefault();
            Clipboard.SetText(active?.Text ?? "");
            MessageBox.Show("복사되었습니다.", "알림");
        };
        btnRefresh.Click += (_, _) => RebuildOutput();
        btnSave.Click += (_, _) => SaveStockOrder();
        btnReset.Click += (_, _) =>
        {
            foreach (TabPage page in _tabPaste.TabPages)
            {
                var txt = page.Controls.OfType<TextBox>().FirstOrDefault();
                if (txt != null) txt.Text = "";
            }
        };
        if (_db == null) btnSave.Enabled = false;
        // 순서: 저장 | 갱신 | 복사 | 초기화 | 닫기 (RightToLeft이므로 역순 추가)
        btnPanel.Controls.AddRange(new Control[] { btnClose, btnReset, btnCopy, btnRefresh, btnSave });

        root.Controls.Add(top, 0, 0);
        root.Controls.Add(_grid, 0, 1);
        root.Controls.Add(info, 0, 2);
        root.Controls.Add(_tabPaste, 0, 3);
        root.Controls.Add(btnPanel, 0, 4);
        Controls.Add(root);

        _grid.CellEndEdit += (_, _) => RebuildOutput();
        _grid.CellValueChanged += (_, e) =>
        {
            if (e.RowIndex >= 0 && (e.ColumnIndex == 5 || e.ColumnIndex == 6))
                RebuildOutput();
        };
        void AppendRowToPaste(DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (_grid.Columns[e.ColumnIndex].Name != "GoPaste") return;

            var row = _grid.Rows[e.RowIndex];
            var detail = (row.Cells[2].Value?.ToString() ?? "").Trim();
            var option = (row.Cells[3].Value?.ToString() ?? "").Trim();
            var qty = ParseInt(row.Cells[5].Value?.ToString());
            var unit = ParseDecimal(row.Cells[6].Value?.ToString());

            var line = $"{detail} /{option} /{(qty > 0 ? qty.ToString() : "0")} ea/{(unit > 0 ? unit.ToString("0.##") : "0")}元";

            // 첫 번째 탭이 없으면 생성
            if (_tabPaste.TabPages.Count == 0)
            {
                var page = new TabPage("붙여넣기 1");
                var txt = new TextBox
                {
                    Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both,
                    WordWrap = false, AcceptsReturn = true, AcceptsTab = true,
                    ShortcutsEnabled = true, ReadOnly = false, Font = new Font("Consolas", 10f)
                };
                page.Controls.Add(txt);
                _tabPaste.TabPages.Add(page);
            }

            // 현재 활성 탭에 한 줄 추가
            var activeTab = _tabPaste.SelectedTab ?? _tabPaste.TabPages[0];
            var activeTxt = activeTab.Controls.OfType<TextBox>().FirstOrDefault();
            if (activeTxt != null)
            {
                if (!string.IsNullOrEmpty(activeTxt.Text))
                    activeTxt.Text += Environment.NewLine;
                activeTxt.Text += line;
                activeTxt.SelectionStart = activeTxt.Text.Length;
                activeTxt.ScrollToCaret();
            }

            _tabPaste.Focus();
        }
        _grid.CellContentClick += (_, e) => AppendRowToPaste(e);
        _grid.CurrentCellDirtyStateChanged += (_, _) =>
        {
            if (_grid.IsCurrentCellDirty)
                _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        };
        // 초기 붙여넣기 탭 생성
        var initPage = new TabPage("붙여넣기 1");
        var initTxt = new TextBox
        {
            Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both,
            WordWrap = false, AcceptsReturn = true, AcceptsTab = true,
            ShortcutsEnabled = true, ReadOnly = false, Font = new Font("Consolas", 10f)
        };
        initPage.Controls.Add(initTxt);
        _tabPaste.TabPages.Add(initPage);

        RebuildOutput();
    }

    private void SaveStockOrder()
    {
        if (_db == null) return;

        var orderLines = new List<StockOrderLine>();
        foreach (DataGridViewRow row in _grid.Rows)
        {
            if (row.IsNewRow) continue;
            var qty = ParseInt(row.Cells[5].Value?.ToString());
            if (qty <= 0) continue;
            var unit = ParseDecimal(row.Cells[6].Value?.ToString());
            orderLines.Add(new StockOrderLine
            {
                ProductCode = row.Cells[0].Value?.ToString() ?? "",
                ImportDetail = row.Cells[2].Value?.ToString() ?? "",
                OptionText = row.Cells[3].Value?.ToString() ?? "",
                OrderQty = qty,
                UnitYuan = unit,
                AmountYuan = qty * unit
            });
        }

        if (orderLines.Count == 0)
        {
            MessageBox.Show("발주수량이 입력된 항목이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var header = new StockOrderHeader
        {
            BaseCodeA = _baseCodeA,
            SiteUrl = _txtSiteUrl.Text.Trim(),
            TotalQty = orderLines.Sum(l => l.OrderQty),
            TotalAmountYuan = orderLines.Sum(l => l.AmountYuan),
            ItemCount = orderLines.Count,
            OrderedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        };

        try
        {
            _db.InsertStockOrder(header, orderLines);
            MessageBox.Show($"발주 이력이 저장되었습니다.\n품목: {header.ItemCount}건, 수량: {header.TotalQty}, 금액: {header.TotalAmountYuan:0.##}元",
                "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"저장 실패: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void RebuildOutput()
    {
        decimal totalYuan = 0m;
        int totalQty = 0;
        int itemCount = 0;

        foreach (DataGridViewRow row in _grid.Rows)
        {
            if (row.IsNewRow) continue;

            var qty = ParseInt(row.Cells[5].Value?.ToString());
            var unit = ParseDecimal(row.Cells[6].Value?.ToString());
            var amt = qty * unit;
            totalYuan += amt;
            totalQty += qty;
            if (qty > 0) itemCount++;

            row.Cells[7].Value = amt > 0 ? amt.ToString("0.##") : "";
        }
        _lblTotal.Text = $"발주품목: {itemCount}건   합계수량: {totalQty}   총 합계금액(元): {totalYuan:0.##}";
        _lblTotal.ForeColor = SystemColors.ControlText;
    }

    private void RebuildPasteTabs(List<string> chunks)
    {
        var selected = _tabPaste.SelectedIndex;
        _tabPaste.TabPages.Clear();

        if (chunks.Count == 0) chunks.Add("");
        for (int i = 0; i < chunks.Count; i++)
        {
            var page = new TabPage($"붙여넣기 {i + 1}");
            var txt = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                AcceptsReturn = true,
                AcceptsTab = true,
                ShortcutsEnabled = true,
                ReadOnly = false,
                Font = new Font("Consolas", 10f),
                Text = chunks[i]
            };
            page.Controls.Add(txt);
            _tabPaste.TabPages.Add(page);
        }

        if (_tabPaste.TabPages.Count > 0)
            _tabPaste.SelectedIndex = Math.Clamp(selected, 0, _tabPaste.TabPages.Count - 1);
    }

    private static List<string> SplitByMaxChars(string text, int maxChars)
    {
        var result = new List<string>();
        if (string.IsNullOrEmpty(text))
        {
            result.Add("");
            return result;
        }

        var rows = text.Replace("\r\n", "\n").Split('\n');
        var current = new System.Text.StringBuilder();

        foreach (var row in rows)
        {
            if (row.Length > maxChars)
            {
                if (current.Length > 0)
                {
                    result.Add(current.ToString());
                    current.Clear();
                }

                var start = 0;
                while (start < row.Length)
                {
                    var len = Math.Min(maxChars, row.Length - start);
                    result.Add(row.Substring(start, len));
                    start += len;
                }
                continue;
            }

            var appendLen = (current.Length == 0 ? 0 : 1) + row.Length;
            if (current.Length + appendLen > maxChars && current.Length > 0)
            {
                result.Add(current.ToString());
                current.Clear();
            }

            if (current.Length > 0)
                current.Append('\n');
            current.Append(row);
        }

        if (current.Length > 0)
            result.Add(current.ToString());
        if (result.Count == 0)
            result.Add("");

        return result;
    }

    private static int ParseInt(string? raw)
    {
        if (int.TryParse((raw ?? "").Trim(), out var v) && v > 0) return v;
        return 0;
    }

    private static decimal ParseDecimal(string? raw)
    {
        var s = (raw ?? "").Replace(",", "").Replace("원", "").Replace("元", "").Trim();
        if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) return v;
        if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out v)) return v;
        return 0m;
    }
}

// ── 주문 이력 검색 다이얼로그 ──
public class StockOrderHistoryDialog : Form
{
    private readonly DatabaseManager _db;
    private readonly DateTimePicker _dtFrom;
    private readonly DateTimePicker _dtTo;
    private readonly TextBox _txtProductCode;
    private readonly DataGridView _dgvHeaders;
    private readonly DataGridView _dgvLines;

    public StockOrderHistoryDialog(DatabaseManager db)
    {
        _db = db;
        Text = "발주 주문이력";
        Size = new Size(960, 680);
        StartPosition = FormStartPosition.CenterParent;

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(8)
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 45f));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 45f));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        // Search panel
        var search = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false };
        _dtFrom = new DateTimePicker { Format = DateTimePickerFormat.Short, Width = 110, Value = DateTime.Today.AddMonths(-3) };
        _dtTo = new DateTimePicker { Format = DateTimePickerFormat.Short, Width = 110, Value = DateTime.Today };
        _txtProductCode = new TextBox { Width = 140, PlaceholderText = "상품코드 필터" };
        var btnSearch = new Button { Text = "검색", Width = 70, Height = 28 };
        var btnChart = new Button { Text = "📊 차트보기", Width = 100, Height = 28 };
        search.Controls.AddRange(new Control[] {
            new Label { Text = "시작일:", AutoSize = true, Padding = new Padding(0, 5, 0, 0) }, _dtFrom,
            new Label { Text = "~", AutoSize = true, Padding = new Padding(4, 5, 4, 0) }, _dtTo,
            _txtProductCode, btnSearch, btnChart
        });

        // Header grid
        _dgvHeaders = new DataGridView
        {
            Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false
        };
        _dgvHeaders.Columns.Add("OrderedAt", "발주일시");
        _dgvHeaders.Columns.Add("BaseCodeA", "기준코드");
        _dgvHeaders.Columns.Add("ItemCount", "품목수");
        _dgvHeaders.Columns.Add("TotalQty", "총수량");
        _dgvHeaders.Columns.Add("TotalAmountYuan", "총금액(元)");
        _dgvHeaders.Columns["OrderedAt"]!.FillWeight = 140;
        _dgvHeaders.Columns["BaseCodeA"]!.FillWeight = 120;

        // Lines grid
        _dgvLines = new DataGridView
        {
            Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false
        };
        _dgvLines.Columns.Add("ProductCode", "주문코드");
        _dgvLines.Columns.Add("OptionText", "옵션");
        _dgvLines.Columns.Add("OrderQty", "수량");
        _dgvLines.Columns.Add("UnitYuan", "단가(元)");
        _dgvLines.Columns.Add("AmountYuan", "금액(元)");
        _dgvLines.Columns["OptionText"]!.FillWeight = 160;

        // Bottom buttons
        var bottom = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft };
        var btnClose = new Button { Text = "닫기", Width = 90, Height = 32, DialogResult = DialogResult.OK };
        bottom.Controls.Add(btnClose);

        root.Controls.Add(search, 0, 0);
        root.Controls.Add(_dgvHeaders, 0, 1);
        root.Controls.Add(_dgvLines, 0, 2);
        root.Controls.Add(bottom, 0, 3);
        Controls.Add(root);

        btnSearch.Click += (_, _) => DoSearch();
        btnChart.Click += (_, _) =>
        {
            using var chart = new StockOrderChartDialog(_db);
            chart.ShowDialog(this);
        };
        _dgvHeaders.SelectionChanged += (_, _) => ShowSelectedHeaderLines();
        _dgvHeaders.CellDoubleClick += (_, e) =>
        {
            if (e.RowIndex < 0) return;
            var header = _dgvHeaders.Rows[e.RowIndex].Tag as StockOrderHeader;
            if (header == null) return;
            using var detail = new StockOrderDetailDialog(_db, header);
            detail.ShowDialog(this);
        };

        DoSearch();
    }

    private void DoSearch()
    {
        var headers = _db.SearchStockOrderHeaders(
            _dtFrom.Value.ToString("yyyy-MM-dd"),
            _dtTo.Value.ToString("yyyy-MM-dd"),
            _txtProductCode.Text);

        _dgvHeaders.Rows.Clear();
        foreach (var h in headers)
        {
            var idx = _dgvHeaders.Rows.Add(h.OrderedAt, h.BaseCodeA, h.ItemCount, h.TotalQty, h.TotalAmountYuan.ToString("0.##"));
            _dgvHeaders.Rows[idx].Tag = h;
        }
        _dgvLines.Rows.Clear();
    }

    private void ShowSelectedHeaderLines()
    {
        _dgvLines.Rows.Clear();
        if (_dgvHeaders.SelectedRows.Count == 0) return;
        var header = _dgvHeaders.SelectedRows[0].Tag as StockOrderHeader;
        if (header == null) return;

        var lines = _db.GetStockOrderLines(header.Id);
        foreach (var l in lines)
        {
            _dgvLines.Rows.Add(l.ProductCode, l.OptionText, l.OrderQty, l.UnitYuan.ToString("0.##"), l.AmountYuan.ToString("0.##"));
        }
    }
}

// ── 발주 상세 다이얼로그 (월별 옵션 수량 표) ──
public class StockOrderDetailDialog : Form
{
    public StockOrderDetailDialog(DatabaseManager db, StockOrderHeader header)
    {
        Text = $"발주 상세 - {header.BaseCodeA} ({header.OrderedAt})";
        Size = new Size(1000, 560);
        StartPosition = FormStartPosition.CenterParent;

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(8)
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 40f));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 60f));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        // Info label
        var lblInfo = new Label
        {
            Text = $"기준코드: {header.BaseCodeA}    발주일: {header.OrderedAt}    품목: {header.ItemCount}건    총수량: {header.TotalQty}    총금액: {header.TotalAmountYuan:0.##}元",
            AutoSize = true, Font = new Font("맑은 고딕", 10f, FontStyle.Bold), Padding = new Padding(0, 4, 0, 4)
        };

        // Detail grid (this order's lines)
        var dgvDetail = new DataGridView
        {
            Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false
        };
        dgvDetail.Columns.Add("ProductCode", "주문코드");
        dgvDetail.Columns.Add("OptionText", "옵션");
        dgvDetail.Columns.Add("OrderQty", "수량");
        dgvDetail.Columns.Add("UnitYuan", "단가(元)");
        dgvDetail.Columns.Add("AmountYuan", "금액(元)");
        dgvDetail.Columns["OptionText"]!.FillWeight = 160;

        var lines = db.GetStockOrderLines(header.Id);
        foreach (var l in lines)
            dgvDetail.Rows.Add(l.ProductCode, l.OptionText, l.OrderQty, l.UnitYuan.ToString("0.##"), l.AmountYuan.ToString("0.##"));

        // Monthly option breakdown grid (pivot table)
        var dgvMonthly = new DataGridView
        {
            Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.CellSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false
        };

        var records = db.GetOptionMonthlyBreakdown(header.BaseCodeA);
        BuildMonthlyPivot(dgvMonthly, records);

        // Bottom
        var bottom = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft };
        var btnClose = new Button { Text = "닫기", Width = 90, Height = 32, DialogResult = DialogResult.OK };
        bottom.Controls.Add(btnClose);

        root.Controls.Add(lblInfo, 0, 0);
        root.Controls.Add(dgvDetail, 0, 1);
        root.Controls.Add(dgvMonthly, 0, 2);
        root.Controls.Add(bottom, 0, 3);
        Controls.Add(root);
    }

    private static void BuildMonthlyPivot(DataGridView dgv, List<OptionMonthlyRecord> records)
    {
        if (records.Count == 0)
        {
            dgv.Columns.Add("Empty", "데이터 없음");
            return;
        }

        // Collect distinct months and options
        var months = records.Select(r => r.Month).Distinct().OrderBy(m => m).ToList();
        var options = records
            .Select(r => (r.ProductCode, r.OptionText))
            .Distinct()
            .OrderBy(o => o.ProductCode)
            .ToList();

        // Build columns: 주문코드 | 옵션 | month1 | month2 | ... | 합계 | 월평균
        dgv.Columns.Add("ProductCode", "주문코드");
        dgv.Columns.Add("OptionText", "옵션");
        dgv.Columns["OptionText"]!.FillWeight = 140;
        foreach (var m in months)
        {
            var col = dgv.Columns.Add(m, m);
            dgv.Columns[col].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }
        var colTotal = dgv.Columns.Add("Total", "합계");
        dgv.Columns[colTotal].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        dgv.Columns[colTotal].DefaultCellStyle.Font = new Font("맑은 고딕", 9f, FontStyle.Bold);
        var colAvg = dgv.Columns.Add("Avg", "월평균");
        dgv.Columns[colAvg].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        dgv.Columns[colAvg].DefaultCellStyle.Font = new Font("맑은 고딕", 9f, FontStyle.Bold);

        // Build lookup
        var lookup = new Dictionary<(string, string, string), long>();
        foreach (var r in records)
            lookup[(r.ProductCode, r.OptionText, r.Month)] = r.TotalQty;

        // Fill rows
        foreach (var (code, option) in options)
        {
            var row = new object[months.Count + 4];
            row[0] = code;
            row[1] = option;
            long sum = 0;
            int activeMonths = 0;
            for (int i = 0; i < months.Count; i++)
            {
                if (lookup.TryGetValue((code, option, months[i]), out var qty))
                {
                    row[i + 2] = qty;
                    sum += qty;
                    activeMonths++;
                }
                else
                {
                    row[i + 2] = "";
                }
            }
            row[months.Count + 2] = sum;
            row[months.Count + 3] = activeMonths > 0 ? $"{(double)sum / months.Count:0.#}" : "";
            dgv.Rows.Add(row);
        }

        // Grand total row
        var totalRow = new object[months.Count + 4];
        totalRow[0] = "";
        totalRow[1] = "합계";
        long grandSum = 0;
        for (int i = 0; i < months.Count; i++)
        {
            long mSum = records.Where(r => r.Month == months[i]).Sum(r => r.TotalQty);
            totalRow[i + 2] = mSum;
            grandSum += mSum;
        }
        totalRow[months.Count + 2] = grandSum;
        totalRow[months.Count + 3] = months.Count > 0 ? $"{(double)grandSum / months.Count:0.#}" : "";
        var totalIdx = dgv.Rows.Add(totalRow);
        dgv.Rows[totalIdx].DefaultCellStyle.BackColor = Color.LightYellow;
        dgv.Rows[totalIdx].DefaultCellStyle.Font = new Font("맑은 고딕", 9f, FontStyle.Bold);
    }
}

// ── 발주 차트 다이얼로그 (GDI+ 수평 막대) ──
public class StockOrderChartDialog : Form
{
    private readonly DatabaseManager _db;
    private readonly ComboBox _cmbPeriod;
    private readonly ComboBox _cmbYear;
    private readonly Panel _chartPanel;
    private List<TopOrderedProduct> _chartData = new();

    public StockOrderChartDialog(DatabaseManager db)
    {
        _db = db;
        Text = "발주 차트";
        Size = new Size(800, 600);
        StartPosition = FormStartPosition.CenterParent;

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(8)
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        // Top controls
        var top = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false };
        _cmbPeriod = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 90 };
        _cmbPeriod.Items.AddRange(new object[] { "월별", "연별" });
        _cmbPeriod.SelectedIndex = 0;

        _cmbYear = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 80 };
        var currentYear = DateTime.Now.Year;
        for (int y = currentYear; y >= currentYear - 5; y--)
            _cmbYear.Items.Add(y.ToString());
        _cmbYear.SelectedIndex = 0;

        var btnQuery = new Button { Text = "조회", Width = 70, Height = 28 };
        top.Controls.AddRange(new Control[] {
            new Label { Text = "기간:", AutoSize = true, Padding = new Padding(0, 5, 0, 0) }, _cmbPeriod,
            new Label { Text = "연도:", AutoSize = true, Padding = new Padding(8, 5, 0, 0) }, _cmbYear,
            btnQuery
        });

        // Chart panel
        _chartPanel = new DoubleBufferedPanel { Dock = DockStyle.Fill, BackColor = Color.White };
        _chartPanel.Paint += ChartPanel_Paint;

        // Bottom
        var bottom = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft };
        var btnClose = new Button { Text = "닫기", Width = 90, Height = 32, DialogResult = DialogResult.OK };
        bottom.Controls.Add(btnClose);

        root.Controls.Add(top, 0, 0);
        root.Controls.Add(_chartPanel, 0, 1);
        root.Controls.Add(bottom, 0, 2);
        Controls.Add(root);

        btnQuery.Click += (_, _) => RefreshChart();
        RefreshChart();
    }

    private void RefreshChart()
    {
        var year = _cmbYear.SelectedItem?.ToString() ?? DateTime.Now.Year.ToString();
        var isMonthly = _cmbPeriod.SelectedIndex == 0;

        string dateFrom, dateTo;
        if (isMonthly)
        {
            dateFrom = $"{year}-01-01";
            dateTo = $"{year}-12-31";
        }
        else
        {
            dateFrom = $"{int.Parse(year) - 4}-01-01";
            dateTo = $"{year}-12-31";
        }

        _chartData = _db.GetTopOrderedProducts(dateFrom, dateTo, 20);
        _chartPanel.Invalidate();
    }

    private void ChartPanel_Paint(object? sender, PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

        if (_chartData.Count == 0)
        {
            g.DrawString("데이터가 없습니다.", new Font("맑은 고딕", 12f), Brushes.Gray,
                _chartPanel.Width / 2f - 60, _chartPanel.Height / 2f - 10);
            return;
        }

        var padding = new { Left = 160, Top = 30, Right = 60, Bottom = 20 };
        var barHeight = 22;
        var barGap = 6;
        var maxQty = _chartData.Max(d => d.TotalQty);
        if (maxQty == 0) maxQty = 1;

        var chartWidth = _chartPanel.Width - padding.Left - padding.Right;
        var titleFont = new Font("맑은 고딕", 11f, FontStyle.Bold);
        var labelFont = new Font("맑은 고딕", 9f);
        var valueFont = new Font("맑은 고딕", 8.5f, FontStyle.Bold);
        var barBrush = new SolidBrush(Color.FromArgb(66, 133, 244));
        var altBrush = new SolidBrush(Color.FromArgb(52, 168, 83));

        var year = _cmbYear.SelectedItem?.ToString() ?? "";
        var period = _cmbPeriod.SelectedIndex == 0 ? "월별" : "연별";
        g.DrawString($"{year} {period} 상위 발주 상품 (수량 기준)", titleFont, Brushes.Black, padding.Left, 6);

        for (int i = 0; i < _chartData.Count; i++)
        {
            var item = _chartData[i];
            var y = padding.Top + i * (barHeight + barGap);
            var barW = (int)((double)item.TotalQty / maxQty * chartWidth);
            if (barW < 2) barW = 2;

            // Label
            g.DrawString(item.ProductCode, labelFont, Brushes.Black, 4, y + 2);
            // Bar
            var brush = i % 2 == 0 ? barBrush : altBrush;
            g.FillRectangle(brush, padding.Left, y, barW, barHeight);
            // Value
            g.DrawString($"{item.TotalQty} ({item.OrderCount}회)", valueFont, Brushes.Black,
                padding.Left + barW + 4, y + 3);
        }

        titleFont.Dispose();
        labelFont.Dispose();
        valueFont.Dispose();
        barBrush.Dispose();
        altBrush.Dispose();
    }
}

// ── DoubleBufferedPanel for flicker-free drawing ──
internal class DoubleBufferedPanel : Panel
{
    public DoubleBufferedPanel()
    {
        DoubleBuffered = true;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
    }
}
