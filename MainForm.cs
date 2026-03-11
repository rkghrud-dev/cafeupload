using System.Text;
using Cafe24ShipmentManager.Data;
using Cafe24ShipmentManager.Models;
using Cafe24ShipmentManager.Services;
using Newtonsoft.Json;

namespace Cafe24ShipmentManager;

public partial class MainForm : Form
{
    // ── Services ──
    private readonly DatabaseManager _db;
    private readonly AppLogger _log;
    private readonly Cafe24ApiClient _api;
    private readonly MatchingEngine _matcher;
    private readonly Cafe24Config _cafe24Config;
    private GoogleSheetsReader? _sheetsReader;
    private readonly string _credentialPath;
    private readonly string _spreadsheetId;
    private readonly string _defaultSheetName;
    private const string StockSpreadsheetId = "1HWR8zdvx0DYbl4ac9hmGuaaIA47nMO0v1CtO99PyC6w";
    private const int StockDefaultSheetGid = 2073400281;

    // ── State ──
    private ExcelReadResult? _excelResult;
    private List<ShipmentSourceRow> _filteredRows = new();
    private List<Cafe24Order> _cafe24Orders = new();
    private List<MatchResult> _matchResults = new();

    // ── Top Bar (공통) ──
    private ComboBox cboSheet = null!;
    private Label lblStatus = null!;

    // ── 출고 탭 컨트롤 ──
    private CheckedListBox clbVendors = null!;
    private Button btnSelectAll = null!;
    private Button btnDeselectAll = null!;
    private DateTimePicker dtpStart = null!;
    private DateTimePicker dtpEnd = null!;
    private Button btnFetch = null!;
    private ComboBox cboShippingCompany = null!;

    // ── Main Tabs ──
    private TabControl tabMain = null!;
    private TabPage tabShipment = null!;
    private TabPage tabPopular = null!;
    private TabPage tabProduct = null!;
    private TabPage tabStock = null!;
    private ComboBox cboStockSheet = null!;
    private Button btnStockLoad = null!;
    private DataGridView dgvStock = null!;

    // ── 출고 탭 내부 서브탭 ──
    private TabControl tabShipSub = null!;
    private DataGridView dgvData = null!;
    private DataGridView dgvMatch = null!;
    private DataGridView dgvResult = null!;
    private Button btnMatch = null!;
    private Button btnPush = null!;
    private Button btnExportFailed = null!;

    // ── Log ──
    private TextBox txtLog = null!;

    public MainForm(DatabaseManager db, AppLogger log, Cafe24Config cafe24Config,
                    string credentialPath, string spreadsheetId, string defaultSheetName)
    {
        _db = db;
        _log = log;
        _cafe24Config = cafe24Config;
        _api = new Cafe24ApiClient(cafe24Config, log);
        _matcher = new MatchingEngine(db, log);
        _credentialPath = credentialPath;
        _spreadsheetId = spreadsheetId;
        _defaultSheetName = defaultSheetName;

        InitializeUI();
        WireEvents();
        InitEnhancedState();

        _log.OnLog += msg =>
        {
            if (txtLog.InvokeRequired)
                txtLog.Invoke(() => AppendLog(msg));
            else
                AppendLog(msg);
        };

        _log.Info("프로그램 시작");
        _ = InitAsync();
    }

    private void AppendLog(string msg)
    {
        txtLog.AppendText(msg + Environment.NewLine);
        txtLog.SelectionStart = txtLog.TextLength;
        txtLog.ScrollToCaret();
    }

    // ═══════════════════════════════════════
    // 초기화: 인증 → 시트목록 → 발주사 자동 로드
    // ═══════════════════════════════════════
    private async Task InitAsync()
    {
        try
        {
            lblStatus.Text = "Google 인증 중...";
            lblStatus.ForeColor = Color.Blue;

            _sheetsReader = await GoogleSheetsReader.CreateAsync(_credentialPath, _log);

            // 시트 목록
            var sheets = _sheetsReader.GetSheetList(_spreadsheetId);
            cboSheet.Items.Clear();
            foreach (var (title, _) in sheets)
                cboSheet.Items.Add(title);

            // 기본 시트 "출고정보" 선택
            var defaultIdx = sheets.FindIndex(s => s.title == _defaultSheetName);
            if (defaultIdx < 0)
                defaultIdx = sheets.FindIndex(s => s.title.Contains("출고"));
            cboSheet.SelectedIndex = defaultIdx >= 0 ? defaultIdx : 0;

            lblStatus.Text = $"✅ 연결 완료 ({sheets.Count}개 시트)";
            lblStatus.ForeColor = Color.DarkGreen;

            // 발주사 자동 로드
            await LoadVendorsAsync();

            // 재고관리 확장 초기화
            await InitEnhancedStockAsync();
        }
        catch (Exception ex)
        {
            _log.Error("초기화 실패", ex);
            lblStatus.Text = "❌ 연결 실패";
            lblStatus.ForeColor = Color.Red;
        }
    }

    private async Task LoadVendorsAsync()
    {
        if (_sheetsReader == null) return;
        var sheetName = cboSheet.SelectedItem?.ToString();
        if (string.IsNullOrEmpty(sheetName)) return;

        try
        {
            lblStatus.Text = $"발주사 목록 로딩 중...";
            lblStatus.ForeColor = Color.Blue;

            var vendors = await Task.Run(() =>
                _sheetsReader.FetchVendorList(_spreadsheetId, sheetName));

            clbVendors.Items.Clear();
            foreach (var v in ApplyShipmentVendorOrder(vendors))
                clbVendors.Items.Add(v);

            SyncStockVendorsFromShipment();

            lblStatus.Text = $"✅ '{sheetName}' — 발주사 {vendors.Count}개 로드됨";
            lblStatus.ForeColor = Color.DarkGreen;
        }
        catch (Exception ex)
        {
            _log.Error("발주사 로드 실패", ex);
        }
    }

    // ═══════════════════════════════════════
    // UI 구성
    // ═══════════════════════════════════════
    private void InitializeUI()
    {
        Text = "Cafe24 출고/송장 관리 매니저";
        Size = new Size(1400, 950);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("맑은 고딕", 9f);

        // ═══ 최상단: 시트 선택 ═══
        var topBar = new Panel { Dock = DockStyle.Top, Height = 36 };
        var lblSheet = new Label { Text = "시트:", Location = new Point(8, 9), AutoSize = true };
        cboSheet = new ComboBox { Location = new Point(42, 5), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList };
        lblStatus = new Label { Text = "초기화 중...", Location = new Point(235, 9), AutoSize = true, ForeColor = Color.Gray };
        topBar.Controls.AddRange(new Control[] { lblSheet, cboSheet, lblStatus });

        // ═══ 메인 탭 ═══
        tabMain = new TabControl { Dock = DockStyle.Fill };

        tabShipment = new TabPage("📦 출고/송장");
        tabPopular = new TabPage("📊 발주 많은 상품");
        tabProduct = new TabPage("📋 상품정보");
        tabStock = new TabPage("🧮 재고관리");

        BuildShipmentTab();
        BuildPlaceholderTab(tabPopular, "발주 많은 상품 분석 — 추후 구현 예정");
        BuildPlaceholderTab(tabProduct, "상품정보 관리 — 추후 구현 예정");
        BuildStockTab();

        tabMain.TabPages.AddRange(new[] { tabShipment, tabPopular, tabProduct, tabStock });

        // ═══ 로그 ═══
        var logPanel = new Panel { Dock = DockStyle.Bottom, Height = 140 };
        var logHeader = new Panel { Dock = DockStyle.Top, Height = 22 };
        logHeader.Controls.Add(new Label { Text = "로그", Dock = DockStyle.Left, Padding = new Padding(4, 3, 0, 0) });
        var btnClear = new Button { Text = "클리어", Dock = DockStyle.Right, Width = 60, Height = 22 };
        btnClear.Click += (_, _) => txtLog.Clear();
        logHeader.Controls.Add(btnClear);

        txtLog = new TextBox
        {
            Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both,
            ReadOnly = true, WordWrap = false,
            BackColor = Color.FromArgb(30, 30, 30), ForeColor = Color.LightGreen,
            Font = new Font("Consolas", 9f)
        };
        logPanel.Controls.Add(txtLog);
        logPanel.Controls.Add(logHeader);

        // ═══ 조립 ═══
        Controls.Add(tabMain);
        Controls.Add(new Panel { Dock = DockStyle.Top, Height = 1, BackColor = Color.LightGray });
        Controls.Add(topBar);
        Controls.Add(logPanel);
    }

    private void BuildShipmentTab()
    {
        // ── 좌측 패널: 발주사 체크리스트 ──
        var leftPanel = new Panel { Dock = DockStyle.Left, Width = 220, Padding = new Padding(4) };

        var lblVendor = new Label { Text = "발주사 (복수 선택)", Dock = DockStyle.Top, Height = 20 };
        clbVendors = new CheckedListBox
        {
            Dock = DockStyle.Fill, CheckOnClick = true,
            Font = new Font("맑은 고딕", 9f)
        };

        var vendorBtnPanel = new Panel { Dock = DockStyle.Bottom, Height = 28 };
        btnSelectAll = new Button { Text = "전체선택", Width = 70, Height = 26, Location = new Point(0, 0) };
        btnDeselectAll = new Button { Text = "전체해제", Width = 70, Height = 26, Location = new Point(74, 0) };
        vendorBtnPanel.Controls.AddRange(new Control[] { btnSelectAll, btnDeselectAll });

        leftPanel.Controls.Add(clbVendors);
        leftPanel.Controls.Add(vendorBtnPanel);
        leftPanel.Controls.Add(lblVendor);

        // ── 상단: 날짜 + 조회 + 택배사 ──
        var filterBar = new Panel { Dock = DockStyle.Top, Height = 40 };
        var lblFrom = new Label { Text = "조회기간:", Location = new Point(8, 11), AutoSize = true };
        dtpStart = new DateTimePicker
        {
            Location = new Point(75, 7), Width = 120, Format = DateTimePickerFormat.Short,
            Value = DateTime.Now.AddDays(-_cafe24Config.OrderFetchDays)
        };
        var lblTo = new Label { Text = "~", Location = new Point(200, 11), AutoSize = true };
        dtpEnd = new DateTimePicker
        {
            Location = new Point(215, 7), Width = 120, Format = DateTimePickerFormat.Short,
            Value = DateTime.Now
        };
        btnFetch = new Button { Text = "📥 조회", Location = new Point(345, 5), Width = 80, Height = 30 };
        var btnReset = new Button { Text = "🔄 초기화", Location = new Point(430, 5), Width = 80, Height = 30 };
        btnReset.Click += (_, _) => ResetAll();

        var lblShip = new Label { Text = "기본택배사:", Location = new Point(520, 11), AutoSize = true };
        cboShippingCompany = new ComboBox { Location = new Point(600, 7), Width = 160, DropDownStyle = ComboBoxStyle.DropDownList };
        foreach (var kv in Cafe24ApiClient.ShippingCompanyCodes)
            cboShippingCompany.Items.Add($"{kv.Key} ({kv.Value})");
        cboShippingCompany.SelectedIndex = 0;

        filterBar.Controls.AddRange(new Control[] { lblFrom, dtpStart, lblTo, dtpEnd, btnFetch, btnReset, lblShip, cboShippingCompany });

        // ── 서브탭: 데이터/매칭/결과 ──
        tabShipSub = new TabControl { Dock = DockStyle.Fill };

        // 서브탭1: 데이터
        var subData = new TabPage("데이터 프리뷰");
        dgvData = CreateGridView();
        var dataBtnPanel = new Panel { Dock = DockStyle.Top, Height = 36 };
        btnMatch = new Button { Text = "🔍 매칭 실행", Width = 130, Height = 30, Location = new Point(4, 3) };
        dataBtnPanel.Controls.Add(btnMatch);
        subData.Controls.Add(dgvData);
        subData.Controls.Add(dataBtnPanel);

        // 서브탭2: 매칭/검토
        var subMatch = new TabPage("매칭/검토");
        dgvMatch = CreateGridView();
        var matchBtnPanel = new Panel { Dock = DockStyle.Top, Height = 36 };
        btnPush = new Button { Text = "✅ 확정 항목 반영", Width = 150, Height = 30, Location = new Point(4, 3) };
        var btnConfirmAll = new Button { Text = "전체 확정", Width = 90, Height = 30, Location = new Point(160, 3) };
        btnConfirmAll.Click += (_, _) => ConfirmAllMatches();
        matchBtnPanel.Controls.AddRange(new Control[] { btnPush, btnConfirmAll });
        subMatch.Controls.Add(dgvMatch);
        subMatch.Controls.Add(matchBtnPanel);

        // 서브탭3: 반영 결과
        var subResult = new TabPage("반영 결과");
        dgvResult = CreateGridView();
        var resBtnPanel = new Panel { Dock = DockStyle.Top, Height = 36 };
        btnExportFailed = new Button { Text = "💾 실패/미매칭 Export", Width = 170, Height = 30, Location = new Point(4, 3) };
        resBtnPanel.Controls.Add(btnExportFailed);
        subResult.Controls.Add(dgvResult);
        subResult.Controls.Add(resBtnPanel);

        tabShipSub.TabPages.AddRange(new[] { subData, subMatch, subResult });

        // ── 조립 ──
        var rightPanel = new Panel { Dock = DockStyle.Fill };
        rightPanel.Controls.Add(tabShipSub);
        rightPanel.Controls.Add(filterBar);

        tabShipment.Controls.Add(rightPanel);
        tabShipment.Controls.Add(new Panel { Dock = DockStyle.Left, Width = 1, BackColor = Color.LightGray });
        tabShipment.Controls.Add(leftPanel);
    }


    private void BuildStockTab()
    {
        var top = new Panel { Dock = DockStyle.Top, Height = 42 };
        var lbl = new Label { Text = "재고 시트:", Location = new Point(8, 12), AutoSize = true };
        cboStockSheet = new ComboBox
        {
            Location = new Point(70, 8),
            Width = 360,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        btnStockLoad = new Button { Text = "📥 불러오기", Location = new Point(440, 6), Width = 100, Height = 30 };
        var lblInfo = new Label
        {
            Text = "기본 연결: 1HWR8zdvx0DYbl4ac9hmGuaaIA47nMO0v1CtO99PyC6w",
            Location = new Point(560, 12),
            AutoSize = true,
            ForeColor = Color.DimGray
        };
        top.Controls.AddRange(new Control[] { lbl, cboStockSheet, btnStockLoad, lblInfo });

        dgvStock = CreateGridView();
        dgvStock.ReadOnly = true;

        tabStock.Controls.Add(dgvStock);
        tabStock.Controls.Add(top);
    }

    private async Task InitStockTabAsync()
    {
        if (_sheetsReader == null || cboStockSheet == null) return;

        try
        {
            var sheets = await Task.Run(() => _sheetsReader.GetSheetList(StockSpreadsheetId));
            cboStockSheet.Items.Clear();

            foreach (var (title, sheetId) in sheets)
                cboStockSheet.Items.Add(new StockSheetItem(title, sheetId));

            var defaultIdx = sheets.FindIndex(s => s.sheetId == StockDefaultSheetGid);
            cboStockSheet.SelectedIndex = defaultIdx >= 0 ? defaultIdx : (cboStockSheet.Items.Count > 0 ? 0 : -1);

            if (cboStockSheet.SelectedIndex >= 0)
                await LoadStockSheetPreviewAsync();
        }
        catch (Exception ex)
        {
            _log.Error("재고 시트 초기화 실패", ex);
        }
    }

    private async Task LoadStockSheetPreviewAsync()
    {
        if (_sheetsReader == null) return;
        if (cboStockSheet.SelectedItem is not StockSheetItem item) return;

        try
        {
            btnStockLoad.Enabled = false;
            btnStockLoad.Text = "로딩 중...";

            var raw = await Task.Run(() => _sheetsReader.ReadRawSheet(StockSpreadsheetId, item.Title, 500));

            dgvStock.Columns.Clear();
            dgvStock.Rows.Clear();

            for (int i = 0; i < raw.Headers.Count; i++)
                dgvStock.Columns.Add($"C{i}", raw.Headers[i]);

            foreach (var row in raw.Rows)
                dgvStock.Rows.Add(row.Cast<object>().ToArray());

            _log.Info($"재고관리 시트 '{item.Title}' 프리뷰 완료: {raw.Rows.Count}행");
        }
        catch (Exception ex)
        {
            _log.Error("재고 시트 로드 실패", ex);
            MessageBox.Show($"재고 시트 로드 오류:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnStockLoad.Enabled = true;
            btnStockLoad.Text = "📥 불러오기";
        }
    }

    private sealed class StockSheetItem
    {
        public string Title { get; }
        public int SheetId { get; }

        public StockSheetItem(string title, int sheetId)
        {
            Title = title;
            SheetId = sheetId;
        }

        public override string ToString() => $"{Title} (gid:{SheetId})";
    }
    private void BuildPlaceholderTab(TabPage tab, string message)
    {
        var lbl = new Label
        {
            Text = $"🚧 {message}",
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("맑은 고딕", 14f),
            ForeColor = Color.Gray
        };
        tab.Controls.Add(lbl);
    }

    private DataGridView CreateGridView()
    {
        return new DataGridView
        {
            Dock = DockStyle.Fill, ReadOnly = false,
            AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor = Color.White, RowHeadersVisible = false,
            Font = new Font("맑은 고딕", 9f)
        };
    }

    private void WireEvents()
    {
        cboSheet.SelectedIndexChanged += async (_, _) => await LoadVendorsAsync();
        btnSelectAll.Click += (_, _) => { for (int i = 0; i < clbVendors.Items.Count; i++) clbVendors.SetItemChecked(i, true); };
        btnDeselectAll.Click += (_, _) => { for (int i = 0; i < clbVendors.Items.Count; i++) clbVendors.SetItemChecked(i, false); };
        btnFetch.Click += async (_, _) => await FetchFilteredDataAsync();
        btnMatch.Click += async (_, _) => await ExecuteMatchingAsync();
        btnPush.Click += async (_, _) => await ExecutePushAsync();
        btnExportFailed.Click += (_, _) => ExportFailed();
        if (btnStockLoad != null) btnStockLoad.Click += async (_, _) => await LoadStockSheetPreviewAsync();
    }

    // ═══════════════════════════════════════
    // 조회: 선택 발주사 데이터 가져오기
    // ═══════════════════════════════════════
    private async Task FetchFilteredDataAsync()
    {
        try
        {
            btnFetch.Enabled = false;
            btnFetch.Text = "Cafe24 조회 중...";

            // ① Cafe24 배송준비중 주문 수집
            var progress = new Progress<string>(msg => _log.Info(msg));
            _cafe24Orders = await _api.FetchRecentOrders(
                dtpStart.Value, dtpEnd.Value, progress, orderStatus: "N20");

            if (_cafe24Orders.Count == 0)
            {
                lblStatus.Text = "⚠️ 배송준비중 주문이 없습니다.";
                lblStatus.ForeColor = Color.DarkOrange;
                _log.Info("배송준비중 주문 0건");
                return;
            }

            // DB 캐시 갱신
            _db.ClearOrderCache();
            foreach (var o in _cafe24Orders)
                _db.InsertOrderCache(o);

            // ② 데이터 프리뷰에 Cafe24 주문 표시
            ShowDataPreview();

            lblStatus.Text = $"✅ Cafe24 배송준비중 주문 {_cafe24Orders.Count}건 조회 완료";
            lblStatus.ForeColor = Color.DarkGreen;
            _log.Info($"Cafe24 배송준비중 주문 {_cafe24Orders.Count}건 캐시 완료");

            tabShipSub.SelectedIndex = 0;
        }
        catch (Exception ex)
        {
            _log.Error("Cafe24 주문 조회 실패", ex);
            MessageBox.Show($"조회 오류:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnFetch.Enabled = true;
            btnFetch.Text = "📥 조회";
        }
    }

    private void ResetAll()
    {
        _excelResult = null;
        _filteredRows = new();
        _cafe24Orders = new();
        _matchResults = new();

        dgvData.Columns.Clear(); dgvData.Rows.Clear();
        dgvMatch.Columns.Clear(); dgvMatch.Rows.Clear();
        dgvResult.Columns.Clear(); dgvResult.Rows.Clear();

        UpdateOrderSelectionSummaryEx();

        tabShipSub.SelectedIndex = 0;
        lblStatus.Text = "🔄 초기화 완료";
        lblStatus.ForeColor = Color.Gray;
        _log.Info("초기화 완료");
    }
    private void ShowDataPreview()
    {
        dgvData.Columns.Clear();
        dgvData.Rows.Clear();

        EnsureOrderSelectionColumnEx();

        dgvData.Columns.Add("No", "#");
        dgvData.Columns.Add("OrderId", "주문번호");
        dgvData.Columns.Add("OrderDate", "주문일");
        dgvData.Columns.Add("ProductName", "상품명");
        dgvData.Columns.Add("Qty", "수량");
        dgvData.Columns.Add("Name", "수령인명");
        dgvData.Columns.Add("Phone", "전화번호");
        dgvData.Columns.Add("OrderStatus", "주문상태");

        dgvData.Columns[OrderSelectColumnNameEx]!.Width = 50;
        dgvData.Columns["No"]!.Width = 40;
        dgvData.Columns["OrderId"]!.Width = 130;
        dgvData.Columns["OrderDate"]!.Width = 90;
        dgvData.Columns["ProductName"]!.Width = 200;
        dgvData.Columns["Qty"]!.Width = 45;
        dgvData.Columns["Name"]!.Width = 75;
        dgvData.Columns["Phone"]!.Width = 110;
        dgvData.Columns["OrderStatus"]!.Width = 80;

        for (int i = 0; i < _cafe24Orders.Count; i++)
        {
            var o = _cafe24Orders[i];
            var phone = string.IsNullOrWhiteSpace(o.RecipientCellPhone) ? o.RecipientPhone : o.RecipientCellPhone;
            var rowIndex = dgvData.Rows.Add(false, i + 1, o.OrderId, o.OrderDate, o.ProductName,
                o.Quantity, o.RecipientName, phone, o.OrderStatus);
            dgvData.Rows[rowIndex].Tag = o;
        }

        UpdateOrderSelectionSummaryEx();
    }
    // ═══════════════════════════════════════
    // 매칭 실행
    // ═══════════════════════════════════════
    private async Task ExecuteMatchingAsync()
    {
        if (_cafe24Orders.Count == 0) { MessageBox.Show("먼저 Cafe24 주문을 조회하세요.", "알림"); return; }
        if (_sheetsReader == null) { MessageBox.Show("시트 연결이 안 되어 있습니다.", "알림"); return; }

        var sheetName = cboSheet.SelectedItem?.ToString();
        if (string.IsNullOrEmpty(sheetName)) { MessageBox.Show("시트를 선택하세요.", "알림"); return; }

        try
        {
            btnMatch.Enabled = false;
            btnMatch.Text = "스프레드시트 로드 중...";

            // ③ 스프레드시트 로드 (발주사 선택 시 필터 적용)
            var selectedVendors = new HashSet<string>();
            foreach (var item in clbVendors.CheckedItems)
                selectedVendors.Add(NormalizeVendorLabel(item.ToString() ?? ""));

            // UI 컨트롤 값을 미리 캡처 (크로스스레드 방지)
            var startDate = dtpStart.Value;
            var endDate = dtpEnd.Value;

            _excelResult = await Task.Run(() =>
                selectedVendors.Count > 0
                    ? _sheetsReader.ReadSheetFiltered(_spreadsheetId, sheetName, selectedVendors, startDate, endDate)
                    : _sheetsReader.ReadSheet(_spreadsheetId, sheetName));

            _filteredRows = _excelResult.Rows;
            _log.Info($"스프레드시트 '{sheetName}' 로드: {_filteredRows.Count}행" +
                (selectedVendors.Count > 0 ? $" (발주사 {selectedVendors.Count}개 필터)" : " (전체)"));

            // DB 저장: 필터된 행만 단일 트랜잭션으로 처리
            await Task.Run(() => _db.BulkInsertSourceRows(_filteredRows));

            btnMatch.Text = "매칭 중...";

            // 역방향 매칭: Cafe24 주문 기준 → 스프레드시트 검색
            _matchResults = await Task.Run(() => _matcher.ExecuteReverseMatching(_cafe24Orders, _filteredRows));
            foreach (var mr in _matchResults)
            {
                if (mr.SourceRowId > 0)
                {
                    try { mr.Id = _db.InsertMatchResult(mr); }
                    catch { /* FK 무시 - 메모리 결과로 진행 */ }
                }
            }

            ShowMatchResults();
            tabShipSub.SelectedIndex = 1;
        }
        catch (Exception ex)
        {
            _log.Error("매칭 실행 오류", ex);
            MessageBox.Show($"매칭 오류:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnMatch.Enabled = true;
            btnMatch.Text = "🔍 매칭 실행";
        }
    }

    private void ShowMatchResults()
    {
        dgvMatch.Columns.Clear();
        dgvMatch.Rows.Clear();

        dgvMatch.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Confirm", HeaderText = "확정", Width = 45 });
        dgvMatch.Columns.Add("MrId", "ID");
        dgvMatch.Columns.Add("SrcPhone", "출고-휴대폰");
        dgvMatch.Columns.Add("SrcName", "출고-수령인");
        dgvMatch.Columns.Add("SrcTracking", "송장번호");
        dgvMatch.Columns.Add("Cafe24Order", "Cafe24 주문번호");
        dgvMatch.Columns.Add("OrdPhone", "주문-휴대폰");
        dgvMatch.Columns.Add("OrdName", "주문-수령인");
        dgvMatch.Columns.Add("OrdProduct", "상품명");
        dgvMatch.Columns.Add("Confidence", "확신도");
        dgvMatch.Columns.Add("MStatus", "상태");

        dgvMatch.Columns["MrId"]!.Width = 40;
        dgvMatch.Columns["Confidence"]!.Width = 70;
        dgvMatch.Columns["MStatus"]!.Width = 80;

        for (int i = 0; i < _matchResults.Count; i++)
        {
            var mr = _matchResults[i];
            bool auto = mr.MatchStatus == "auto_confirmed";
            var idx = dgvMatch.Rows.Add(auto, i, mr.SourcePhone, mr.SourceName,
                mr.SourceTracking, mr.Cafe24OrderId, mr.OrderPhone, mr.OrderName,
                mr.OrderProduct, ConfLabel(mr.Confidence), StatLabel(mr.MatchStatus));

            dgvMatch.Rows[idx].DefaultCellStyle.BackColor = mr.Confidence switch
            {
                "exact" => Color.FromArgb(220, 255, 220),
                "probable" => Color.FromArgb(255, 255, 210),
                "candidate" => Color.FromArgb(255, 240, 200),
                "no_tracking" => Color.FromArgb(255, 230, 180),
                _ => Color.FromArgb(255, 220, 220)
            };
        }

        _log.Info($"매칭: 전체 {_matchResults.Count} | " +
            $"확정 {_matchResults.Count(m => m.Confidence == "exact")} | " +
            $"유력 {_matchResults.Count(m => m.Confidence == "probable")} | " +
            $"후보 {_matchResults.Count(m => m.Confidence == "candidate")} | " +
            $"송장없음 {_matchResults.Count(m => m.Confidence == "no_tracking")} | " +
            $"미매칭 {_matchResults.Count(m => m.Confidence == "none")}");
    }

    static string ConfLabel(string c) => c switch { "exact" => "✅ 확정", "probable" => "🟡 유력", "candidate" => "🟠 후보", "no_tracking" => "⚠️ 송장없음", _ => "❌ 없음" };
    static string StatLabel(string s) => s switch { "auto_confirmed" => "자동확정", "confirmed" => "수동확정", "pending" => "대기", "pushed" => "반영완료", "push_failed" => "반영실패", "unmatched" => "미매칭", "no_tracking" => "송장미등록", _ => s };

    private void ConfirmAllMatches()
    {
        for (int i = 0; i < dgvMatch.Rows.Count; i++)
        {
            var c = dgvMatch.Rows[i].Cells["Confidence"]?.Value?.ToString() ?? "";
            // 미매칭(없음)만 제외하고 전부 체크
            if (!c.Contains("없음"))
                dgvMatch.Rows[i].Cells["Confirm"].Value = true;
        }
    }

    // ═══════════════════════════════════════
    // 반영 실행
    // ═══════════════════════════════════════
    private async Task ExecutePushAsync()
    {
        var confirmed = new List<MatchResult>();
        for (int i = 0; i < dgvMatch.Rows.Count; i++)
        {
            if (dgvMatch.Rows[i].Cells["Confirm"]?.Value is true)
            {
                var mrIdx = Convert.ToInt32(dgvMatch.Rows[i].Cells["MrId"]?.Value);
                if (mrIdx >= 0 && mrIdx < _matchResults.Count)
                {
                    var mr = _matchResults[mrIdx];
                    if (!string.IsNullOrEmpty(mr.Cafe24OrderId))
                        confirmed.Add(mr);
                }
            }
        }

        if (confirmed.Count == 0) { MessageBox.Show("확정된 항목이 없습니다.", "알림"); return; }
        if (MessageBox.Show($"{confirmed.Count}건을 Cafe24에 반영합니다.\n계속?",
            "반영 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

        var shipCode = _cafe24Config.DefaultShippingCompanyCode;
        if (cboShippingCompany.SelectedItem != null)
        {
            var sel = cboShippingCompany.SelectedItem.ToString()!;
            var s = sel.LastIndexOf('(') + 1; var e = sel.LastIndexOf(')');
            if (s > 0 && e > s) shipCode = sel[s..e];
        }

        btnPush.Enabled = false; btnPush.Text = "반영 중...";

        dgvResult.Columns.Clear(); dgvResult.Rows.Clear();
        dgvResult.Columns.Add("OrderId", "주문번호");
        dgvResult.Columns.Add("Tracking", "송장번호");
        dgvResult.Columns.Add("Result", "결과");
        dgvResult.Columns.Add("Detail", "상세");

        int ok = 0, fail = 0;
        foreach (var mr in confirmed)
        {
            var src = _filteredRows.FirstOrDefault(r => r.Id == mr.SourceRowId);
            var tracking = src?.TrackingNumber ?? "";
            var code = !string.IsNullOrEmpty(src?.ShippingCompany) ? ResolveShipCode(src.ShippingCompany) : shipCode;

            if (string.IsNullOrEmpty(tracking)) { dgvResult.Rows.Add(mr.Cafe24OrderId, "", "SKIP", "송장번호 없음"); continue; }

            var (success, resp, status) = await _api.PushTrackingNumber(mr.Cafe24OrderId, mr.Cafe24OrderItemCode, tracking, code);

            _db.InsertPushLog(new PushLog
            {
                MatchResultId = mr.Id, Cafe24OrderId = mr.Cafe24OrderId,
                RequestBody = JsonConvert.SerializeObject(new { tracking, code }),
                ResponseBody = resp, HttpStatusCode = status,
                Result = success ? "success" : "fail",
                ErrorMessage = success ? "" : resp,
                PushedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            });

            var ns = success ? "pushed" : "push_failed";
            _db.UpdateMatchStatus(mr.Id, ns, true);
            _db.UpdateSourceRowStatus(mr.SourceRowId, success ? "pushed" : "failed", mr.Cafe24OrderId);

            var idx = dgvResult.Rows.Add(mr.Cafe24OrderId, tracking,
                success ? "✅ 성공" : $"❌ 실패 ({status})", resp.Length > 200 ? resp[..200] : resp);
            dgvResult.Rows[idx].DefaultCellStyle.BackColor = success ? Color.FromArgb(220, 255, 220) : Color.FromArgb(255, 220, 220);
            if (success) ok++; else fail++;
        }

        _log.Info($"반영 완료: 성공 {ok}, 실패 {fail}");
        tabShipSub.SelectedIndex = 2;
        btnPush.Enabled = true; btnPush.Text = "✅ 확정 항목 반영";
    }

    private string ResolveShipCode(string name)
    {
        foreach (var kv in Cafe24ApiClient.ShippingCompanyCodes)
            if (name.Contains(kv.Key, StringComparison.OrdinalIgnoreCase)) return kv.Value;
        return _cafe24Config.DefaultShippingCompanyCode;
    }

    // ═══════════════════════════════════════
    // Export
    // ═══════════════════════════════════════
    private void ExportFailed()
    {
        var failed = _matchResults.Where(m => m.MatchStatus is "unmatched" or "push_failed" or "pending").ToList();
        if (failed.Count == 0) { MessageBox.Show("미매칭/실패 항목이 없습니다.", "알림"); return; }

        using var sfd = new SaveFileDialog { Filter = "CSV|*.csv", FileName = $"failed_{DateTime.Now:yyyyMMdd_HHmmss}.csv" };
        if (sfd.ShowDialog() != DialogResult.OK) return;

        var sb = new StringBuilder();
        sb.AppendLine("SourceRowId,출고-휴대폰,출고-수령인,송장번호,Cafe24주문번호,주문-수령인,확신도,상태");
        foreach (var m in failed)
            sb.AppendLine($"{m.SourceRowId},{m.SourcePhone},{m.SourceName},{m.SourceTracking},{m.Cafe24OrderId},{m.OrderName},{m.Confidence},{m.MatchStatus}");

        File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
        _log.Info($"Export: {sfd.FileName} ({failed.Count}건)");
        MessageBox.Show($"{failed.Count}건 저장:\n{sfd.FileName}", "Export");
    }
}

// ═══════════════════════════════════════
// 컬럼 선택 다이얼로그
// ═══════════════════════════════════════
public class ColumnSelectDialog : Form
{
    public int SelectedIndex { get; private set; } = -1;
    private readonly ListBox _listBox;

    public ColumnSelectDialog(List<string> headers)
    {
        Text = "수령인 휴대폰 컬럼 선택"; Size = new Size(350, 400);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false;

        var label = new Label { Text = "휴대폰 컬럼을 선택하세요:", Dock = DockStyle.Top, Height = 30, Padding = new Padding(8) };
        _listBox = new ListBox { Dock = DockStyle.Fill };
        for (int i = 0; i < headers.Count; i++)
            _listBox.Items.Add($"[{(char)('A' + (i < 26 ? i : 0))}열] {headers[i]}");

        var btnOk = new Button { Text = "선택", DialogResult = DialogResult.OK, Dock = DockStyle.Bottom, Height = 35 };
        btnOk.Click += (_, _) => { SelectedIndex = _listBox.SelectedIndex; };

        Controls.Add(_listBox); Controls.Add(label); Controls.Add(btnOk);
        AcceptButton = btnOk;
    }
}


