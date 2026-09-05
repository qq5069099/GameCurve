using System.Globalization;
using GameCurve.Excel;
using GameCurve.Models;
using GameCurve.Services;

namespace GameCurve.Ui;

public sealed class MainForm : Form
{
    private const int LeftPanelWidth = 200;   // 左侧功能面板宽度
    private const int RightPanelWidth = 260;  // 可完整容纳编辑/批量/统计控件

    private readonly CurveEditor _curve = new() { Dock = DockStyle.Fill };
    private readonly DataGridView _grid = new() { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false };
    private readonly ToolTip _tip = new() { AutoPopDelay = 12000, InitialDelay = 500, ReshowDelay = 120 };

    private readonly ToolStrip _tool = new();
    private readonly ToolStripButton _autoSaveCheck = new("自动保存") { Checked = false, CheckOnClick = true };
    private readonly ToolStripDropDownButton _openMenu = new("打开");
    private readonly ToolStripDropDownButton _layoutButton = new("布局");
    private readonly StatusStrip _status = new();
    private readonly ToolStripStatusLabel _statusLabel = new() { Text = "就绪" };
    private readonly ToolStripStatusLabel _hoverLabel = new() { Spring = true, TextAlign = ContentAlignment.MiddleLeft };

    // 底部工作表标签
    private readonly FlowLayoutPanel _sheetStrip = new()
    {
        Dock = DockStyle.Bottom,
        FlowDirection = FlowDirection.LeftToRight,
        WrapContents = false,
        AutoScroll = true,
        Height = 46,
        Padding = new Padding(2),
        BackColor = Color.FromArgb(226, 230, 236)
    };
    private readonly List<Button> _sheetButtons = new();

    // 左：列选择面板
    private readonly CheckedListBox _colsChecked = new() { CheckOnClick = true, BorderStyle = BorderStyle.FixedSingle };
    private readonly ComboBox _xCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList };

    // 右：编辑/批量/统计
    private readonly Label _selInfo = new() { AutoSize = true, Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold) };
    private readonly NumericUpDown _valUpDown = new() { DecimalPlaces = 6, Minimum = -1.0E12m, Maximum = 1.0E12m, Width = 215 };
    private readonly NumericUpDown _stepUpDown = new() { DecimalPlaces = 6, Minimum = 1.0E-8m, Maximum = 1.0E6m, Value = 1m, Width = 215 };
    private readonly NumericUpDown _offsetUpDown = new() { DecimalPlaces = 6, Minimum = -1.0E12m, Maximum = 1.0E12m, Width = 215 };
    private readonly NumericUpDown _scaleUpDown = new() { DecimalPlaces = 6, Minimum = -1.0E6m, Maximum = 1.0E6m, Value = 1m, Width = 215 };
    private readonly NumericUpDown _clampMin = new() { DecimalPlaces = 6, Minimum = -1.0E12m, Maximum = 1.0E12m, Width = 215 };
    private readonly NumericUpDown _clampMax = new() { DecimalPlaces = 6, Minimum = -1.0E12m, Maximum = 1.0E12m, Value = 100m, Width = 215 };
    private readonly NumericUpDown _randUpDown = new() { DecimalPlaces = 6, Minimum = 0m, Maximum = 1.0E9m, Value = 5m, Width = 215 };
    private readonly Label _statLabel = new()
    {
        AutoSize = false,
        Height = 96,
        Font = new Font("Microsoft YaHei UI", 7.5f),
        ForeColor = Color.FromArgb(70, 76, 84)
    };

    private SplitContainer _chartGridSplit = null!;
    private TableLayoutPanel _chartArea = null!;
    private readonly ContextMenuStrip _menu = new();
    private Panel _gridPane = null!;
    private readonly Panel _rightPanel = new() { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(6) };
    private readonly Panel _leftPanel = new() { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(6) };

    private WorkbookModel? _wb;
    private SheetSnapshot? _snapshot;
    private List<CurveColumnOption> _checkedCols = new();
    private CurveColumnOption? _xColumn;
    private CurveColumnOption? _activeYColumn;
    private readonly List<CurveSeriesView> _series = new();

    private readonly Dictionary<int, (double X, double Y)> _committed = new();
    private readonly Dictionary<int, (double X, double Y)> _editing = new();
    private readonly Dictionary<int, int> _rowToGridIndex = new();
    private readonly HashSet<(int Col, int Row)> _dirtyCells = new();
    private readonly HashSet<int> _pendingUndoRows = new();

    private bool _gridAtRight;
    private bool _suppressRebuild;
    private bool _syncFromGrid;
    private bool _updatingGridSelection;
    private string _activeSheet = "";
    private string _autoFocusSheet = "";
    private int _autoFocusCol = -1;

    private readonly System.Windows.Forms.Timer _autoSaveTimer = new() { Interval = 600 };
    private readonly System.Windows.Forms.Timer _reloadTimer = new() { Interval = 600 };
    private FileSystemWatcher? _watcher;
    private bool _selfWrite;

    private readonly string? _startupFile;
    private readonly string _settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GameCurve", "layout.json");
    private LayoutSettings _settings = new();
    private sealed record EditCmd(List<(int Col, int Row, string Old, string New)> Cells);
    private readonly List<EditCmd> _undo = new();
    private readonly List<EditCmd> _redo = new();

    public MainForm(string? startupFile = null)
    {
        _startupFile = startupFile;
        LoadSettings();
        Text = "GameCurve - 游戏数据曲线编辑器";
        Width = 1440; Height = 920;
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(1150, 720);
        Font = new Font("Microsoft YaHei UI", 9f);

        BuildUi();
        BuildEvents();
        Shown += (s, e) => OnShown();
    }

    private sealed class LayoutSettings
    {
        public int LeftWidth { get; set; } = LeftPanelWidth;
        public int RightWidth { get; set; } = RightPanelWidth;
        public int SplitterDistance { get; set; } = -1;
        public bool GridAtRight { get; set; }
        public bool Maximized { get; set; } = true;
        public string LastFile { get; set; } = "";
    }

    private void LoadSettings()
    {
        try
        {
            if (File.Exists(_settingsPath))
                _settings = System.Text.Json.JsonSerializer.Deserialize<LayoutSettings>(File.ReadAllText(_settingsPath)) ?? new LayoutSettings();
        }
        catch { _settings = new LayoutSettings(); }
    }

    private void SaveSettings()
    {
        _settings.LeftWidth = (int)_chartArea.ColumnStyles[0].Width;
        _settings.RightWidth = (int)_chartArea.ColumnStyles[2].Width;
        _settings.GridAtRight = _gridAtRight;
        _settings.Maximized = WindowState == FormWindowState.Maximized;
        _settings.SplitterDistance = _chartGridSplit.Height > 0 ? _chartGridSplit.SplitterDistance : -1;
        _settings.LastFile = _wb?.Path ?? _settings.LastFile;
        try
        {
            var dir = Path.GetDirectoryName(_settingsPath)!;
            Directory.CreateDirectory(dir);
            File.WriteAllText(_settingsPath, System.Text.Json.JsonSerializer.Serialize(_settings));
        }
        catch { }
    }

    private void BuildUi()
    {
        _tool.GripStyle = ToolStripGripStyle.Hidden;
        _openMenu.DropDownOpening += (s, e) => RebuildOpenMenu();
        _tool.Items.Add(_openMenu);
        AddButton("保存", "把当前改动写回 Excel 文件（Ctrl+S）", OnSave);
        AddButton("另存为", "复制一份并另存为新文件", OnSaveAs);
        AddButton("刷新", "重新从磁盘读取当前工作表", OnReload);
        _tool.Items.Add(new ToolStripSeparator());
        AddButton("撤销", "撤销上次编辑（Ctrl+Z）", () => Undo());
        AddButton("重做", "重做上次撤销（Ctrl+Y）", () => Redo());
        _tool.Items.Add(new ToolStripSeparator());
        AddButton("导出PNG", "把当前曲线图导出为 PNG 图片", OnExport);
        _tool.Items.Add(_autoSaveCheck);
        _layoutButton.DropDownItems.Add(MakeMenu("表格置底", "表格放在曲线区下方", () => SetGridSide(false), true));
        _layoutButton.DropDownItems.Add(MakeMenu("表格靠右", "表格与曲线区左右并列", () => SetGridSide(true), false));
        _tool.Items.Add(_layoutButton);
        _tool.Dock = DockStyle.Top;
        Controls.Add(_tool);

        // 左：列选择
        int lw = LeftPanelWidth - 16;
        int rw = RightPanelWidth - 16;
        int ly = 30;
        bool compactAfterSection = false;
        void L(Control c, int h = 0)
        {
            int gap = compactAfterSection ? 2 : 6;
            compactAfterSection = false;
            if (c.AutoSize)
            {
                c.Location = new Point(8, ly);
                c.Width = lw;
                _leftPanel.Controls.Add(c);
                ly += Math.Max(22, c.GetPreferredSize(Size.Empty).Height) + gap;
            }
            else
            {
                int hh = h == 0 ? (c.Height > 0 ? c.Height : 26) : h;
                c.SetBounds(8, ly, lw, hh);
                _leftPanel.Controls.Add(c);
                ly += hh + gap;
            }
        }
        L(Section("曲线列 (Y) 多选"));
        compactAfterSection = true;
        L(MakeButton("清空所有选择", () =>
        {
            for (int i = 0; i < _colsChecked.Items.Count; i++)
                _colsChecked.SetItemChecked(i, false);
        }), 28);
        L(_colsChecked, 475);
        ly += 8;
        L(Section("X 轴"));
        compactAfterSection = true;
        L(_xCombo, 26);
        L(MakeHint("行号 或 选某数值列作为 X（此时可拖动横移该列）"));

        // 右：编辑/批量/统计（功能面板，始终在最右侧）
        int ry = 30;
        void R(Control c, int h = 0)
        {
            if (c.AutoSize)
            {
                c.Location = new Point(8, ry);
                c.Width = rw;
                _rightPanel.Controls.Add(c);
                ry += Math.Max(22, c.GetPreferredSize(Size.Empty).Height) + 6;
            }
            else
            {
                int hh = h == 0 ? (c.Height > 0 ? c.Height : 26) : h;
                c.SetBounds(8, ry, rw, hh);
                _rightPanel.Controls.Add(c);
                ry += hh + 6;
            }
        }
        R(_selInfo, 24);
        R(MakeLabel("选中点数值:"), 20);
        R(_valUpDown, 28);
        R(MakeButton("应用值", OnApplyValue), 30);
        R(MakeLabel("键盘微调步长:"), 20);
        R(_stepUpDown, 28);
        ry += 10;
        R(Section("批量操作"));
        R(MakeLabel("偏移 (+/-):"), 20);
        R(_offsetUpDown, 28);
        R(MakeButton("偏移 Δ", () => BatchOffset((double)_offsetUpDown.Value)), 30);
        R(MakeLabel("缩放 (×):"), 20);
        R(_scaleUpDown, 28);
        R(MakeButton("缩放 ×", () => BatchScale((double)_scaleUpDown.Value)), 30);
        R(MakeLabel("钳制 最小 / 最大:"), 20);
        R(_clampMin, 28);
        R(_clampMax, 28);
        R(MakeButton("钳制", () => BatchClamp((double)_clampMin.Value, (double)_clampMax.Value)), 30);
        R(MakeButton("整列平滑", BatchSmooth), 30);
        R(MakeButton("整列归一化", BatchNormalize), 30);
        R(MakeLabel("随机扰动幅度:"), 20);
        R(_randUpDown, 28);
        R(MakeButton("随机扰动", () => BatchRandom((double)_randUpDown.Value)), 30);
        R(MakeButton("右键更多操作", OpenContextAtChart), 30);
        ry += 10;
        R(Section("统计"));
        R(_statLabel, 132);

        // 顶部三栏：左侧选择 | 曲线 | 右侧功能面板（功能面板只在这一栏，不遮挡下方表格）
        int leftCol = Math.Clamp(_settings.LeftWidth, LeftPanelWidth, 460);
        int rightCol = Math.Clamp(_settings.RightWidth, RightPanelWidth, 460);
        _chartArea = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1 };
        _chartArea.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, leftCol));
        _chartArea.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        _chartArea.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, rightCol));
        _chartArea.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        _chartArea.Controls.Add(_leftPanel, 0, 0);
        _chartArea.Controls.Add(_curve, 1, 0);
        _chartArea.Controls.Add(_rightPanel, 2, 0);

        // 面板宽度变化时同步内部控件宽度，避免调窄后出现横向滚动条
        static void FitPanelWidth(Panel p)
        {
            int w = Math.Max(2, p.ClientSize.Width - 16);
            foreach (Control c in p.Controls)
            {
                if (c.AutoSize) continue;
                c.Width = w;
            }
        }
        _leftPanel.Resize += (s, e) => FitPanelWidth(_leftPanel);
        _leftPanel.ClientSizeChanged += (s, e) => FitPanelWidth(_leftPanel);
        _rightPanel.Resize += (s, e) => FitPanelWidth(_rightPanel);
        _rightPanel.ClientSizeChanged += (s, e) => FitPanelWidth(_rightPanel);

        // 数据表格 + 底部工作表标签（作为一个整体，独占下方整行）
        _gridPane = new Panel { Dock = DockStyle.Fill };
        _gridPane.Controls.Add(_grid);
        _gridPane.Controls.Add(_sheetStrip);

        // 顶部（功能面板区域） 与 表格 之间可拖动分隔，表格可置底/置右
        _chartGridSplit = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            Panel1MinSize = 120,
            Panel2MinSize = 120,
            SplitterWidth = 6
        };
        _chartGridSplit.Panel1.Controls.Add(_chartArea);
        _chartGridSplit.Panel2.Controls.Add(_gridPane);
        Controls.Add(_chartGridSplit);

        // 状态栏（最底部）
        _status.Items.Add(_statusLabel);
        _status.Items.Add(_hoverLabel);
        _status.Dock = DockStyle.Bottom;
        Controls.Add(_status);

        _grid.AllowUserToResizeRows = true;
        _grid.AllowUserToResizeColumns = true;
        _grid.RowHeadersVisible = true;
        _grid.EditMode = DataGridViewEditMode.EditOnEnter;
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
        _grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
        _grid.BackgroundColor = Color.White;
        _grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

        _menu.ShowImageMargin = false;
        _curve.ContextMenuStrip = _menu;

        BuildTooltips();
    }

    private void BuildEvents()
    {
        _colsChecked.ItemCheck += (s, e) => BeginInvoke((Action)OnSelectionChanged);
        _xCombo.SelectedIndexChanged += (s, e) => OnSelectionChangedSafe();
        _stepUpDown.ValueChanged += (s, e) => _curve.KeyboardStep = (double)_stepUpDown.Value;

        _curve.PointsChanged += OnPointsChanged;
        _curve.EditCommitted += OnEditCommitted;
        _curve.SelectionChanged += OnSelectionChangedUi;
        _curve.HoverChanged += s => _hoverLabel.Text = s;
        _grid.CellEndEdit += OnGridCellEdited;
        _grid.SelectionChanged += OnGridSelectionChanged;
        _menu.Opening += (s, e) => BuildContextMenu();

        _autoSaveTimer.Tick += (s, e) => { _autoSaveTimer.Stop(); CommitPending(); };
        _reloadTimer.Tick += (s, e) => { _reloadTimer.Stop(); ReloadFromDisk(); };

        KeyPreview = true;
        KeyDown += (s, e) =>
        {
            if (e.Control && e.KeyCode == Keys.S) { OnSave(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.Z) { Undo(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.Y) { Redo(); e.Handled = true; }
        };
    }

    // ---------- 打开 / 工作表 ----------
    private void OnShown()
    {
        WindowState = _settings.Maximized ? FormWindowState.Maximized : FormWindowState.Normal;
        _chartArea.ColumnStyles[0].Width = Math.Clamp(_settings.LeftWidth, LeftPanelWidth, 460);
        _chartArea.ColumnStyles[2].Width = Math.Clamp(_settings.RightWidth, RightPanelWidth, 460);
        SetGridSide(_settings.GridAtRight);
        int min = Math.Max(_chartGridSplit.Panel1MinSize, 120);
        int max = Math.Max(min, (int)((_settings.GridAtRight ? _chartGridSplit.Width : _chartGridSplit.Height)) - Math.Max(_chartGridSplit.Panel2MinSize, 120));
        int dist = _settings.SplitterDistance > 0
            ? Math.Max(min, Math.Min(_settings.SplitterDistance, max))
            : Math.Max(min, (int)((_settings.GridAtRight ? _chartGridSplit.Width : _chartGridSplit.Height) * 0.58));
        try { _chartGridSplit.SplitterDistance = dist; } catch { }
        if (File.Exists(_startupFile)) OpenFile(_startupFile!);
        else if (File.Exists(_settings.LastFile)) OpenFile(_settings.LastFile!);
    }

    private void OnOpen()
    {
        using var ofd = new OpenFileDialog
        {
            Title = "选择游戏数据工作簿",
            Filter = "Excel 工作簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|所有文件 (*.*)|*.*"
        };
        if (ofd.ShowDialog(this) != DialogResult.OK) return;
        OpenFile(ofd.FileName);
    }

    private void OpenFile(string path)
    {
        _wb?.Dispose();
        _wb = new WorkbookModel();
        try { _wb.Open(path); }
        catch (Exception ex)
        {
            MessageBox.Show(this, "无法打开文件：\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        _settings.LastFile = path;
        SetupWatcher(path);
        BuildSheetStrip();
        var best = FindBestDataColumn();
        var target = best?.Sheet ?? _wb.SheetNames.FirstOrDefault(s => !ShouldHideSheet(s)) ?? "";
        _autoFocusSheet = best?.Sheet ?? "";
        _autoFocusCol = best?.Col.ColumnIndex ?? -1;
        ActivateSheet(target);
        UpdateTitle();
    }

    /// <summary>判断是否为“序号 / 主键”列，默认选中时应跳过。</summary>
    private static bool IsOrdinalColumn(ColumnMeta c)
    {
        if (c.Name.Equals("ID", StringComparison.OrdinalIgnoreCase)) return true;
        string t = $"{c.Name} {c.Label} {c.HeaderRaw}";
        return t.Contains("序号") || t.Contains("编号");
    }

    /// <summary>名含连续三个 1（111）的列不显示。</summary>
    private static bool ShouldHideColumn(ColumnMeta c)
        => $"{c.Name} {c.Label} {c.HeaderRaw}".Contains("111", StringComparison.OrdinalIgnoreCase);

    /// <summary>名字含连续三个 1（111）的工作表不显示。</summary>
    private static bool ShouldHideSheet(string name)
        => name.Contains("111", StringComparison.OrdinalIgnoreCase);

    private (string Sheet, ColumnMeta Col)? FindBestDataColumn()
    {
        if (_wb == null) return null;
        (string Sheet, ColumnMeta Col, int Count)? best = null;
        foreach (var sheet in _wb.SheetNames)
        {
            if (ShouldHideSheet(sheet)) continue;
            SheetSnapshot sn;
            try { sn = _wb.LoadSheet(sheet); } catch { continue; }
            foreach (var col in sn.Columns.Where(c => c.IsNumericScalar))
            {
                if (IsOrdinalColumn(col) || ShouldHideColumn(col)) continue;
                int cnt = sn.GetNumericColumn(col.ColumnIndex).Count;
                if (best == null || cnt > best.Value.Count) best = (sheet, col, cnt);
            }
        }
        return best == null ? null : (best.Value.Sheet, best.Value.Col);
    }

    private void BuildSheetStrip()
    {
        _sheetStrip.SuspendLayout();
        _sheetStrip.Controls.Clear();
        _sheetButtons.Clear();
        if (_wb == null) { _sheetStrip.ResumeLayout(); return; }
        foreach (var name in _wb.SheetNames)
        {
            if (ShouldHideSheet(name)) continue;
            var b = new Button
            {
                Text = name,
                Tag = name,
                AutoSize = true,
                FlatStyle = FlatStyle.Flat,
                Height = 26,
                Margin = new Padding(1, 2, 1, 2),
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(214, 220, 228),
                ForeColor = Color.FromArgb(50, 56, 64)
            };
            b.FlatAppearance.BorderSize = 0;
            b.Click += (s, e) => ActivateSheet(((Button)s!).Tag!.ToString()!);
            _sheetStrip.Controls.Add(b);
            _sheetButtons.Add(b);
        }
        _sheetStrip.ResumeLayout();
        HighlightActiveSheet();
    }

    private void HighlightActiveSheet()
    {
        foreach (var b in _sheetButtons)
        {
            bool active = (string)b.Tag! == _activeSheet;
            b.BackColor = active ? Color.FromArgb(49, 110, 244) : Color.FromArgb(214, 220, 228);
            b.ForeColor = active ? Color.White : Color.FromArgb(50, 56, 64);
            b.FlatAppearance.BorderSize = active ? 0 : 0;
        }
    }

    private void ActivateSheet(string name)
    {
        if (_wb == null || name == "" ) return;
        if (_activeSheet != name && HasUnsavedChanges())
        {
            var r = MessageBox.Show(this, "切换工作表前是否保存当前改动？", "确认",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (r == DialogResult.Cancel) return;
            if (r == DialogResult.Yes) CommitPending();
            else DiscardUnsaved();
        }
        _activeSheet = name;
        _autoFocusCol = name == _autoFocusSheet ? _autoFocusCol : -1;
        HighlightActiveSheet();
        LoadSheetAndColumns();
    }

    private bool HasUnsavedChanges() => _dirtyCells.Count > 0;

    private void DiscardUnsaved()
    {
        _dirtyCells.Clear();
        _pendingUndoRows.Clear();
        _editing.Clear();
        _committed.Clear();
        _undo.Clear();
        _redo.Clear();
        UpdateTitle();
    }

    private void LoadSheetAndColumns()
    {
        if (_wb == null || string.IsNullOrEmpty(_activeSheet)) return;
        try { _snapshot = _wb.LoadSheet(_activeSheet); }
        catch (Exception ex) { _statusLabel.Text = "读取工作表失败：" + ex.Message; return; }

        var nums = _snapshot.Columns.Where(c => c.IsNumericScalar && !ShouldHideColumn(c)).ToList();
        var curveOptions = SubCurveHelper.BuildOptions(_snapshot)
            .Where(o => !ShouldHideColumn(o.Column))
            .ToList();
        _colsChecked.Items.Clear();
        foreach (var o in curveOptions) _colsChecked.Items.Add(o);
        _autoFocusCol = -1;

        _xCombo.Items.Clear();
        _xCombo.Items.Add("行号");
        foreach (var o in curveOptions) _xCombo.Items.Add(o);
        _xCombo.SelectedIndex = 0;

        RebuildFromSelection();
        _statusLabel.Text = $"已加载 [{_activeSheet}]：{_snapshot.DataRowCount} 行 × {_snapshot.ColumnCount} 列，曲线子列 {curveOptions.Count} 条";
    }

    private void RebuildFromSelection()
    {
        if (_wb == null || _snapshot == null) return;
        _suppressRebuild = true;
        try
        {
            _checkedCols = _colsChecked.CheckedItems.Cast<CurveColumnOption>().ToList();
            if (_checkedCols.Count == 0)
            {
                _series.Clear();
                _curve.SetSeries(new List<CurveSeriesView>(), -1);
                _activeYColumn = null;
                BuildGrid();
                UpdateStats();
                return;
            }
            _xColumn = _xCombo.SelectedIndex > 0 ? _xCombo.SelectedItem as CurveColumnOption : null;
            _activeYColumn = _checkedCols[0];

            _series.Clear();
            for (int i = 0; i < _checkedCols.Count; i++)
                _series.Add(BuildSeriesForOption(_checkedCols[i], i == 0));

            _curve.SetSeries(_series, 0);
            _curve.XAxisLabel = _xColumn != null ? _xColumn.DisplayName : "行号";
            _curve.YAxisLabel = _activeYColumn.DisplayName;
            SyncActiveColumnHighlight(0);

            _committed.Clear(); _editing.Clear();
            _pendingUndoRows.Clear();
            _undo.Clear(); _redo.Clear();
            foreach (var p in _curve.Points)
            {
                _committed[p.RowNumber] = (p.X, p.Y);
                _editing[p.RowNumber] = (p.X, p.Y);
            }
            BuildGrid();
            UpdateStats();
            UpdateTitle();
            _curve.ClearSelection();
        }
        finally { _suppressRebuild = false; }
    }

    private CurveSeriesView BuildSeriesForOption(CurveColumnOption opt, bool editable)
    {
        var points = new List<CurvePoint>();
        var snap = _snapshot!;
        for (int gi = 0; gi < snap.RowNumbers.Count; gi++)
        {
            int row = snap.RowNumbers[gi];
            if (!SubCurveHelper.TryReadValue(snap, opt, gi, out var val)) continue;
            double x = row; bool xEditable = false;
            if (_xColumn != null)
            {
                if (SubCurveHelper.TryReadValue(snap, _xColumn, gi, out var xv))
                { x = xv; xEditable = true; }
                else continue;
            }
            points.Add(new CurvePoint(x, val, row, xEditable));
        }
        var view = new CurveSeriesView
        {
            Name = opt.DisplayName,
            Color = CurveEditor.Palette[_series.Count % CurveEditor.Palette.Length],
            IsEditable = editable
        };
        view.Points.AddRange(points);
        return view;
    }

    private void OnSelectionChanged() => OnSelectionChangedSafe();
    private void OnSelectionChangedSafe()
    {
        if (_suppressRebuild || _wb == null) return;
        CommitPending();
        RebuildFromSelection();
    }

    private void SetActiveSeries(int seriesIndex)
    {
        if (seriesIndex < 0 || seriesIndex >= _checkedCols.Count) return;
        CommitPending();
        _activeYColumn = _checkedCols[seriesIndex];
        _curve.SetActiveSeries(seriesIndex);
        _curve.YAxisLabel = _activeYColumn.DisplayName;
        _committed.Clear(); _editing.Clear();
        _pendingUndoRows.Clear();
        foreach (var p in _curve.Points)
        {
            _committed[p.RowNumber] = (p.X, p.Y);
            _editing[p.RowNumber] = (p.X, p.Y);
        }
        HighlightEditableColumn();
        SyncActiveColumnHighlight(seriesIndex);
        UpdateStats();
    }

    /// <summary>把当前编辑列在“曲线列 (Y) 多选”里高亮选中，但不改变勾选状态。</summary>
    private void SyncActiveColumnHighlight(int seriesIndex)
    {
        if (seriesIndex < 0 || seriesIndex >= _checkedCols.Count) return;
        var meta = _checkedCols[seriesIndex];
        int idx = _colsChecked.Items.IndexOf(meta);
        if (idx >= 0) _colsChecked.SetSelected(idx, true);
    }

    private void HighlightEditableColumn()
    {
        int activeCol = _activeYColumn?.Column.ColumnIndex ?? -1;
        for (int i = 0; i < _grid.Columns.Count; i++)
        {
            var col = _grid.Columns[i];
            if (i == 0) continue; // 行号列
            col.DefaultCellStyle.BackColor = (i - 1 == activeCol) ? Color.FromArgb(255, 250, 235) : Color.White;
        }
        _grid.Refresh();
    }

    // ---------- 同步 ----------
    private void OnPointsChanged(IReadOnlyList<int> rows)
    {
        if (rows.Count == 0) return;
        foreach (var row in rows)
        {
            if (!TryGetPoint(row, out var p)) continue;
            _editing[row] = (p.X, p.Y);
            if (_activeYColumn != null)
            {
                int col = _activeYColumn.Column.ColumnIndex;
                _dirtyCells.Add((col, row));
                if (_activeYColumn.IsSubCurve)
                {
                    string oldRaw = CellText(col, row);
                    UpdateSnapshotCell(row, col, SubCurveHelper.SetValue(oldRaw, _activeYColumn, p.Y));
                }
                else
                    UpdateSnapshotCell(row, col, FormatCellValue(p.Y, _activeYColumn.IsInteger));
            }
            if (_xColumn != null && p.XEditable)
            {
                int xCol = _xColumn.Column.ColumnIndex;
                _dirtyCells.Add((xCol, row));
                UpdateSnapshotCell(row, xCol, _xColumn.IsSubCurve
                    ? SubCurveHelper.SetValue(CellText(xCol, row), _xColumn, p.X)
                    : FormatCellValue(p.X, _xColumn.IsInteger));
            }
            _pendingUndoRows.Add(row);
        }
        UpdateGridCells(rows);
        UpdateStats();
        UpdateTitle();
    }

    private void OnEditCommitted()
    {
        RecordDragUndo(_pendingUndoRows.ToList());
        _pendingUndoRows.Clear();
        if (_autoSaveCheck.Checked) { _autoSaveTimer.Stop(); _autoSaveTimer.Start(); }
    }

    private void RecordDragUndo(IReadOnlyList<int> rows)
    {
        if (rows.Count == 0) return;
        var items = new List<(int, int, string, string)>();
        foreach (var row in rows)
        {
            if (!_editing.TryGetValue(row, out var cur)) continue;
            var old = _committed.TryGetValue(row, out var o) ? o : cur;
            if (Math.Abs(old.X - cur.X) < 1e-12 && Math.Abs(old.Y - cur.Y) < 1e-12) continue;
            if (_activeYColumn != null && !_activeYColumn.IsSubCurve)
                items.Add((_activeYColumn.Column.ColumnIndex, row, FormatCellValue(old.Y, _activeYColumn.IsInteger), FormatCellValue(cur.Y, _activeYColumn.IsInteger)));
            if (_xColumn != null && !_xColumn.IsSubCurve && TryGetPoint(row, out var p) && p.XEditable && Math.Abs(old.X - cur.X) > 1e-12)
                items.Add((_xColumn.Column.ColumnIndex, row, FormatCellValue(old.X, _xColumn.IsInteger), FormatCellValue(cur.X, _xColumn.IsInteger)));
            _committed[row] = cur;
        }
        if (items.Count > 0) { _undo.Add(new EditCmd(items)); _redo.Clear(); }
    }

    private bool TryGetPoint(int row, out CurvePoint p)
    {
        foreach (var pp in _curve.Points)
            if (pp.RowNumber == row) { p = pp; return true; }
        p = null!;
        return false;
    }

    private void CommitPending()
    {
        if (_wb == null || _dirtyCells.Count == 0) return;
        _autoSaveTimer.Stop();
        var nums = new List<CellWrite>();
        var strs = new List<CellWriteString>();
        foreach (var (col, row) in _dirtyCells)
        {
            if (col < 0 || col >= _snapshot!.Columns.Count) continue;
            var meta = _snapshot.Columns[col];
            string text = CellText(col, row);
            string cellRef = CellHelper.ToCellReference(col, row);
            if (meta.IsNumericScalar && CellHelper.TryParseDouble(text, out var v))
                nums.Add(new CellWrite(_snapshot.SheetName, cellRef, v, meta.IsInteger, 6));
            else
                strs.Add(new CellWriteString(_snapshot.SheetName, cellRef, text));
        }
        _selfWrite = true;
        bool ok = _wb.TryWriteCells(nums, out var errNum);
        bool okStr = _wb.TryWriteCellsString(strs, out var errStr);
        _selfWrite = false;
        if (!ok) _statusLabel.Text = "⚠ " + errNum;
        else if (!okStr) _statusLabel.Text = "⚠ " + errStr;
        else
        {
            _statusLabel.Text = $"已保存 {nums.Count + strs.Count} 个单元格";
            _dirtyCells.Clear();
        }
        UpdateTitle();
    }

    private string CellText(int col, int row)
    {
        if (_snapshot != null && _rowToGridIndex.TryGetValue(row, out var gi) &&
            gi >= 0 && gi < _snapshot.Grid.Count && col >= 0 && col < _snapshot.Grid[gi].Length)
            return _snapshot.Grid[gi][col] ?? "";
        return "";
    }

    private double SnapshotValue(int row, int colIndex, double fallback)
    {
        if (_snapshot != null && _rowToGridIndex.TryGetValue(row, out var gi) &&
            gi >= 0 && gi < _snapshot.Grid.Count && colIndex >= 0 && colIndex < _snapshot.Grid[gi].Length)
            return CellHelper.TryParseDouble(_snapshot.Grid[gi][colIndex], out var v) ? v : fallback;
        return fallback;
    }

    private void UpdateSnapshotCell(int rowNumber, int colIndex, string value)
    {
        if (_snapshot == null) return;
        if (_rowToGridIndex.TryGetValue(rowNumber, out var gi) &&
            gi >= 0 && gi < _snapshot.Grid.Count && colIndex >= 0 && colIndex < _snapshot.Grid[gi].Length)
            _snapshot.Grid[gi][colIndex] = value;
    }

    private static string FormatCellValue(double v, bool integer)
    {
        if (integer) return ((long)Math.Round(v)).ToString(CultureInfo.InvariantCulture);
        return v.ToString("0.######", CultureInfo.InvariantCulture);
    }

    // ---------- 网格 ----------
    private void BuildGrid()
    {
        _grid.Columns.Clear();
        if (_snapshot == null) { _grid.Rows.Clear(); return; }
        var header = new DataGridViewTextBoxColumn
        {
            HeaderText = "行号",
            Name = "行号",
            ReadOnly = true,
            Resizable = DataGridViewTriState.True,
            Width = 70,
            DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
        };
        _grid.Columns.Add(header);
        foreach (var col in _snapshot.Columns)
        {
            var c = new DataGridViewTextBoxColumn
            {
                HeaderText = col.DisplayName,
                Name = "col" + col.ColumnIndex,
                ReadOnly = false,
                Resizable = DataGridViewTriState.True,
                MinimumWidth = 60,
                Width = 130,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            int ci = col.ColumnIndex;
            var align = ci >= 0 && ci < _snapshot.ColumnAlignments.Count
                ? _snapshot.ColumnAlignments[ci]
                : HorizontalAlign.Default;
            // 默认对齐：数值列按 Excel 惯例右对齐，文本列左对齐
            if (align == HorizontalAlign.Default)
                align = col.IsNumericScalar ? HorizontalAlign.Right : HorizontalAlign.Left;
            c.DefaultCellStyle.Alignment = MapAlign(align);
            c.Visible = !ShouldHideColumn(col);
            if (_activeYColumn != null && col.ColumnIndex == _activeYColumn.Column.ColumnIndex)
                c.DefaultCellStyle.BackColor = Color.FromArgb(255, 250, 235);
            _grid.Columns.Add(c);
        }

        _grid.Rows.Clear();
        _rowToGridIndex.Clear();
        for (int gi = 0; gi < _snapshot.RowNumbers.Count; gi++)
        {
            var cells = new object[_snapshot.ColumnCount + 1];
            int rowNum = _snapshot.RowNumbers[gi];
            cells[0] = rowNum;
            _rowToGridIndex[rowNum] = gi;
            var row = _snapshot.Grid[gi];
            for (int c = 0; c < _snapshot.ColumnCount; c++)
                cells[c + 1] = c < row.Length ? (row[c] ?? "") : "";
            _grid.Rows.Add(cells);
        }
        _grid.ClearSelection();
    }

    private void UpdateGridCells(IReadOnlyList<int> rows)
    {
        if (_snapshot == null || _activeYColumn == null) return;
        int yCol = _activeYColumn.Column.ColumnIndex;
        int yGridCol = yCol + 1;
        foreach (var row in rows)
        {
            if (!_rowToGridIndex.TryGetValue(row, out var gi) || gi >= _snapshot.Grid.Count) continue;
            if (_editing.TryGetValue(row, out var v))
            {
                _grid.Rows[gi].Cells[yGridCol].Value = _activeYColumn.IsSubCurve
                    ? CellText(yCol, row)
                    : FormatCellValue(v.Y, _activeYColumn.IsInteger);
                if (_xColumn != null)
                {
                    int xCol = _xColumn.Column.ColumnIndex;
                    int xGridCol = xCol + 1;
                    _grid.Rows[gi].Cells[xGridCol].Value = _xColumn.IsSubCurve
                        ? CellText(xCol, row)
                        : FormatCellValue(v.X, _xColumn.IsInteger);
                }
            }
        }
    }

    private void OnGridCellEdited(object? sender, DataGridViewCellEventArgs e)
    {
        if (_snapshot == null || e.RowIndex < 0 || e.ColumnIndex < 0) return;
        int col = e.ColumnIndex - 1;
        if (col < 0 || col >= _snapshot.ColumnCount) return;
        int row = _snapshot.RowNumbers[e.RowIndex];
        var cell = _grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
        string text = cell.Value?.ToString() ?? "";
        string oldText = CellText(col, row);
        if (text == oldText) return;

        // 更新曲线（若该列被绘制）
        ApplyPlottedCellChange(col, row, text);

        // 当前编辑列（可拖动那条）的编辑基线同步
        if (_activeYColumn != null && col == _activeYColumn.Column.ColumnIndex && !_activeYColumn.IsSubCurve)
        {
            if (CellHelper.TryParseDouble(text, out var y))
            {
                double x = _editing.TryGetValue(row, out var ce) ? ce.X : row;
                y = _activeYColumn.IsInteger ? Math.Round(y) : y;
                _editing[row] = (x, y);
                _committed[row] = (x, y);
            }
            else { _editing.Remove(row); _committed.Remove(row); }
        }

        UpdateSnapshotCell(row, col, text);
        _dirtyCells.Add((col, row));
        _undo.Add(new EditCmd(new List<(int, int, string, string)> { (col, row, oldText, text) }));
        _redo.Clear();
        UpdateStats();
        UpdateTitle();
        if (_autoSaveCheck.Checked) CommitPending();
    }

    private void ApplyPlottedCellChange(int col, int row, string text)
    {
        int si = _checkedCols.FindIndex(c => c.Column.ColumnIndex == col);
        if (si < 0) return;
        // 子曲线列整格为数组/JSON，直接编辑按非标量处理，避免误删曲线点
        if (_checkedCols[si].IsSubCurve && !CellHelper.TryParseDouble(text, out _))
            return;
        if (CellHelper.TryParseDouble(text, out var y))
        {
            double x = row; bool xEditable = false;
            if (_xColumn != null &&
                _rowToGridIndex.TryGetValue(row, out var xgi) &&
                SubCurveHelper.TryReadValue(_snapshot!, _xColumn, xgi, out var xv))
            {
                x = xv;
                xEditable = true;
            }
            _curve.SetSeriesPoint(si, row, x, y, xEditable);
        }
        else
        {
            _curve.RemoveSeriesPoint(si, row);
        }
        _curve.Invalidate();
    }

    private void OnGridSelectionChanged(object? sender, EventArgs e)
    {
        // 由曲线联动表格或处于重建期间时不回写曲线
        if (_updatingGridSelection || _syncFromGrid || _suppressRebuild) return;
        if (_snapshot == null) return;

        var cells = _grid.SelectedCells.Cast<DataGridViewCell>()
            .Where(c => c.RowIndex >= 0 && c.ColumnIndex > 0)
            .ToList();
        if (cells.Count == 0) return;

        // 取选中单元格所属第一列（去掉“行号”列）
        int tableCol = cells[0].ColumnIndex - 1;

        // 在“曲线列 (Y) 多选”完整列表里找到该列，高亮并滚动到可见区（不改变勾选）
        var allItems = _colsChecked.Items.Cast<CurveColumnOption>().ToList();
        int listIdx = allItems.FindIndex(m => m.Column.ColumnIndex == tableCol);
        if (listIdx < 0) return;
        if (_colsChecked.SelectedIndex != listIdx)
            _colsChecked.SetSelected(listIdx, true);
        _colsChecked.TopIndex = Math.Max(0, listIdx - 2);

        if (_checkedCols.Count == 0) return;
        // 若该列已是勾选的曲线列，则设为当前编辑列并同步选中点
        int si = _checkedCols.FindIndex(cc => cc.Column.ColumnIndex == tableCol);
        if (si < 0) return;
        var rows = cells
            .Where(c => c.ColumnIndex - 1 == tableCol)
            .Select(c => _snapshot.RowNumbers[c.RowIndex])
            .Distinct()
            .ToList();
        if (rows.Count == 0) return;

        _syncFromGrid = true;
        if (si != _curve.ActiveSeriesIndex) SetActiveSeries(si);
        _curve.SelectPointsByRows(rows);
        _syncFromGrid = false;
    }

    // ---------- 统计 ----------
    private void UpdateStats()
    {
        _statLabel.Text = Statistics.Summarize(_curve.Points.Select(p => p.Y)).AsText();
    }

    // ---------- 批量 ----------
    private void OnApplyValue() { double v = (double)_valUpDown.Value; _curve.ApplyToSelected(p => (p.X, v)); }
    private void OpenContextAtChart()
    {
        BuildContextMenu();
        _menu.Show(_curve, new Point(_curve.Width / 2, _curve.Height / 2));
    }
    private void BatchOffset(double d) => _curve.ApplyToSelected(p => (p.X, p.Y + d));
    private void BatchScale(double k) => _curve.ApplyToSelected(p => (p.X, p.Y * k));
    private void BatchClamp(double lo, double hi) => _curve.ApplyToSelected(p => (p.X, CurveMath.Clamp(p.Y, lo, hi)));
    private void BatchSmooth()
    {
        var ys = _curve.Points.Select(p => p.Y).ToList();
        if (ys.Count == 0) return;
        var sm = CurveMath.MovingAverage(ys, 3);
        int i = 0;
        _curve.ApplyToAll(_ => (_.X, sm[i++]));
    }
    private void BatchNormalize()
    {
        var ys = _curve.Points.Select(p => p.Y).ToList();
        if (ys.Count < 2) return;
        double min = ys.Min(), max = ys.Max();
        if (Math.Abs(max - min) < 1e-12) return;
        _curve.ApplyToAll(p => (p.X, (p.Y - min) / (max - min)));
    }
    private void BatchRandom(double amp)
    {
        var rand = new Random();
        _curve.ApplyToSelected(p => (p.X, p.Y + (rand.NextDouble() * 2 - 1) * amp));
    }

    // ---------- 撤销/重做 ----------
    private void Undo()
    {
        if (_undo.Count == 0) { _statusLabel.Text = "没有可撤销的操作"; return; }
        var cmd = _undo[^1]; _undo.RemoveAt(_undo.Count - 1); _redo.Add(cmd);
        ApplyCmd(cmd, true);
    }
    private void Redo()
    {
        if (_redo.Count == 0) { _statusLabel.Text = "没有可重做的操作"; return; }
        var cmd = _redo[^1]; _redo.RemoveAt(_redo.Count - 1); _undo.Add(cmd);
        ApplyCmd(cmd, false);
    }

    private void ApplyCmd(EditCmd cmd, bool toOld)
    {
        if (_wb == null || _snapshot == null) return;
        foreach (var (col, row, o, n) in cmd.Cells)
        {
            string val = toOld ? o : n;
            UpdateSnapshotCell(row, col, val);
            ApplyPlottedCellChange(col, row, val);
            if (_activeYColumn != null && col == _activeYColumn.Column.ColumnIndex)
            {
                if (CellHelper.TryParseDouble(val, out var y))
                {
                    double x = _editing.TryGetValue(row, out var ce) ? ce.X : row;
                    y = _activeYColumn.IsInteger ? Math.Round(y) : y;
                    _editing[row] = (x, y);
                    _committed[row] = (x, y);
                }
                else { _editing.Remove(row); _committed.Remove(row); }
            }
            _dirtyCells.Add((col, row));
            if (_rowToGridIndex.TryGetValue(row, out var gi) && gi >= 0 && gi < _grid.Rows.Count)
                _grid.Rows[gi].Cells[col + 1].Value = val;
        }
        _curve.Invalidate();
        UpdateStats();
        UpdateTitle();
        CommitPending();
    }

    // ---------- 保存/刷新/导出 ----------
    private void OnSave() => CommitPending();

    private void OnSaveAs()
    {
        if (_wb == null) return;
        CommitPending();
        using var sfd = new SaveFileDialog
        {
            Filter = "Excel 工作簿 (*.xlsm)|*.xlsm|Excel 工作簿 (*.xlsx)|*.xlsx",
            FileName = Path.GetFileName(_wb.Path)
        };
        if (sfd.ShowDialog(this) == DialogResult.OK)
        {
            if (_wb.TrySaveAs(sfd.FileName, out var err)) _statusLabel.Text = "已另存为 " + sfd.FileName;
            else MessageBox.Show(this, err, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void OnExport()
    {
        using var sfd = new SaveFileDialog { Filter = "PNG 图片 (*.png)|*.png", FileName = "curve.png" };
        if (sfd.ShowDialog(this) == DialogResult.OK) _curve.SaveBitmap(sfd.FileName);
    }

    private void OnReload()
    {
        if (_wb == null) return;
        CommitPending();
        ReloadFromDisk();
    }

    private void ReloadFromDisk()
    {
        if (_wb == null || string.IsNullOrEmpty(_activeSheet)) return;
        try
        {
            _wb.RefreshMeta();
            _snapshot = _wb.LoadSheet(_activeSheet);
            RebuildFromSelection();
            _statusLabel.Text = "已从磁盘重新加载";
        }
        catch (Exception ex)
        {
            _statusLabel.Text = "重新加载失败（文件可能被占用）：" + ex.Message;
        }
    }

    // ---------- 布局切换 ----------
    private void SetGridSide(bool atRight)
    {
        _gridAtRight = atRight;
        foreach (ToolStripMenuItem item in _layoutButton.DropDownItems) item.Checked = false;
        ((ToolStripMenuItem)_layoutButton.DropDownItems[atRight ? 1 : 0]).Checked = true;
        _chartGridSplit.Orientation = atRight ? Orientation.Vertical : Orientation.Horizontal;
        try
        {
            if (atRight) _chartGridSplit.SplitterDistance = Math.Max(200, (int)(_chartGridSplit.Width * 0.72));
            else _chartGridSplit.SplitterDistance = Math.Max(140, (int)(_chartGridSplit.Height * 0.58));
        }
        catch { }
        SaveSettings();
    }

    // ---------- 右键菜单 ----------
    private int _menuSeries = -1;
    private void BuildContextMenu()
    {
        _menu.Items.Clear();
        var loc = _curve.PointToClient(Cursor.Position);
        _menuSeries = _curve.HitTestAnySeries(loc);

        if (_menuSeries >= 0)
        {
            string name = _curve.GetSeriesName(_menuSeries);
            var setActive = new ToolStripMenuItem("设为当前编辑列：" + name);
            setActive.Click += (s, e) => SetActiveSeries(_menuSeries);
            _menu.Items.Add(setActive);

            var hide = new ToolStripMenuItem("隐藏该曲线");
            hide.Click += (s, e) => _curve.SetSeriesVisible(_menuSeries, false);
            _menu.Items.Add(hide);

            var only = new ToolStripMenuItem("仅显示该曲线");
            only.Click += (s, e) => { for (int i = 0; i < _curve.SeriesCount; i++) _curve.SetSeriesVisible(i, i == _menuSeries); };
            _menu.Items.Add(only);

            var all = new ToolStripMenuItem("显示全部曲线");
            all.Click += (s, e) => { for (int i = 0; i < _curve.SeriesCount; i++) _curve.SetSeriesVisible(i, true); };
            _menu.Items.Add(all);
        }
        else
        {
            var fit = new ToolStripMenuItem("自动适配视图");
            fit.Click += (s, e) => _curve.AutoFitView();
            _menu.Items.Add(fit);
        }

        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add(BuildBatchMenu());
        var stats = new ToolStripMenuItem("查看统计");
        stats.Click += (s, e) => MessageBox.Show(this, _statLabel.Text, "当前编辑列统计", MessageBoxButtons.OK, MessageBoxIcon.Information);
        _menu.Items.Add(stats);
        _menu.Items.Add(new ToolStripSeparator());

        if (_menuSeries >= 0)
        {
            var fit = new ToolStripMenuItem("自动适配视图");
            fit.Click += (s, e) => _curve.AutoFitView();
            _menu.Items.Add(fit);
        }
        else
        {
            AddCheck(_menu, "平滑曲线", () => _curve.ShowSpline, v => _curve.ShowSpline = v);
            AddCheck(_menu, "网格", () => _curve.ShowGrid, v => _curve.ShowGrid = v);
            AddCheck(_menu, "数据点", () => _curve.ShowPoints, v => _curve.ShowPoints = v);
            AddCheck(_menu, "坐标标签", () => _curve.ShowLabels, v => _curve.ShowLabels = v);
            _menu.Items.Add(new ToolStripSeparator());
            var exp = new ToolStripMenuItem("导出 PNG");
            exp.Click += (s, e) => OnExport();
            _menu.Items.Add(exp);
            var reload = new ToolStripMenuItem("刷新重载");
            reload.Click += (s, e) => OnReload();
            _menu.Items.Add(reload);
            var open = new ToolStripMenuItem("打开工作簿...");
            open.Click += (s, e) => OnOpen();
            _menu.Items.Add(open);
        }
    }

    private void EnsureActiveHitSeries()
    {
        if (_menuSeries >= 0) SetActiveSeries(_menuSeries);
    }

    private ToolStripMenuItem BuildBatchMenu()
    {
        var menu = new ToolStripMenuItem("批量操作");
        AddBatch(menu, "设为值...", () =>
        {
            var v = PromptDouble("设为值", 0);
            if (v.HasValue) { EnsureActiveHitSeries(); _curve.ApplyToSelected(p => (p.X, v.Value)); }
        });
        AddBatch(menu, "偏移...", () =>
        {
            var d = PromptDouble("偏移量（加减值）", 0);
            if (d.HasValue) { EnsureActiveHitSeries(); BatchOffset(d.Value); }
        });
        AddBatch(menu, "缩放...", () =>
        {
            var k = PromptDouble("缩放倍数", 1);
            if (k.HasValue) { EnsureActiveHitSeries(); BatchScale(k.Value); }
        });
        AddBatch(menu, "钳制 最小,最大...", () =>
        {
            var lo = PromptDouble("最小值", 0);
            if (lo.HasValue)
            {
                var hi = PromptDouble("最大值", 100);
                if (hi.HasValue) { EnsureActiveHitSeries(); BatchClamp(lo.Value, hi.Value); }
            }
        });
        AddBatch(menu, "整列平滑", () => { EnsureActiveHitSeries(); BatchSmooth(); });
        AddBatch(menu, "整列归一化", () => { EnsureActiveHitSeries(); BatchNormalize(); });
        AddBatch(menu, "随机扰动...", () =>
        {
            var a = PromptDouble("扰动幅度（±）", 5);
            if (a.HasValue) { EnsureActiveHitSeries(); BatchRandom(a.Value); }
        });
        return menu;
    }

    private static void AddBatch(ToolStripMenuItem parent, string text, Action action)
    {
        var item = new ToolStripMenuItem(text);
        item.Click += (s, e) => action();
        parent.DropDownItems.Add(item);
    }

    private double? PromptDouble(string title, double defaultValue)
    {
        using var f = new Form
        {
            Text = title,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterParent,
            Size = new Size(280, 130),
            MaximizeBox = false,
            MinimizeBox = false
        };
        f.Controls.Add(new Label { Text = "请输入数值：", AutoSize = true, Location = new Point(16, 18) });
        var num = new NumericUpDown
        {
            DecimalPlaces = 6,
            Minimum = -1.0E12m,
            Maximum = 1.0E12m,
            Value = Math.Clamp((decimal)defaultValue, -1.0E12m, 1.0E12m),
            Location = new Point(130, 15),
            Width = 120
        };
        f.Controls.Add(num);
        var ok = new Button { Text = "确定", DialogResult = DialogResult.OK, Location = new Point(78, 66), Width = 80 };
        var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Location = new Point(168, 66), Width = 80 };
        f.Controls.Add(ok);
        f.Controls.Add(cancel);
        f.AcceptButton = ok;
        f.CancelButton = cancel;
        return f.ShowDialog(this) == DialogResult.OK ? (double?)num.Value : null;
    }

    private static void AddCheck(ContextMenuStrip menu, string text, Func<bool> get, Action<bool> set)
    {
        var item = new ToolStripMenuItem(text) { Checked = get() };
        item.Click += (s, e) => { bool nv = !item.Checked; item.Checked = nv; set(nv); };
        menu.Items.Add(item);
    }

    // ---------- 监视 ----------
    private void SetupWatcher(string path)
    {
        _watcher?.Dispose();
        var dir = Path.GetDirectoryName(path)!;
        _watcher = new FileSystemWatcher(dir, Path.GetFileName(path))
        {
            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName,
            EnableRaisingEvents = true
        };
        _watcher.Changed += OnFileChanged;
    }
    private void OnFileChanged(object sender, FileSystemEventArgs e)
    {
        if (_selfWrite) return;
        _reloadTimer.Stop();
        _reloadTimer.Start();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (_wb != null && _dirtyCells.Count > 0)
        {
            var r = MessageBox.Show(this, "有未保存的改动，是否保存？", "确认", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (r == DialogResult.Cancel) { e.Cancel = true; return; }
            if (r == DialogResult.Yes) CommitPending();
        }
        _watcher?.Dispose();
        _autoSaveTimer.Stop();
        _reloadTimer.Stop();
        SaveSettings();
        _wb?.Dispose();
        base.OnFormClosing(e);
    }

    private void UpdateTitle()
    {
        var dirty = _dirtyCells.Count > 0 ? "（有未保存改动）" : "";
        Text = "GameCurve - 游戏数据曲线编辑器" + (_wb != null ? " - " + Path.GetFileName(_wb.Path) : "") + dirty;
    }

    // ---------- UI 辅助 ----------
    private static Label Section(string t) => new()
    {
        Text = t,
        Font = new Font("Microsoft YaHei UI", 9.5f, FontStyle.Bold),
        ForeColor = Color.FromArgb(49, 110, 244),
        AutoSize = false,
        Height = 20
    };
    private static Label MakeLabel(string t) => new() { Text = t, AutoSize = true };
    private static Label MakeHint(string t) => new()
    {
        Text = t,
        AutoSize = false,
        Height = 36,
        Font = new Font("Microsoft YaHei UI", 8f),
        ForeColor = Color.FromArgb(100, 106, 114)
    };
    private static Button MakeButton(string text, Action onClick)
    {
        var b = new Button { Text = text, Height = 30 };
        b.Click += (s, e) => onClick();
        return b;
    }

    private static DataGridViewContentAlignment MapAlign(HorizontalAlign a) => a switch
    {
        HorizontalAlign.Center => DataGridViewContentAlignment.MiddleCenter,
        HorizontalAlign.Right => DataGridViewContentAlignment.MiddleRight,
        _ => DataGridViewContentAlignment.MiddleLeft
    };

    private static ToolStripMenuItem MakeMenu(string text, string tooltip, Action onClick, bool check)
    {
        var m = new ToolStripMenuItem(text) { Checked = check, CheckOnClick = false };
        m.Click += (s, e) => onClick();
        return m;
    }

    private void RebuildOpenMenu()
    {
        _openMenu.DropDownItems.Clear();
        var pick = new ToolStripMenuItem("选择文件...");
        pick.Click += (s, e) => OnOpen();
        _openMenu.DropDownItems.Add(pick);
        _openMenu.DropDownItems.Add(new ToolStripSeparator());

        string? folder = Path.GetDirectoryName(_wb?.Path ?? _settings.LastFile);
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            _openMenu.DropDownItems.Add(new ToolStripMenuItem("（暂无历史文件夹）") { Enabled = false });
            return;
        }

        var files = Directory.EnumerateFiles(folder, "*", SearchOption.TopDirectoryOnly)
            .Where(IsExcelFile)
            .OrderByDescending(f => File.GetLastWriteTimeUtc(f))
            .ToList();
        if (files.Count == 0)
        {
            _openMenu.DropDownItems.Add(new ToolStripMenuItem("（该文件夹暂无 Excel 文件）") { Enabled = false });
            return;
        }

        foreach (var file in files)
        {
            var item = new ToolStripMenuItem(Path.GetFileName(file)) { Tag = file };
            item.Click += (s, e) => OpenFile((string)((ToolStripMenuItem)s!).Tag!);
            _openMenu.DropDownItems.Add(item);
        }
    }

    private static bool IsExcelFile(string path)
        => path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
           path.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase);

    private void AddButton(string text, string tooltip, Action onClick)
    {
        var b = new ToolStripButton(text);
        b.Click += (s, e) => onClick();
        _tool.Items.Add(b);
    }

    private void BuildTooltips()
    {
        // 仅保留右侧功能面板控件的提示
        Tip(_valUpDown, "输入数值后点“应用到选中”把选中点都设为该值");
        Tip(_stepUpDown, "方向键微调步长；按住 Shift 为 ×10，按住 Ctrl 为 ×0.1");
        Tip(_offsetUpDown, "对选中点整体增加或减少的数值（配合“偏移 Δ”按钮）");
        Tip(_scaleUpDown, "对选中点整体放大/缩小的倍数（配合“缩放 ×”按钮）");
        Tip(_clampMin, "钳制最小值");
        Tip(_clampMax, "钳制最大值");
        Tip(_randUpDown, "随机扰动幅度（±，配合“随机扰动”按钮）");
        Tip(_statLabel, "当前编辑列的统计信息（最大/平均/总和/标准差等）");
    }

    private void Tip(Control c, string t) => _tip.SetToolTip(c, t);

    private void OnSelectionChangedUi()
    {
        EnsureActiveSeriesSynced();
        _selInfo.Text = "选中: " + _curve.SelectedCount;
        if (_curve.SelectedCount > 0)
        {
            var first = _curve.Points.Where(p => _curve.SelectedRows.Contains(p.RowNumber)).FirstOrDefault();
            if (first != null)
            {
                _valUpDown.Value = Math.Clamp((decimal)first.Y, _valUpDown.Minimum, _valUpDown.Maximum);
                // 来自表格点击时不再滚动/定位（避免跳动）
                if (_syncFromGrid) return;
                if (_activeYColumn != null)
                {
                    int col = _activeYColumn.Column.ColumnIndex + 1;
                    if (col < _grid.Columns.Count)
                    {
                        // 只选中对应的单元格，而不是整行
                        _updatingGridSelection = true;
                        try
                        {
                            _grid.ClearSelection();
                            foreach (var row in _curve.SelectedRows)
                            {
                                if (_rowToGridIndex.TryGetValue(row, out var gi) && gi >= 0 && gi < _grid.Rows.Count)
                                    _grid.Rows[gi].Cells[col].Selected = true;
                            }
                            if (_rowToGridIndex.TryGetValue(first.RowNumber, out var firstGi) && firstGi >= 0 && firstGi < _grid.Rows.Count)
                            {
                                _grid.FirstDisplayedScrollingRowIndex = Math.Max(0, firstGi);
                                _grid.FirstDisplayedScrollingColumnIndex = Math.Max(0, col - 1);
                                _grid.CurrentCell = _grid.Rows[firstGi].Cells[col];
                            }
                        }
                        finally
                        {
                            _updatingGridSelection = false;
                        }
                    }
                }
            }
        }
    }

    /// <summary>曲线控件内部切换当前编辑列时，同步主界面的编辑列元数据。</summary>
    private void EnsureActiveSeriesSynced()
    {
        int idx = _curve.ActiveSeriesIndex;
        if (idx < 0 || idx >= _checkedCols.Count) return;
        if (_activeYColumn == _checkedCols[idx]) return;
        CommitPending();
        _activeYColumn = _checkedCols[idx];
        _curve.YAxisLabel = _activeYColumn.DisplayName;
        _committed.Clear();
        _editing.Clear();
        _pendingUndoRows.Clear();
        foreach (var p in _curve.Points)
        {
            _committed[p.RowNumber] = (p.X, p.Y);
            _editing[p.RowNumber] = (p.X, p.Y);
        }
        HighlightEditableColumn();
        SyncActiveColumnHighlight(idx);
    }
}
