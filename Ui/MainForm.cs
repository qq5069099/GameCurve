using System.Globalization;
using System.Diagnostics;
using System.Text;
using GameCurve.Excel;
using GameCurve.Models;
using GameCurve.Services;

namespace GameCurve.Ui;

public sealed class MainForm : Form
{
    private const int LeftPanelWidth = 260;   // 左侧功能面板宽度
    private const int RightPanelWidth = 260;  // 可完整容纳编辑/批量/统计控件

    private readonly CurveEditor _curve = new() { Dock = DockStyle.Fill };
    private readonly DataGridView _grid = new() { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false };
    private readonly ToolTip _tip = new() { AutoPopDelay = 12000, InitialDelay = 500, ReshowDelay = 120 };
    /// <summary>当前已在“曲线列 (Y) 多选”上显示的提示文本，避免每次移动都重设 ToolTip。</summary>
    private string? _colCheckedTipText;
    private readonly ContextMenuStrip _gridMenu = new();
    private (int Row, int Col, GridArea Area) _gridTarget = (-1, -1, GridArea.None);
    private int _rowHeaderAnchor = -1;   // 行头拖拽的起始行
    private int _rowHeaderLastRow = -1;  // 行头拖拽过程中最新经过的行
    private bool _rowHeaderDragging;     // 是否正在行头拖拽
    private List<int> _gridTargetSel = new(); // 右键发生时捕获的已选中行
    private bool _gridTargetRowBlock;           // 右键目标行是否属于“整行”选中块
    private readonly List<StructuralOp> _pendingStructure = new();
    private readonly List<(int Col, string Text)> _pendingHeaderRename = new();
    private readonly HashSet<int> _pendingHeaderAlign = new();
    private readonly HashSet<int> _pendingColumnAlign = new();
    private bool _structureQueued;
    private bool _columnOrderDirty;
    private bool _suppressColumnOrderDirty;
    private bool _columnWidthDirty;
    private bool _suppressColumnWidthDirty;
    private bool _sheetOrderDirty;
    private TextBox? _headerEdit;
    private int _headerEditPhysicalCol = -1;   // -1 表示没有在编辑列名

    private readonly ToolStrip _tool = new();
    private readonly ToolStripMenuItem _autoSaveCheck = new("自动保存") { Checked = false, CheckOnClick = true };
    private readonly ToolStripDropDownButton _openMenu = new("文件");
    private readonly ToolStripDropDownButton _editMenu = new("编辑");
    private readonly ToolStripDropDownButton _gridButton = new("表格");
    private readonly ToolStripDropDownButton _viewButton = new("视图");
    private ToolStripMenuItem _gridMaxItem = null!;
    private ToolStripMenuItem _gridBottomItem = null!;
    private ToolStripMenuItem _gridRightItem = null!;
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
    private Button? _addSheetButton;
    private Button? _dragSheetBtn;
    private Point _dragSheetStart;
    private bool _draggingSheet;
    private bool _suppressSheetClick;
    private TextBox? _sheetRenameEdit;
    private string _sheetRenameOldName = "";

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
    private readonly ComboBox _fitTypeCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly NumericUpDown _fitDegree = new() { DecimalPlaces = 0, Minimum = 2, Maximum = 8, Value = 2 };
    private readonly Label _fitInfo = new() { AutoSize = false, Height = 84, Font = new Font("Microsoft YaHei UI", 7.5f), ForeColor = Color.FromArgb(70, 76, 84) };
    private FitDialog? _fitDialog;
    private SplashForm? _splash;
    private bool _busyLoading;
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
    private readonly Dictionary<(int Col, int Row), string> _subEditOldText = new();

    private bool _gridAtRight;
    private bool _gridMaximized;
    private bool _suppressRebuild;
    private bool _syncFromGrid;
    private bool _updatingGridSelection;
    private int _preferredListIndex = -1;
    private bool _preferredListChecked;
    private string _activeSheet = "";
    private string _autoFocusSheet = "";
    private int _autoFocusCol = -1;

    private enum GridArea
    {
        None = 0,
        Cell = 1,
        ColumnHeader = 2,
        RowHeader = 3
    }

    private sealed class SheetLoadData
    {
        public SheetSnapshot? Snapshot;
        public List<CurveColumnOption> CurveOptions = new();
        public string? Error;
    }
    private sealed record OpenPrepared(
        WorkbookModel Wb,
        (string Sheet, ColumnMeta Col)? Best,
        string TargetSheet,
        SheetLoadData TargetLoad,
        string? Error,
        bool ShowOccupiedHint = false);

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
        // 显示前就设为最大化，避免先出现默认小窗再“闪大”
        WindowState = FormWindowState.Maximized;

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
        // “编辑”菜单统一放置：刷新、撤销、重做
        var reloadItem = MakeMenu("刷新", "重新从磁盘读取当前工作表", OnReload, false);
        var undoItem = MakeMenu("撤销", "撤销上次编辑（Ctrl+Z）", () => Undo(), false);
        var redoItem = MakeMenu("重做", "重做上次撤销（Ctrl+Y）", () => Redo(), false);
        _editMenu.DropDownItems.Add(reloadItem);
        _editMenu.DropDownItems.Add(new ToolStripSeparator());
        _editMenu.DropDownItems.Add(undoItem);
        _editMenu.DropDownItems.Add(redoItem);
        _editMenu.DropDownOpening += (s, e) =>
        {
            reloadItem.Enabled = _wb != null;
            undoItem.Enabled = _undo.Count > 0;
            redoItem.Enabled = _redo.Count > 0;
        };
        _tool.Items.Add(_editMenu);
        // “视图”菜单统一放置表格的视图切换：最大化/还原、置底/靠右
        _gridMaxItem = MakeMenu("表格最大化", "把表格铺满主窗口，便于编辑数据（F11）", ToggleGridMaximize, false);
        _gridBottomItem = MakeMenu("表格置底", "表格放在曲线区下方", () => SetGridSide(false), true);
        _gridRightItem = MakeMenu("表格靠右", "表格与曲线区左右并列", () => SetGridSide(true), false);
        _viewButton.DropDownItems.Add(_gridMaxItem);
        _viewButton.DropDownItems.Add(new ToolStripSeparator());
        _viewButton.DropDownItems.Add(_gridBottomItem);
        _viewButton.DropDownItems.Add(_gridRightItem);
        _tool.Items.Add(_viewButton);
        _gridButton.DropDownOpening += (s, e) => BuildGridToolbarMenu();
        _tool.Items.Add(_gridButton);
        UpdateLayoutMenu();
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
        var refreshBtn = MakeButton("刷新列名", RefreshColumnNames);
        var clearBtn = MakeButton("清空选择", () =>
        {
            for (int i = 0; i < _colsChecked.Items.Count; i++)
                _colsChecked.SetItemChecked(i, false);
        });
        int btnGap = 4;
        int halfW = (lw - btnGap) / 2;
        refreshBtn.Size = new Size(halfW, 26);
        refreshBtn.Margin = new Padding(0);
        clearBtn.Size = new Size(lw - btnGap - halfW, 26);
        clearBtn.Margin = new Padding(btnGap, 0, 0, 0);
        var colBtnPanel = new FlowLayoutPanel
        {
            AutoSize = false,
            Width = lw,
            Height = 26,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0),
            Margin = new Padding(0),
            BackColor = Color.Transparent
        };
        colBtnPanel.Controls.Add(refreshBtn);
        colBtnPanel.Controls.Add(clearBtn);
        L(colBtnPanel, 26);
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
        R(MakeButton("填充空白", FillBlanks), 30);
		R(MakeLabel("随机扰动幅度:"), 20);
        R(_randUpDown, 28);
        R(MakeButton("随机扰动(选中点)", () => BatchRandom((double)_randUpDown.Value)), 30);
        R(MakeButton("曲线平滑(选中点)", BatchSmoothSelected), 30);
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
        R(Section("拟合"));
        foreach (var m in System.Enum.GetValues<FitModel>())
            _fitTypeCombo.Items.Add(CurveFit.LabelOf(m));
        _fitTypeCombo.SelectedIndex = 0;
        _fitDegree.Enabled = false;
        R(MakeLabel("类型:"), 20);
        R(_fitTypeCombo, 28);
        R(MakeLabel("多项式次数(仅多项式):"), 20);
        R(_fitDegree, 28);
        var fitButtons = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Height = 30 };
        fitButtons.Controls.Add(MakeButton("预览", OnFitPreview));
        fitButtons.Controls.Add(MakeButton("应用", OnFitApply));
        fitButtons.Controls.Add(MakeButton("清除", ClearFitPreview));
        R(fitButtons, 30);
        R(MakeButton("高级拟合...", OpenAdvancedFit), 30);
        R(_fitInfo, 88);
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
        _grid.AllowUserToOrderColumns = true;           // 列名按住拖拽移动顺序
        _grid.ColumnDisplayIndexChanged += (s, e) => { if (!_suppressColumnOrderDirty) _columnOrderDirty = true; };
        _grid.ColumnWidthChanged += OnGridColumnWidthChanged;
        _grid.RowHeadersVisible = true;
        _grid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;  // 点击只选中，键入/F2/双击进入编辑（Excel 式）
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
        _grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
        _grid.MultiSelect = true;                        // 支持框选多个单元格
        _grid.BackgroundColor = Color.White;
        _grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        _grid.ContextMenuStrip = _gridMenu;
        _gridMenu.ShowImageMargin = false;

        _menu.ShowImageMargin = false;
        _curve.ContextMenuStrip = _menu;

        BuildTooltips();
    }

    private void BuildEvents()
    {
        _colsChecked.ItemCheck += (s, e) =>
        {
            if (_suppressRebuild) return;
            _preferredListIndex = e.Index;
            _preferredListChecked = e.NewValue == CheckState.Checked;
            // 勾选集合变化只影响“曲线序列”：不重建整个网格、
            // 不清空撤销历史、不滚动列表，否则重绘错位会让用户误以为
            // 别项也被取消了（点击两次时尤其明显）。
            BeginInvokeSafe(ApplyCheckedColumns);
        };
        _colsChecked.MouseMove += OnColumnListMouseMove;
        _xCombo.SelectedIndexChanged += (s, e) => OnSelectionChangedSafe();
        _stepUpDown.ValueChanged += (s, e) => _curve.KeyboardStep = (double)_stepUpDown.Value;
        _fitTypeCombo.SelectedIndexChanged += (s, e) => _fitDegree.Enabled = SelectedFitModel() == FitModel.Polynomial;

        _curve.PointsChanged += OnPointsChanged;
        _curve.EditCommitted += OnEditCommitted;
        _curve.SelectionChanged += OnSelectionChangedUi;
        _curve.HoverChanged += s => _hoverLabel.Text = s;
        _grid.CellEndEdit += OnGridCellEdited;
        _grid.SelectionChanged += OnGridSelectionChanged;
        _grid.MouseDown += OnGridMouseDown;
        _grid.MouseMove += OnGridMouseMove;
        _grid.MouseUp += OnGridMouseUp;
        _grid.ColumnHeaderMouseClick += OnGridColumnHeaderClick;
        _gridMenu.Opening += (s, e) => BuildGridContextMenu();
        _grid.ColumnHeaderMouseDoubleClick += OnGridColumnHeaderDoubleClick;
        _menu.Opening += (s, e) => BuildContextMenu();

        _autoSaveTimer.Tick += (s, e) => { _autoSaveTimer.Stop(); CommitPending(); };
        _reloadTimer.Tick += (s, e) => { _reloadTimer.Stop(); ReloadFromDisk(); };

        KeyPreview = true;
        KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.F11) { ToggleGridMaximize(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.S) { OnSave(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.Z) { Undo(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.Y) { Redo(); e.Handled = true; }
            else if (_grid.IsCurrentCellInEditMode)
            {
                // 单元格正在编辑：C/X/V/Delete 交给编辑框自己处理
                if (e.KeyCode == Keys.Delete || (e.Control && (e.KeyCode == Keys.C || e.KeyCode == Keys.X || e.KeyCode == Keys.V)))
                    return;
            }
            else if (e.Control && e.KeyCode == Keys.C) { CopySelection(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.X) { CutSelection(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.V) { PasteFromClipboard(); e.Handled = true; }
            // Delete 只清除“表格(excel)”拥有焦点时的单元格值；曲线选中点、曲线获得焦点时不触发删除
            else if (e.KeyCode == Keys.Delete && !e.Control && _grid.ContainsFocus) { ClearSelection(); e.Handled = true; }
        };
    }

    /// <summary>曲线列 (Y) 多选：当鼠标悬停且该项文字显示不全时，提示完整名称。</summary>
    private void OnColumnListMouseMove(object? sender, MouseEventArgs e)
    {
        int idx = _colsChecked.IndexFromPoint(e.Location);
        string? text = idx >= 0 ? _colsChecked.Items[idx]?.ToString() : null;
        if (string.IsNullOrEmpty(text))
        {
            if (_colCheckedTipText != null)
            {
                _colCheckedTipText = null;
                _tip.Hide(_colsChecked);
                _tip.SetToolTip(_colsChecked, "");
            }
            return;
        }

        var rect = _colsChecked.GetItemRectangle(idx);
        float scale = _colsChecked.DeviceDpi / 96f;
        int checkboxGap = (int)Math.Ceiling(16 * scale) + 2; // 复选框 + 文字左边距
        int avail = Math.Max(0, rect.Width - checkboxGap);
        int textWidth = TextRenderer.MeasureText(text, _colsChecked.Font).Width;
        string tip = textWidth > avail ? text : "";
        if (tip != _colCheckedTipText)
        {
            _colCheckedTipText = tip;
            if (tip.Length == 0) _tip.Hide(_colsChecked);
            _tip.SetToolTip(_colsChecked, tip);
        }
    }

    // ---------- 打开 / 工作表 ----------
    private void OnShown()
    {
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
        if (_busyLoading) return;
        _busyLoading = true;
        ShowSplash("正在打开 " + Path.GetFileName(path) + " ...");
        Task.Run(() => PrepareOpen(path))
            .ContinueWith(t => BeginInvokeSafe(() => FinishOpen(path, t.Result)), TaskScheduler.Default);
    }

    private static OpenPrepared PrepareOpen(string path)
    {
        try
        {
            var wb = new WorkbookModel();
            wb.Open(path);
            var best = FindBestDataColumn(wb);
            string target = best?.Sheet ?? wb.SheetNames.FirstOrDefault(s => !ShouldHideSheet(s)) ?? "";
            var tl = LoadSheetDataSync(wb, target);
            return new OpenPrepared(wb, best, target, tl, null);
        }
        catch (Exception ex)
        {
            // 文件被占用（如 Excel 已打开）时 OpenXml 通常会抛 IOException
            bool occupied = ex is IOException or UnauthorizedAccessException;
            return new OpenPrepared(null!, null, "", new SheetLoadData { Error = ex.Message }, ex.Message, occupied);
        }
    }

    private void FinishOpen(string path, OpenPrepared prep)
    {
        _busyLoading = false;
        if (prep.Error != null)
        {
            CloseSplash();
            string hint = prep.ShowOccupiedHint
                ? "\n\n该文件当前可能被其他程序（如 Excel）占用，请关闭后重试。"
                : "";
            MessageBox.Show(this, "无法打开文件：\n" + prep.Error + hint, "错误",
                MessageBoxButtons.OK, MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            return;
        }
        _wb?.Dispose();
        _wb = prep.Wb;
        _settings.LastFile = path;
        SetupWatcher(path);
        BuildSheetStrip();
        _autoFocusSheet = prep.Best?.Sheet ?? "";
        _autoFocusCol = prep.Best?.Col.ColumnIndex ?? -1;
        _activeSheet = prep.TargetSheet;
        HighlightActiveSheet();
        ApplySheetData(prep.TargetLoad);
        UpdateTitle();
        CloseSplash();
    }

    /// <summary>关闭当前工作簿并清空界面状态。</summary>
    private void CloseCurrentFile()
    {
        if (_busyLoading || _wb == null) return;
        if (HasUnsavedChanges())
        {
            var r = MessageBox.Show(this, "有未保存的改动，是否保存？", "确认", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (r == DialogResult.Cancel) return;
            if (r == DialogResult.Yes)
            {
                if (!CommitPending(out var err))
                {
                    MessageBox.Show(this, string.IsNullOrEmpty(err) ? "保存失败，请检查文件是否被占用。" : err,
                        "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
        }
        _watcher?.Dispose();
        _watcher = null;
        _autoSaveTimer.Stop();
        _reloadTimer.Stop();
        _wb?.Dispose();
        _wb = null;
        _snapshot = null;
        _activeSheet = "";
        _autoFocusSheet = "";
        _autoFocusCol = -1;
        _checkedCols = new List<CurveColumnOption>();
        _xColumn = null;
        _activeYColumn = null;
        _series.Clear();
        _curve.SetSeries(new List<CurveSeriesView>(), -1);
        _curve.ClearSelection();
        _curve.XAxisLabel = "";
        _curve.YAxisLabel = "";
        _committed.Clear();
        _editing.Clear();
        _rowToGridIndex.Clear();
        _dirtyCells.Clear();
        _pendingUndoRows.Clear();
        _subEditOldText.Clear();
        _undo.Clear();
        _redo.Clear();
        _pendingStructure.Clear();
        _pendingHeaderRename.Clear();
        _pendingHeaderAlign.Clear();
        _pendingColumnAlign.Clear();
        _structureQueued = false;
        _columnOrderDirty = false;
        _columnWidthDirty = false;
        _sheetOrderDirty = false;
        _grid.Columns.Clear();
        _grid.Rows.Clear();
        _colsChecked.Items.Clear();
        _xCombo.Items.Clear();
        _xCombo.Items.Add("行号");
        _xCombo.SelectedIndex = 0;
        _statLabel.Text = "";
        _sheetStrip.SuspendLayout();
        _sheetStrip.Controls.Clear();
        _sheetButtons.Clear();
        _addSheetButton = null;
        _sheetStrip.ResumeLayout();
        _statusLabel.Text = "已关闭当前文件";
        RebuildOpenMenu();
        UpdateTitle();
    }

    /// <summary>用系统关联程序（Excel）打开当前工作簿。</summary>
    private void OpenCurrentInExcel()
    {
        if (_wb == null) return;
        string path = _wb.Path;
        if (!File.Exists(path))
        {
            MessageBox.Show(this, "当前文件不存在，无法用 Excel 打开。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        try
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "无法用 Excel 打开当前文件：\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void StartSheetLoad(string name)
    {
        if (_busyLoading || _wb == null) return;
        _busyLoading = true;
        _statusLabel.Text = "正在加载工作表 " + name + " ...";
        var wb = _wb;
        Task.Run(() => LoadSheetDataSync(wb, name))
            .ContinueWith(t => BeginInvokeSafe(() =>
            {
                _busyLoading = false;
                ApplySheetData(t.Result);
            }), TaskScheduler.Default);
    }

    private static SheetLoadData LoadSheetDataSync(WorkbookModel wb, string sheet)
    {
        var d = new SheetLoadData();
        try
        {
            d.Snapshot = wb.LoadSheet(sheet);
            d.CurveOptions = SubCurveHelper.BuildOptions(d.Snapshot)
                .Where(o => !ShouldHideColumn(o.Column))
                .ToList();
        }
        catch (Exception ex)
        {
            d.Error = ex.Message;
        }
        return d;
    }

    private void ApplySheetData(SheetLoadData d)
    {
        if (d.Error != null)
        {
            _statusLabel.Text = "读取工作表失败：" + d.Error;
            return;
        }
        _snapshot = d.Snapshot;
        _columnOrderDirty = false;
        _columnWidthDirty = false;
        _sheetOrderDirty = false;
        var curveOptions = d.CurveOptions;
        _autoFocusCol = -1;
        _colsChecked.Items.Clear();
        foreach (var o in curveOptions) _colsChecked.Items.Add(o);
        _xCombo.Items.Clear();
        _xCombo.Items.Add("行号");
        foreach (var o in curveOptions) _xCombo.Items.Add(o);
        _xCombo.SelectedIndex = 0;
        RebuildFromSelection();
        _statusLabel.Text = $"已加载 [{_activeSheet}]：{d.Snapshot!.DataRowCount} 行 × {d.Snapshot.ColumnCount} 列，曲线子列 {curveOptions.Count} 条";
    }

    private void BeginInvokeSafe(Action action)
    {
        if (IsDisposed || !IsHandleCreated) return;
        try { BeginInvoke(action); }
        catch { }
    }

    private void ShowSplash(string text)
    {
        if (_splash == null || _splash.IsDisposed)
            _splash = new SplashForm();
        _splash.SetText(text);
        if (!_splash.Visible) _splash.Show();
        _splash.BringToFront();
        _splash.Update();
    }

    private void CloseSplash()
    {
        if (_splash != null && !_splash.IsDisposed)
        {
            _splash.Hide();
            _splash.Dispose();
        }
        _splash = null;
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

    private static (string Sheet, ColumnMeta Col)? FindBestDataColumn(WorkbookModel wb)
    {
        if (wb == null) return null;
        (string Sheet, ColumnMeta Col, int Count)? best = null;
        foreach (var sheet in wb.SheetNames)
        {
            if (ShouldHideSheet(sheet)) continue;
            SheetSnapshot sn;
            try { sn = wb.LoadSheet(sheet); } catch { continue; }
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
        CancelSheetRename();
        var prevOrder = _sheetButtons.Count > 0
            ? _sheetButtons.Select(b => (string)b.Tag!).ToList()
            : null;
        _sheetStrip.SuspendLayout();
        _sheetStrip.Controls.Clear();
        _sheetButtons.Clear();
        _addSheetButton = null;
        if (_wb == null) { _sheetStrip.ResumeLayout(); return; }
        foreach (var name in GetDisplayOrderedSheetNames(prevOrder))
        {
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
            b.Click += (s, e) =>
            {
                if (_suppressSheetClick) { _suppressSheetClick = false; return; }
                ActivateSheet(((Button)s!).Tag!.ToString()!);
            };
            b.MouseDown += OnSheetMouseDown;
            b.MouseMove += OnSheetMouseMove;
            b.MouseUp += OnSheetMouseUp;
            _sheetStrip.Controls.Add(b);
            _sheetButtons.Add(b);
        }
        var addBtn = new Button
        {
            Text = "+",
            Tag = "__add__",
            Size = new Size(32, 26),
            FlatStyle = FlatStyle.Flat,
            Margin = new Padding(3, 2, 1, 2),
            Cursor = Cursors.Hand,
            BackColor = Color.FromArgb(226, 230, 236),
            ForeColor = Color.FromArgb(50, 56, 64),
            Font = new Font("Microsoft YaHei UI", 12f, FontStyle.Regular)
        };
        addBtn.FlatAppearance.BorderSize = 0;
        addBtn.Click += (s, e) => OnAddSheet();
        _addSheetButton = addBtn;
        _sheetStrip.Controls.Add(addBtn);
        _tip.SetToolTip(addBtn, "新建工作表");
        _sheetStrip.ResumeLayout();
        HighlightActiveSheet();
    }

    /// <summary>按保存的工作表显示顺序返回用户可见工作表名列表（未记录的排在末尾）。</summary>
    private List<string> GetDisplayOrderedSheetNames(IReadOnlyList<string>? priorityOrder = null)
    {
        var visible = _wb!.SheetNames.Where(s => !ShouldHideSheet(s)).ToList();
        // 优先使用当前标签栏既有的显示顺序（含未保存的拖拽），否则用保存的顺序
        var baseOrder = priorityOrder != null && priorityOrder.Count > 0
            ? priorityOrder
            : _wb.GetSheetDisplayOrder();
        if (baseOrder == null || baseOrder.Count == 0) return visible;
        var result = new List<string>();
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in baseOrder)
            if (visible.Contains(name, StringComparer.OrdinalIgnoreCase) && used.Add(name))
                result.Add(name);
        foreach (var name in visible)
            if (used.Add(name)) result.Add(name);
        return result;
    }

    /// <summary>把当前 <see cref="_sheetStrip"/> 中除“+”外的按钮顺序同步到 <see cref="_sheetButtons"/>。</summary>
    private void SyncSheetButtons()
    {
        _sheetButtons.Clear();
        foreach (var c in _sheetStrip.Controls.Cast<Control>())
            if (c is Button b && b != _addSheetButton && !Equals(b.Tag, "__add__"))
                _sheetButtons.Add(b);
    }

    /// <summary>把“+”按钮固定在标签栏最右侧。</summary>
    private void EnsureAddSheetLast()
    {
        if (_addSheetButton == null) return;
        int addIndex = _sheetStrip.Controls.IndexOf(_addSheetButton);
        int lastIndex = _sheetStrip.Controls.Count - 1;
        if (addIndex != lastIndex)
            _sheetStrip.Controls.SetChildIndex(_addSheetButton, lastIndex);
    }

    /// <summary>sheet 标签按住拖拽：记录起点。</summary>
    private void OnSheetMouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button != MouseButtons.Left) return;
        // 双击进入改名（MouseDown 的 Clicks==2 比 DoubleClick 事件更可靠）
        if (e.Clicks == 2)
        {
            BeginSheetRename((Button)sender!);
            return;
        }
        _dragSheetBtn = (Button)sender!;
        _dragSheetStart = e.Location;
        _draggingSheet = false;
        _suppressSheetClick = false;
        _dragSheetBtn.Capture = true;
    }

    /// <summary>sheet 标签拖拽过程中：移动到其它标签位置时交换顺序。</summary>
    private void OnSheetMouseMove(object? sender, MouseEventArgs e)
    {
        if (_dragSheetBtn == null || sender != _dragSheetBtn) return;
        if (!_draggingSheet)
        {
            if (Math.Abs(e.X - _dragSheetStart.X) < SystemInformation.DragSize.Width &&
                Math.Abs(e.Y - _dragSheetStart.Y) < SystemInformation.DragSize.Height)
                return;
            _draggingSheet = true;
        }
        var pos = _sheetStrip.PointToClient(Cursor.Position);
        if (_sheetStrip.GetChildAtPoint(pos) is not Button target) return;
        if (target == _dragSheetBtn || target == _addSheetButton) return;
        int from = _sheetStrip.Controls.IndexOf(_dragSheetBtn);
        int to = _sheetStrip.Controls.IndexOf(target);
        if (from == to) return;
        _sheetStrip.Controls.SetChildIndex(_dragSheetBtn, to);
        EnsureAddSheetLast();
        SyncSheetButtons();
        _sheetStrip.PerformLayout();
    }

    /// <summary>sheet 标签拖拽结束：标记工作表顺序为待保存。</summary>
    private void OnSheetMouseUp(object? sender, MouseEventArgs e)
    {
        if (_dragSheetBtn == null) return;
        _dragSheetBtn.Capture = false;
        bool changed = _draggingSheet;
        _dragSheetBtn = null;
        _draggingSheet = false;
        if (changed)
        {
            _suppressSheetClick = true;
            _sheetOrderDirty = true;
            HighlightActiveSheet();
            PersistSheetOrder();
            UpdateTitle();
        }
    }

    /// <summary>把当前标签栏顺序写入文件；成功则清除脏标记，失败保留以便下次保存重试。</summary>
    private void PersistSheetOrder()
    {
        if (_wb == null) return;
        var order = GetDisplayedSheetOrder();
        if (order.Count == 0) return;
        var saved = _wb.GetSheetDisplayOrder();
        if (saved != null && saved.SequenceEqual(order))
        {
            _sheetOrderDirty = false;
            return;
        }
        _selfWrite = true;
        bool ok = _wb.TryWriteSheetOrder(order, out var err);
        _selfWrite = false;
        if (ok) _sheetOrderDirty = false;
        else _statusLabel.Text = "⚠ 保存工作表顺序失败：" + (err ?? "");
    }

    /// <summary>返回当前标签栏显示的工作表顺序（名称序列）。</summary>
    private List<string> GetDisplayedSheetOrder()
        => _sheetButtons.Select(b => (string)b.Tag!).ToList();

    /// <summary>点击“+”新建一张工作表并切换到它。</summary>
    private void OnAddSheet()
    {
        if (_wb == null) return;
        string name = MakeUniqueSheetName();
        _selfWrite = true;
        bool ok = _wb.TryAddSheet(name, out var err);
        _selfWrite = false;
        if (!ok)
        {
            _statusLabel.Text = "⚠ " + (err ?? "新建工作表失败");
            MessageBox.Show(this, err ?? "新建工作表失败", "新建工作表失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        _statusLabel.Text = "已新建工作表 " + name;
        _autoFocusSheet = name;
        _autoFocusCol = 1;
        BuildSheetStrip();
        ActivateSheet(name);
    }

    /// <summary>生成一个不与现有工作表重名的名称（新表1、新表2…）。</summary>
    private string MakeUniqueSheetName()
    {
        var existing = new HashSet<string>(_wb!.SheetNames, StringComparer.OrdinalIgnoreCase);
        int n = 1;
        while (existing.Contains("新表" + n)) n++;
        return "新表" + n;
    }

    /// <summary>双击 sheet 标签：在按钮原位弹出内联编辑框修改工作表名。</summary>
    private void BeginSheetRename(Button btn)
    {
        if (_wb == null || btn.Tag is not string name) return;
        CommitHeaderEdit(); // 先结束可能存在的列名编辑
        // 双击进入改名，结束拖拽状态并抑制随后的第二次 Click 触发工作表切换
        if (_dragSheetBtn != null)
        {
            _dragSheetBtn.Capture = false;
            _dragSheetBtn = null;
            _draggingSheet = false;
        }
        _suppressSheetClick = true;
        var tb = new TextBox
        {
            Text = name,
            TextAlign = HorizontalAlignment.Center,
            Font = btn.Font,
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = Color.White,
            Multiline = false,
            TabStop = false,
            Bounds = btn.Bounds
        };
        tb.KeyDown += OnSheetRenameKeyDown;
        tb.Leave += OnSheetRenameLeave;
        // 放在父容器上叠盖（FlowLayoutPanel 会把子控件当作流式项，不适合叠加）
        if (_gridPane == null)
        {
            tb.Dispose();
            return;
        }
        Point pos = _gridPane.PointToClient(_sheetStrip.PointToScreen(btn.Location));
        tb.Bounds = new Rectangle(pos, btn.Size);
        _gridPane.Controls.Add(tb);
        tb.BringToFront();
        tb.Focus();
        tb.SelectAll();
        _sheetRenameEdit = tb;
        _sheetRenameOldName = name;
    }

    private void OnSheetRenameKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter) { CommitSheetRename(); e.Handled = true; e.SuppressKeyPress = true; }
        else if (e.KeyCode == Keys.Escape) { CancelSheetRename(); e.Handled = true; e.SuppressKeyPress = true; }
    }

    private void OnSheetRenameLeave(object? sender, EventArgs e) => CommitSheetRename();

    /// <summary>提交 sheet 改名并写回文件。</summary>
    private void CommitSheetRename()
    {
        var tb = _sheetRenameEdit;
        if (tb == null) return;
        _sheetRenameEdit = null;
        string newName = tb.Text.Trim();
        string oldName = _sheetRenameOldName;
        _sheetRenameOldName = "";
        tb.Dispose();

        if (_wb == null || string.IsNullOrWhiteSpace(oldName)) return;
        if (string.IsNullOrWhiteSpace(newName))
        {
            _statusLabel.Text = "⚠ 工作表名不能为空";
            return;
        }
        if (string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase)) return;

        // 正在改活动表且存在未保存改动：先保存/丢弃，避免后续结构改动引用旧名
        if (string.Equals(oldName, _activeSheet, StringComparison.OrdinalIgnoreCase) && HasUnsavedChanges())
        {
            var r = MessageBox.Show(this, "重命名前是否保存当前改动？", "确认",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (r == DialogResult.Cancel) return;
            if (r == DialogResult.Yes) CommitPending();
            else DiscardUnsaved();
        }

        _selfWrite = true;
        bool ok = _wb.TryRenameSheet(oldName, newName, out var err);
        _selfWrite = false;
        if (!ok)
        {
            _statusLabel.Text = "⚠ " + (err ?? "重命名失败");
            MessageBox.Show(this, err ?? "重命名失败", "重命名工作表失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        bool wasActive = string.Equals(oldName, _activeSheet, StringComparison.OrdinalIgnoreCase);
        if (wasActive)
        {
            _activeSheet = newName;
            if (_snapshot != null && string.Equals(_snapshot.SheetName, oldName, StringComparison.OrdinalIgnoreCase))
                _snapshot.SheetName = newName;
            if (string.Equals(_autoFocusSheet, oldName, StringComparison.OrdinalIgnoreCase))
                _autoFocusSheet = newName;
        }
        // 原地更新标签文字与 Tag，避免整条重建在用户点击其它标签时吞掉那次点击
        foreach (var b in _sheetButtons)
            if (string.Equals((string)b.Tag!, oldName, StringComparison.OrdinalIgnoreCase))
            {
                b.Text = newName;
                b.Tag = newName;
                break;
            }
        HighlightActiveSheet();
        _statusLabel.Text = "已重命名工作表 " + oldName + " → " + newName;
        UpdateTitle();
    }

    /// <summary>取消 sheet 改名，直接关闭内联编辑框。</summary>
    private void CancelSheetRename()
    {
        var tb = _sheetRenameEdit;
        if (tb == null) return;
        _sheetRenameEdit = null;
        _sheetRenameOldName = "";
        tb.Dispose();
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
        StartSheetLoad(name);
    }

    private bool HasUnsavedChanges() => _dirtyCells.Count > 0 || _structureQueued || _pendingHeaderRename.Count > 0 || _pendingHeaderAlign.Count > 0 || _pendingColumnAlign.Count > 0 || _columnOrderDirty || _columnWidthDirty || _sheetOrderDirty;

    private void DiscardUnsaved()
    {
        bool hadStructure = _structureQueued;
        _dirtyCells.Clear();
        _pendingStructure.Clear();
        _pendingHeaderRename.Clear();
        _pendingHeaderAlign.Clear();
        _pendingColumnAlign.Clear();
        _structureQueued = false;
        _columnOrderDirty = false;
        _columnWidthDirty = false;
        _sheetOrderDirty = false;
        ApplySavedColumnOrder();
        _pendingUndoRows.Clear();
        _subEditOldText.Clear();
        _editing.Clear();
        _committed.Clear();
        _undo.Clear();
        _redo.Clear();
        if (hadStructure && _wb != null) _wb.RefreshMeta();
        UpdateTitle();
    }

    private void RebuildFromSelection()
    {
        if (_wb == null || _snapshot == null) return;
        _suppressRebuild = true;
        try
        {
            var preserved = _activeYColumn;
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
            int activeIndex = 0;
            if (_preferredListIndex >= 0 && _preferredListChecked &&
                _preferredListIndex < _colsChecked.Items.Count &&
                _colsChecked.Items[_preferredListIndex] is CurveColumnOption preferred)
            {
                int pi = _checkedCols.FindIndex(o => ReferenceEquals(o, preferred));
                if (pi >= 0) activeIndex = pi;
            }
            else if (preserved != null)
            {
                int pi = _checkedCols.FindIndex(o => IsSameCurveOption(o, preserved));
                if (pi >= 0) activeIndex = pi;
            }
            _activeYColumn = _checkedCols[activeIndex];

            _series.Clear();
            for (int i = 0; i < _checkedCols.Count; i++)
                _series.Add(BuildSeriesForOption(_checkedCols[i], i == activeIndex));

            _curve.SetSeries(_series, activeIndex);
            _curve.XAxisLabel = _xColumn != null ? _xColumn.DisplayName : "行号";
            _curve.YAxisLabel = _activeYColumn.DisplayName;
            SyncActiveColumnHighlight(activeIndex);

            _committed.Clear(); _editing.Clear();
            _pendingUndoRows.Clear();
            _subEditOldText.Clear();
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
        finally
        {
            _preferredListIndex = -1;
            _preferredListChecked = false;
            _suppressRebuild = false;
        }
    }

    /// <summary>
    /// 勾选“曲线列 (Y) 多选”后的轻量刷新：只按最新勾选集合重建曲线序列并同步高亮，
    /// 不重建数据网格、不清空撤销历史、不滚动列表。
    /// 解决“点一项取消勾选时感觉别项也被取消 / 列表卡顿”的问题。
    /// </summary>
    private void ApplyCheckedColumns()
    {
        if (_suppressRebuild || _wb == null || _snapshot == null) return;
        CommitPending();

        _suppressRebuild = true;
        try
        {
            var preserved = _activeYColumn;
            _checkedCols = _colsChecked.CheckedItems.Cast<CurveColumnOption>().ToList();

            if (_checkedCols.Count == 0)
            {
                _series.Clear();
                _curve.SetSeries(new List<CurveSeriesView>(), -1);
                _activeYColumn = null;
                _committed.Clear();
                _editing.Clear();
                _curve.XAxisLabel = "行号";
                _curve.YAxisLabel = "";
                HighlightEditableColumn();
                UpdateStats();
                UpdateTitle();
                return;
            }

            _xColumn = _xCombo.SelectedIndex > 0 ? _xCombo.SelectedItem as CurveColumnOption : null;
            int activeIndex = 0;
            if (_preferredListIndex >= 0 && _preferredListChecked &&
                _preferredListIndex < _colsChecked.Items.Count &&
                _colsChecked.Items[_preferredListIndex] is CurveColumnOption preferred)
            {
                int pi = _checkedCols.FindIndex(o => ReferenceEquals(o, preferred));
                if (pi >= 0) activeIndex = pi;
            }
            else if (preserved != null)
            {
                int pi = _checkedCols.FindIndex(o => IsSameCurveOption(o, preserved));
                if (pi >= 0) activeIndex = pi;
            }
            _activeYColumn = _checkedCols[activeIndex];

            bool activeChanged = preserved == null ||
                                 !IsSameCurveOption(_checkedCols[activeIndex], preserved);

            _series.Clear();
            for (int i = 0; i < _checkedCols.Count; i++)
                _series.Add(BuildSeriesForOption(_checkedCols[i], i == activeIndex));

            _curve.SetSeries(_series, activeIndex);
            _curve.XAxisLabel = _xColumn != null ? _xColumn.DisplayName : "行号";
            _curve.YAxisLabel = _activeYColumn.DisplayName;
            // 取消勾选时不要移动蓝色高亮：让高亮停留在用户点击的那一项，
            // 避免“取消一个把蓝色选择跳到别的项去”的观感。
            HighlightEditableColumn();

            // 仅当当前编辑列发生变化时才重设提交/编辑基线；撤销历史保留。
            if (activeChanged)
            {
                _committed.Clear();
                _editing.Clear();
                _pendingUndoRows.Clear();
                _subEditOldText.Clear();
                foreach (var p in _curve.Points)
                {
                    _committed[p.RowNumber] = (p.X, p.Y);
                    _editing[p.RowNumber] = (p.X, p.Y);
                }
            }
            UpdateStats();
            UpdateTitle();
        }
        finally
        {
            _preferredListIndex = -1;
            _preferredListChecked = false;
            _suppressRebuild = false;
        }
    }

    private CurveSeriesView BuildSeriesForOption(CurveColumnOption opt, bool editable)
    {
        var points = new List<CurvePoint>();
        var snap = _snapshot!;
        for (int gi = 0; gi < snap.RowNumbers.Count; gi++)
        {
            int row = snap.RowNumbers[gi];
            if (!SubCurveHelper.TryReadValue(snap, opt, gi, out var val)) continue;
            // 行号模式用“排除表头行的数据行序号”，与表格“行号”列保持一致（物理行号仍保留给编辑映射）
            double x = gi + 1; bool xEditable = false;
            if (_xColumn != null)
            {
                if (SubCurveHelper.TryReadValue(snap, _xColumn, gi, out var xv))
                { x = xv; xEditable = true; }
                else continue;
            }
            points.Add(new CurvePoint(x, val, row, xEditable) { DataRow = gi + 1 });
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
        _subEditOldText.Clear();
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

    private static bool IsSameCurveOption(CurveColumnOption a, CurveColumnOption b)
        => a.Column.ColumnIndex == b.Column.ColumnIndex &&
           a.SubIndex == b.SubIndex &&
           a.JsonIndex == b.JsonIndex &&
           string.Equals(a.JsonId, b.JsonId, StringComparison.Ordinal) &&
           a.IsJsonValue == b.IsJsonValue;

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
                    _subEditOldText.TryAdd((col, row), oldRaw);
                    UpdateSnapshotCell(row, col, SubCurveHelper.SetValue(oldRaw, _activeYColumn, p.Y));
                }
                else
                    UpdateSnapshotCell(row, col, FormatCellValue(p.Y, _activeYColumn.IsInteger));
            }
            if (_xColumn != null && p.XEditable)
            {
                int xCol = _xColumn.Column.ColumnIndex;
                _dirtyCells.Add((xCol, row));
                if (_xColumn.IsSubCurve)
                {
                    string oldRaw = CellText(xCol, row);
                    _subEditOldText.TryAdd((xCol, row), oldRaw);
                    UpdateSnapshotCell(row, xCol, SubCurveHelper.SetValue(oldRaw, _xColumn, p.X));
                }
                else
                    UpdateSnapshotCell(row, xCol, FormatCellValue(p.X, _xColumn.IsInteger));
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
            if (_activeYColumn != null)
            {
                int col = _activeYColumn.Column.ColumnIndex;
                if (_activeYColumn.IsSubCurve)
                {
                    if (_subEditOldText.Remove((col, row), out var oldRaw))
                        items.Add((col, row, oldRaw, CellText(col, row)));
                }
                else
                    items.Add((col, row, FormatCellValue(old.Y, _activeYColumn.IsInteger), FormatCellValue(cur.Y, _activeYColumn.IsInteger)));
            }
            if (_xColumn != null && TryGetPoint(row, out var p) && p.XEditable && Math.Abs(old.X - cur.X) > 1e-12)
            {
                int xCol = _xColumn.Column.ColumnIndex;
                if (_xColumn.IsSubCurve)
                {
                    if (_subEditOldText.Remove((xCol, row), out var oldRaw))
                        items.Add((xCol, row, oldRaw, CellText(xCol, row)));
                }
                else
                    items.Add((xCol, row, FormatCellValue(old.X, _xColumn.IsInteger), FormatCellValue(cur.X, _xColumn.IsInteger)));
            }
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

    /// <summary>把当前未保存的改动写回工作簿。返回是否成功。</summary>
    private bool CommitPending() => CommitPending(out _);

    /// <summary>把当前未保存的改动写回工作簿。返回是否成功；失败时把原因写入 <paramref name="error"/>。</summary>
    private bool CommitPending(out string? error)
    {
        error = null;
        if (_wb == null) return true;
        bool hasStructure = _pendingStructure.Count > 0;
        if (_dirtyCells.Count == 0 && !hasStructure && _pendingHeaderRename.Count == 0 && _pendingHeaderAlign.Count == 0 && _pendingColumnAlign.Count == 0 && !_columnOrderDirty && !_columnWidthDirty && !_sheetOrderDirty) return true;
        _autoSaveTimer.Stop();
        var sheet = _snapshot?.SheetName ?? _activeSheet;

        // 1) 先应用结构改动（行/列增删），使现有单元格平移到位
        if (hasStructure)
        {
            _selfWrite = true;
            bool okStructure = _wb.TryApplyStructure(_pendingStructure, out var errStructure);
            _selfWrite = false;
            if (!okStructure)
            {
                _statusLabel.Text = "⚠ " + errStructure;
                error = errStructure;
                UpdateTitle();
                return false;
            }
            _pendingStructure.Clear();
            _structureQueued = false;
        }

        var nums = new List<CellWrite>();
        var strs = new List<CellWriteString>();

        // 2) 表头重命名写回
        foreach (var (col, text) in _pendingHeaderRename)
            strs.Add(new CellWriteString(sheet, CellHelper.ToCellReference(col, _snapshot?.HeaderRow ?? 1), text));

        // 3) 常规单元格写回
        foreach (var (col, row) in _dirtyCells)
        {
            if (col < 0 || col >= _snapshot!.Columns.Count) continue;
            if (row == _snapshot.HeaderRow) continue; // 表头行不存放在数据网格中
            var meta = _snapshot.Columns[col];
            string text = CellText(col, row);
            if (meta.IsNumericScalar && CellHelper.TryParseDouble(text, out var v))
                nums.Add(new CellWrite(sheet, CellHelper.ToCellReference(col, row), v, meta.IsInteger, 6));
            else
                strs.Add(new CellWriteString(sheet, CellHelper.ToCellReference(col, row), text));
        }
        _selfWrite = true;
        bool ok = _wb.TryWriteCells(nums, out var errNum);
        bool okStr = _wb.TryWriteCellsString(strs, out var errStr);
        _selfWrite = false;

        // 4) 表头对齐写回（在字符串写回之后，保证重命名新建的表头单元格已存在）
        bool okHeader = true;
        string? errHeader = null;
        if (_pendingHeaderAlign.Count > 0)
        {
            int headerRow = _snapshot?.HeaderRow ?? 1;
            var headerAligns = new List<(string, int, int, HorizontalAlign)>();
            foreach (var col in _pendingHeaderAlign)
                if (col >= 0 && col < _snapshot!.Columns.Count)
                    headerAligns.Add((sheet, col, headerRow, _snapshot.HeaderAlignments[col]));
            _selfWrite = true;
            okHeader = _wb.TryWriteHeaderAlignments(headerAligns, out errHeader);
            _selfWrite = false;
            if (okHeader) _pendingHeaderAlign.Clear();
        }

        // 5) 列内容对齐写回（只处理已存在的数据单元格，覆盖整列）
        bool okColAlign = true;
        string? errColAlign = null;
        if (_pendingColumnAlign.Count > 0)
        {
            int headerRow = _snapshot?.HeaderRow ?? 1;
            int maxRow = _snapshot?.MaxRow ?? headerRow;
            var colAligns = new List<(string, int, int, int, HorizontalAlign)>();
            foreach (var col in _pendingColumnAlign)
                if (col >= 0 && col < _snapshot!.Columns.Count)
                    colAligns.Add((sheet, col, headerRow, maxRow, _snapshot.ColumnAlignments[col]));
            _selfWrite = true;
            okColAlign = _wb.TryWriteColumnAlignments(colAligns, out errColAlign);
            _selfWrite = false;
            if (okColAlign) _pendingColumnAlign.Clear();
        }

        // 6) 列显示顺序写回（用户按住表头拖拽移动列后的顺序，随文件持久化）
        bool okOrder = true;
        string? errOrder = null;
        bool orderChanged = _columnOrderDirty;
        if (orderChanged && _snapshot != null)
        {
            var order = BuildCurrentColumnOrder();
            _selfWrite = true;
            okOrder = _wb.TryWriteColumnOrder(_snapshot.SheetName, order, out errOrder);
            _selfWrite = false;
            if (okOrder) _columnOrderDirty = false;
        }

        // 7) 列宽写回（用户调整列宽后，按当前显示宽度写回 Excel 工作表的 <cols>）
        bool okWidth = true;
        string? errWidth = null;
        bool widthChanged = _columnWidthDirty;
        if (widthChanged && _snapshot != null)
        {
            var widths = BuildCurrentColumnWidths();
            _selfWrite = true;
            okWidth = _wb.TryWriteColumnWidths(_snapshot.SheetName, widths, out errWidth);
            _selfWrite = false;
            if (okWidth) _columnWidthDirty = false;
        }

        // 8) 工作表显示顺序写回（用户拖拽 sheet 标签后的顺序，随文件持久化）
        bool okSheetOrder = true;
        string? errSheetOrder = null;
        bool sheetOrderChanged = _sheetOrderDirty;
        if (sheetOrderChanged && _wb != null)
        {
            var sheetOrder = GetDisplayedSheetOrder();
            _selfWrite = true;
            okSheetOrder = _wb.TryWriteSheetOrder(sheetOrder, out errSheetOrder);
            _selfWrite = false;
            if (okSheetOrder) _sheetOrderDirty = false;
        }

        if (!ok)
        {
            _statusLabel.Text = "⚠ " + errNum;
            error = errNum;
            UpdateTitle();
            return false;
        }
        if (!okStr)
        {
            _statusLabel.Text = "⚠ " + errStr;
            error = errStr;
            UpdateTitle();
            return false;
        }
        if (!okHeader)
        {
            _statusLabel.Text = "⚠ " + errHeader;
            error = errHeader;
            UpdateTitle();
            return false;
        }
        if (!okColAlign)
        {
            _statusLabel.Text = "⚠ " + errColAlign;
            error = errColAlign;
            UpdateTitle();
            return false;
        }
        if (!okOrder)
        {
            _statusLabel.Text = "⚠ " + errOrder;
            error = errOrder;
            UpdateTitle();
            return false;
        }
        if (!okWidth)
        {
            _statusLabel.Text = "⚠ " + errWidth;
            error = errWidth;
            UpdateTitle();
            return false;
        }
        if (!okSheetOrder)
        {
            _statusLabel.Text = "⚠ " + errSheetOrder;
            error = errSheetOrder;
            UpdateTitle();
            return false;
        }
        else
        {
            int cellCount = nums.Count + strs.Count;
            bool metaSaved = orderChanged || widthChanged || sheetOrderChanged;
            _statusLabel.Text = metaSaved
                ? (cellCount > 0 ? $"已保存 {cellCount} 个单元格并保存布局信息" : "已保存布局信息")
                : $"已保存 {cellCount} 个单元格";
            _structureQueued = false;
            _pendingHeaderRename.Clear();
            _dirtyCells.Clear();
        }
        error = null;
        UpdateTitle();
        return true;
    }

    /// <summary>按当前 DataGridView 的显示顺序返回物理列索引序列（用于持久化）。</summary>
    private List<int> BuildCurrentColumnOrder()
    {
        var result = new List<int>();
        if (_grid == null) return result;
        foreach (var col in _grid.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.DisplayIndex))
        {
            if (col.Name == "行号") continue;
            int physical = col.Index - 1;
            if (physical >= 0) result.Add(physical);
        }
        return result;
    }

    /// <summary>把保存的列显示顺序应用到当前表格（行号列固定在最前）。</summary>
    private void ApplySavedColumnOrder()
    {
        if (_wb == null || _snapshot == null) return;
        var order = _wb.GetColumnOrder(_snapshot.SheetName);
        if (order == null || order.Count == 0) return;
        _suppressColumnOrderDirty = true;
        try
        {
            var wanted = new List<int>();
            var used = new HashSet<int>();
            foreach (int physical in order)
            {
                int gridIndex = physical + 1;
                if (gridIndex <= 0 || gridIndex >= _grid.Columns.Count) continue;
                if (!used.Add(gridIndex)) continue;
                wanted.Add(gridIndex);
            }
            // 补上未出现在顺序里的列，保持它们默认的相对顺序
            for (int g = 1; g < _grid.Columns.Count; g++)
                if (!used.Contains(g)) wanted.Add(g);
            for (int i = 0; i < wanted.Count; i++)
                _grid.Columns[wanted[i]].DisplayIndex = i + 1;
        }
        finally
        {
            _suppressColumnOrderDirty = false;
        }
    }

    /// <summary>用户调整列宽时同步到快照并标记为待保存（程序化设置列宽时被抑制）。</summary>
    private void OnGridColumnWidthChanged(object? sender, DataGridViewColumnEventArgs e)
    {
        if (_suppressColumnWidthDirty || _snapshot == null) return;
        var col = e.Column;
        if (col == null || col.Name == "行号") return;
        int physical = col.Index - 1;
        if (physical < 0 || physical >= _snapshot.Columns.Count) return;
        _snapshot.Columns[physical].Width = CellHelper.PixelsToCharacterWidth(col.Width);
        _columnWidthDirty = true;
    }

    /// <summary>按当前 DataGridView 的列宽构建写回清单（物理列索引 0 基，宽度为字符串位）。</summary>
    private List<(int Col, double Width)> BuildCurrentColumnWidths()
    {
        var result = new List<(int, double)>();
        if (_snapshot == null) return result;
        for (int physical = 0; physical < _snapshot.Columns.Count; physical++)
        {
            int gridIndex = physical + 1;
            if (gridIndex < 0 || gridIndex >= _grid.Columns.Count) continue;
            var col = _grid.Columns[gridIndex];
            if (col.Name == "行号") continue;
            result.Add((physical, CellHelper.PixelsToCharacterWidth(col.Width)));
        }
        return result;
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
        CancelHeaderEdit();
        _suppressColumnOrderDirty = true;
        _suppressColumnWidthDirty = true;
        _grid.Columns.Clear();
        if (_snapshot == null)
        {
            _grid.Rows.Clear();
            _suppressColumnOrderDirty = false;
            _suppressColumnWidthDirty = false;
            return;
        }
        var header = new DataGridViewTextBoxColumn
        {
            HeaderText = "行号",
            Name = "行号",
            ReadOnly = true,
            Frozen = true,
            Resizable = DataGridViewTriState.True,
            Width = 70,
            DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
        };
        _grid.Columns.Add(header);
        foreach (var col in _snapshot.Columns)
        {
            int colWidth = col.Width.HasValue
                ? CellHelper.CharacterWidthToPixels(col.Width.Value)
                : 130;
            var c = new DataGridViewTextBoxColumn
            {
                HeaderText = col.DisplayName,
                Name = "col" + col.ColumnIndex,
                ReadOnly = false,
                Resizable = DataGridViewTriState.True,
                MinimumWidth = 60,
                Width = colWidth,
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
            var headerAlign = ci >= 0 && ci < _snapshot.HeaderAlignments.Count
                ? _snapshot.HeaderAlignments[ci]
                : HorizontalAlign.Default;
            c.HeaderCell.Style.Alignment = MapHeaderAlign(headerAlign);
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
            // 行号列从 1 开始，即排除表头行后的数据行序号（物理行号仍通过映射保留）
            cells[0] = gi + 1;
            _rowToGridIndex[rowNum] = gi;
            var row = _snapshot.Grid[gi];
            for (int c = 0; c < _snapshot.ColumnCount; c++)
                cells[c + 1] = c < row.Length ? (row[c] ?? "") : "";
            _grid.Rows.Add(cells);
        }
        _grid.ClearSelection();
        ApplySavedColumnOrder();
        _suppressColumnOrderDirty = false;
        _suppressColumnWidthDirty = false;
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
        ApplyGridCellEdit(col, row, text);
        if (_autoSaveCheck.Checked) CommitPending();
    }

    /// <summary>
    /// 把一个单元格的新值提交到快照/曲线/撤销栈。供网格编辑结束、剪切、粘贴、删除键共用。
    /// </summary>
    private void ApplyGridCellEdit(int col, int row, string text)
        => ApplyCellEdits(new[] { (col, row, text) });

    /// <summary>
    /// 批量应用单元格新值（剪切/粘贴/清除共用）：一次性提交一条撤销记录，并同步曲线与脏点。
    /// 返回实际被改动的单元格（物理列、物理行），供调用方刷新网格显示。
    /// </summary>
    private List<(int Col, int Row)> ApplyCellEdits(IEnumerable<(int Col, int Row, string Text)> edits)
    {
        var changed = new List<(int, int, string, string)>();
        var touched = new List<(int Col, int Row)>();
        if (_snapshot == null) return touched;
        foreach (var (col, row, text) in edits)
        {
            if (col < 0 || col >= _snapshot.ColumnCount) continue;
            if (!_rowToGridIndex.TryGetValue(row, out var gi) || gi < 0 || gi >= _snapshot.Grid.Count) continue;
            string oldText = CellText(col, row);
            if (text == oldText) continue;

            // 先更新快照，再刷新曲线：子曲线列（数组/JSON）需要读到新值才能决定点的新增/删除
            UpdateSnapshotCell(row, col, text);
            ApplyPlottedCellChange(col, row, text);
            SyncActiveBaseline(col, row, text);
            _dirtyCells.Add((col, row));
            changed.Add((col, row, oldText, text));
            touched.Add((col, row));
        }
        if (changed.Count > 0)
        {
            _undo.Add(new EditCmd(changed));
            _redo.Clear();
            UpdateStats();
            UpdateTitle();
        }
        return touched;
    }

    private void SyncActiveBaseline(int col, int row, string text)
    {
        if (_activeYColumn != null && col == _activeYColumn.Column.ColumnIndex && !_activeYColumn.IsSubCurve)
        {
            if (CellHelper.TryParseDouble(text, out var y))
            {
                int gi = _rowToGridIndex.TryGetValue(row, out var g) ? g : -1;
                double x = _editing.TryGetValue(row, out var ce) ? ce.X : (gi >= 0 ? gi + 1 : row);
                y = _activeYColumn.IsInteger ? Math.Round(y) : y;
                _editing[row] = (x, y);
                _committed[row] = (x, y);
            }
            else { _editing.Remove(row); _committed.Remove(row); }
        }
    }

    private void ApplyPlottedCellChange(int col, int row, string text)
    {
        int si = _checkedCols.FindIndex(c => c.Column.ColumnIndex == col);
        if (si < 0) return;
        // 子曲线列整格为数组/JSON：按拆分读取值决定该行点的增删。空白/不可解析则删除点；
        // 手工填入合法 JSON 后曲线按拆分值补充显示对应点，空白一律不留点。
        if (_checkedCols[si].IsSubCurve)
        {
            RefreshSubSeriesForColumnRow(col, row);
            SyncSubCurveBaseline(col, row);
            return;
        }
        if (CellHelper.TryParseDouble(text, out var y))
        {
            int xgi = _rowToGridIndex.TryGetValue(row, out var g) ? g : -1;
            double x = xgi >= 0 ? xgi + 1 : row; bool xEditable = false;
            if (_xColumn != null && xgi >= 0 &&
                SubCurveHelper.TryReadValue(_snapshot!, _xColumn, xgi, out var xv))
            {
                x = xv;
                xEditable = true;
            }
            _curve.SetSeriesPoint(si, row, x, y, xEditable, xgi + 1);
        }
        else
        {
            _curve.RemoveSeriesPoint(si, row);
        }
        _curve.Invalidate();
    }

    /// <summary>子曲线列（数组/JSON）手动编辑后，同步当前编辑列的基线值，避免拖点撤销回退到旧值。</summary>
    private void SyncSubCurveBaseline(int col, int row)
    {
        if (_activeYColumn == null || !_activeYColumn.IsSubCurve || col != _activeYColumn.Column.ColumnIndex) return;
        if (TryGetPoint(row, out var sp))
        {
            _editing[row] = (sp.X, sp.Y);
            _committed[row] = (sp.X, sp.Y);
        }
        else { _editing.Remove(row); _committed.Remove(row); }
    }

    private void OnGridSelectionChanged(object? sender, EventArgs e)
    {
        // 由曲线联动表格或处于重建期间时不回写曲线
        if (_updatingGridSelection || _syncFromGrid || _suppressRebuild) return;
        if (_snapshot == null) return;

        var cells = _grid.SelectedCells.Cast<DataGridViewCell>()
            .Where(c => c.RowIndex >= 0 && c.ColumnIndex > 0 && c.ColumnIndex <= _snapshot.ColumnCount)
            .ToList();
        if (cells.Count == 0) return;

        // 取选中区域的“活动列”：优先当前单元格所在列，否则取选中数最多的列
        int tableCol;
        if (_grid.CurrentCell is { ColumnIndex: > 0 } cur && cur.ColumnIndex <= _snapshot.ColumnCount)
            tableCol = cur.ColumnIndex - 1;
        else
            tableCol = cells.GroupBy(c => c.ColumnIndex).OrderByDescending(g => g.Count()).First().Key - 1;

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

    // ---------- 网格：行列结构编辑 ----------
    /// <summary>Excel 式：单击列名（表头）选中整列。</summary>
    private void OnGridColumnHeaderClick(object? sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.Button != MouseButtons.Left) return;
        int col = e.ColumnIndex;
        BeginInvokeSafe(() => SelectGridColumn(col));
    }

    /// <summary>选中整个网格列；期间抑制选中回调，结束后手动同步一次曲线。</summary>
    private void SelectGridColumn(int gridCol)
    {
        if (gridCol < 0 || gridCol >= _grid.Columns.Count) return;
        _updatingGridSelection = true;
        try
        {
            _grid.ClearSelection();
            foreach (DataGridViewRow row in _grid.Rows)
                row.Cells[gridCol].Selected = true;
        }
        finally
        {
            _updatingGridSelection = false;
        }
        OnGridSelectionChanged(_grid, EventArgs.Empty);
    }

    /// <summary>选中 [from, to] 区间内的整行；期间抑制选中回调，结束后手动同步一次曲线。</summary>
    private void SelectGridRows(int from, int to)
    {
        if (_grid.Rows.Count == 0) return;
        int lo = Math.Max(0, Math.Min(from, to));
        int hi = Math.Min(_grid.Rows.Count - 1, Math.Max(from, to));
        if (lo > hi) return;

        _updatingGridSelection = true;
        try
        {
            _grid.ClearSelection();
            for (int gi = lo; gi <= hi; gi++)
            {
                var row = _grid.Rows[gi];
                foreach (DataGridViewColumn col in _grid.Columns)
                    if (col.Visible)
                        row.Cells[col.Index].Selected = true;
            }
        }
        finally
        {
            _updatingGridSelection = false;
        }
        OnGridSelectionChanged(_grid, EventArgs.Empty);
    }

    /// <summary>开始行头拖拽：记录锚点并选中起始行。</summary>
    private void BeginRowHeaderDrag(int row)
    {
        _rowHeaderAnchor = row;
        _rowHeaderLastRow = row;
        _rowHeaderDragging = true;
    }

    /// <summary>行头拖拽中：鼠标悬停到其它行头时，扩展选中区间。</summary>
    private void OnGridMouseMove(object? sender, MouseEventArgs e)
    {
        if (!_rowHeaderDragging) return;
        if (_grid.Rows.Count == 0) return;
        var hit = _grid.HitTest(e.X, e.Y);
        if (hit.Type != DataGridViewHitTestType.RowHeader && hit.Type != DataGridViewHitTestType.Cell) return;
        if (hit.RowIndex < 0) return;
        int row = Math.Clamp(hit.RowIndex, 0, _grid.Rows.Count - 1);
        if (row != _rowHeaderLastRow)
        {
            _rowHeaderLastRow = row;
            SelectGridRows(_rowHeaderAnchor, row);
        }
    }

    /// <summary>结束行头拖拽：重新确认最终选中区间并恢复光标。</summary>
    private void OnGridMouseUp(object? sender, MouseEventArgs e)
    {
        if (!_rowHeaderDragging) return;
        _rowHeaderDragging = false;
        int anchor = _rowHeaderAnchor;
        int lastRow = _rowHeaderLastRow;
        _rowHeaderAnchor = -1;
        _rowHeaderLastRow = -1;
        if (e.Button == MouseButtons.Left)
            BeginInvokeSafe(() => SelectGridRows(anchor, lastRow));
    }

    private void OnGridMouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            var hitL = _grid.HitTest(e.X, e.Y);
            if (hitL.Type == DataGridViewHitTestType.RowHeader && hitL.RowIndex >= 0
                && !IsRowHeaderResizeZone(e.Y, hitL.RowIndex))
                BeginRowHeaderDrag(hitL.RowIndex);
            return;
        }
        if (e.Button != MouseButtons.Right) return;
        var hit = _grid.HitTest(e.X, e.Y);
        _gridTargetRowBlock = hit.RowIndex >= 0 && hit.Type != DataGridViewHitTestType.ColumnHeader
            && IsRowFullySelected(hit.RowIndex);
        _gridTargetSel = _gridTargetRowBlock ? GetSelectedGridRows() : new List<int>();
        _gridTarget = hit.Type switch
        {
            DataGridViewHitTestType.ColumnHeader => (-1, hit.ColumnIndex, GridArea.ColumnHeader),
            DataGridViewHitTestType.RowHeader => (hit.RowIndex, -1, GridArea.RowHeader),
            DataGridViewHitTestType.Cell => (hit.RowIndex, hit.ColumnIndex, GridArea.Cell),
            _ => (-1, -1, GridArea.None)
        };
    }

    /// <summary>光标是否落在行头单元格底部的“行高调节”区域（此时应让行高拖拽工作，而不是选中行）。</summary>
    private bool IsRowHeaderResizeZone(int y, int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= _grid.Rows.Count) return true;
        var rect = _grid.GetRowDisplayRectangle(rowIndex, false);
        return y >= rect.Bottom - 6;
    }

    private void BuildGridContextMenu()
    {
        _gridMenu.Items.Clear();
        var (row, col, area) = _gridTarget;
        bool rowTarget = area != GridArea.ColumnHeader && row >= 0;
        bool colTarget = !(area == GridArea.RowHeader) && col > 0;

        if (rowTarget)
        {
            var deleteRows = ResolveRowsToDelete(row, _gridTargetSel, _gridTargetRowBlock);
            _gridMenu.Items.Add(MakeMenu("在上方插入行", "在选中行上方插入一个空行", () => InsertGridRow(row, false), false));
            _gridMenu.Items.Add(MakeMenu("在下方插入行", "在选中行下方插入一个空行", () => InsertGridRow(row, true), false));
            _gridMenu.Items.Add(BuildCustomRowMenu(row));
            _gridMenu.Items.Add(MakeMenu(
                deleteRows.Count > 1 ? $"删除选中的 {deleteRows.Count} 行" : "删除当前行",
                "删除选中行及其数据",
                () => DeleteGridRows(deleteRows), false));
        }
        if (rowTarget && colTarget)
            _gridMenu.Items.Add(new ToolStripSeparator());
        if (colTarget)
        {
            int physical = col - 1;
            _gridMenu.Items.Add(MakeMenu("在左侧插入列", "在选中列左侧插入一个新列", () => InsertGridColumn(col, false), false));
            _gridMenu.Items.Add(MakeMenu("在右侧插入列", "在选中列右侧插入一个新列", () => InsertGridColumn(col, true), false));
            _gridMenu.Items.Add(MakeMenu("删除当前列", "删除选中列及其数据", () => DeleteGridColumn(col), false));
            _gridMenu.Items.Add(MakeMenu("重命名列", "直接修改当前列表头（字段名）", () => BeginHeaderEdit(col), false));
            _gridMenu.Items.Add(new ToolStripSeparator());
            _gridMenu.Items.Add(BuildAlignMenu(physical));
            _gridMenu.Items.Add(BuildHeaderAlignMenu(physical));
        }
        if (_gridMenu.Items.Count == 0)
            _gridMenu.Items.Add(new ToolStripMenuItem("（无可操作项）") { Enabled = false });
    }

    private void BuildGridToolbarMenu()
    {
        _gridButton.DropDownItems.Clear();
        if (_snapshot == null) return;
        var cur = _grid.CurrentCell;
        int row = _grid.CurrentRow?.Index ?? -1;
        int col = cur?.ColumnIndex ?? -1;
        bool hasRow = row >= 0;
        bool hasCol = col > 0;
        ToolStripMenuItem RowMenu(string text, string tip, Action action)
        {
            var m = MakeMenu(text, tip, action, false);
            m.Enabled = hasRow;
            return m;
        }
        ToolStripMenuItem ColMenu(string text, string tip, Action action)
        {
            var m = MakeMenu(text, tip, action, false);
            m.Enabled = hasCol;
            return m;
        }
        _gridButton.DropDownItems.Add(RowMenu("在上方插入行", "在当前行上方插入一个空行", () => InsertGridRow(row, false)));
        _gridButton.DropDownItems.Add(RowMenu("在下方插入行", "在当前行下方插入一个空行", () => InsertGridRow(row, true)));
        _gridButton.DropDownItems.Add(BuildCustomRowMenu(row, hasRow));
        var deleteRows = ResolveRowsToDelete(row, GetSelectedGridRows(), row >= 0 && IsRowFullySelected(row));
        _gridButton.DropDownItems.Add(RowMenu(
            deleteRows.Count > 1 ? $"删除选中的 {deleteRows.Count} 行" : "删除当前行",
            "删除选中行及其数据",
            () => DeleteGridRows(deleteRows)));
        _gridButton.DropDownItems.Add(new ToolStripSeparator());
        _gridButton.DropDownItems.Add(ColMenu("在左侧插入列", "在当前列左侧插入一个新列", () => InsertGridColumn(col, false)));
        _gridButton.DropDownItems.Add(ColMenu("在右侧插入列", "在当前列右侧插入一个新列", () => InsertGridColumn(col, true)));
        _gridButton.DropDownItems.Add(ColMenu("删除当前列", "删除当前列及其数据", () => DeleteGridColumn(col)));
        _gridButton.DropDownItems.Add(ColMenu("重命名列", "直接修改当前列表头（字段名）", () => BeginHeaderEdit(col)));
        _gridButton.DropDownItems.Add(new ToolStripSeparator());
        if (hasCol)
        {
            _gridButton.DropDownItems.Add(BuildAlignMenu(col - 1));
            _gridButton.DropDownItems.Add(BuildHeaderAlignMenu(col - 1));
        }
    }

    /// <summary>构建“插入自定义行数”子菜单：输入任意数量，一次性插入多行。</summary>
    private ToolStripMenuItem BuildCustomRowMenu(int row, bool enabled = true)
    {
        var menu = new ToolStripMenuItem("插入自定义行数");
        menu.DropDownItems.Add(MakeMenu("在上方插入...", "输入行数并一次性插入多行", () => InsertCustomRows(row, false), false));
        menu.DropDownItems.Add(MakeMenu("在下方插入...", "输入行数并一次性插入多行", () => InsertCustomRows(row, true), false));
        menu.Enabled = enabled && row >= 0;
        return menu;
    }

    /// <summary>构建“内容对齐”子菜单（只改列内容，不影响表头文字）。</summary>
    private ToolStripMenuItem BuildAlignMenu(int physical)
    {
        var menu = new ToolStripMenuItem("内容对齐");
        var cur = physical >= 0 && _snapshot != null && physical < _snapshot.ColumnAlignments.Count
            ? _snapshot.ColumnAlignments[physical]
            : HorizontalAlign.Default;
        menu.DropDownItems.Add(MakeMenu("左对齐", "列内容靠左", () => SetColumnAlignment(physical, HorizontalAlign.Left), cur == HorizontalAlign.Left));
        menu.DropDownItems.Add(MakeMenu("居中", "列内容居中", () => SetColumnAlignment(physical, HorizontalAlign.Center), cur == HorizontalAlign.Center));
        menu.DropDownItems.Add(MakeMenu("右对齐", "列内容靠右", () => SetColumnAlignment(physical, HorizontalAlign.Right), cur == HorizontalAlign.Right));
        return menu;
    }

    /// <summary>构建“表头对齐”子菜单（控制表头文字显示，保存时会写回 Excel）。</summary>
    private ToolStripMenuItem BuildHeaderAlignMenu(int physical)
    {
        var menu = new ToolStripMenuItem("表头对齐");
        var cur = physical >= 0 && _snapshot != null && physical < _snapshot.HeaderAlignments.Count
            ? _snapshot.HeaderAlignments[physical]
            : HorizontalAlign.Default;
        // 表头默认即居中，因此 Default 在菜单中按“居中”显示
        var shown = cur == HorizontalAlign.Default ? HorizontalAlign.Center : cur;
        menu.DropDownItems.Add(MakeMenu("左对齐", "表头文字靠左", () => SetHeaderAlignment(physical, HorizontalAlign.Left), shown == HorizontalAlign.Left));
        menu.DropDownItems.Add(MakeMenu("居中", "表头文字居中", () => SetHeaderAlignment(physical, HorizontalAlign.Center), shown == HorizontalAlign.Center));
        menu.DropDownItems.Add(MakeMenu("右对齐", "表头文字靠右", () => SetHeaderAlignment(physical, HorizontalAlign.Right), shown == HorizontalAlign.Right));
        return menu;
    }

    /// <summary>设置某列的内容对齐方式，并标记待写回 Excel 的整列数据单元格。</summary>
    private void SetColumnAlignment(int physical, HorizontalAlign align)
    {
        if (_snapshot == null || physical < 0 || physical >= _snapshot.ColumnAlignments.Count) return;
        _snapshot.ColumnAlignments[physical] = align;
        _pendingColumnAlign.Add(physical);
        int gridCol = physical + 1;
        if (gridCol >= 0 && gridCol < _grid.Columns.Count)
            _grid.Columns[gridCol].DefaultCellStyle.Alignment = MapAlign(align);
        _grid.Refresh();
        UpdateTitle();
    }

    /// <summary>设置某列表头文字的对齐方式，并标记待写回 Excel。</summary>
    private void SetHeaderAlignment(int physical, HorizontalAlign align)
    {
        if (_snapshot == null || physical < 0 || physical >= _snapshot.Columns.Count) return;
        _snapshot.HeaderAlignments[physical] = align;
        _pendingHeaderAlign.Add(physical);
        int gridCol = physical + 1;
        if (gridCol >= 0 && gridCol < _grid.Columns.Count)
            _grid.Columns[gridCol].HeaderCell.Style.Alignment = MapHeaderAlign(align);
        _grid.Refresh();
        UpdateTitle();
    }

    private void OnGridColumnHeaderDoubleClick(object? sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.ColumnIndex <= 0) return;
        BeginHeaderEdit(e.ColumnIndex);
    }

    /// <summary>在网格行 gridRow 的上方(false)/下方(true)插入空行。</summary>
    private void InsertGridRow(int gridRow, bool below)
    {
        if (_snapshot == null) return;
        int physical;
        if (gridRow >= 0 && gridRow < _snapshot.RowNumbers.Count)
            physical = _snapshot.RowNumbers[gridRow] + (below ? 1 : 0);
        else
            physical = (_snapshot.RowNumbers.Count > 0 ? _snapshot.RowNumbers[^1] : _snapshot.MaxRow) + 1;

        _pendingStructure.Add(new StructuralOp(_snapshot.SheetName, StructuralKind.InsertRow, physical));
        _structureQueued = true;
        RemapRowKeys(physical, +1, false);
        _snapshot.InsertRow(physical);
        RebuildAfterStructure();
    }

    /// <summary>询问用户要插入的行数，并在选中行上方/下方一次性插入多行。</summary>
    private void InsertCustomRows(int gridRow, bool below)
    {
        int count = PromptRowCount();
        if (count <= 0) return;
        InsertGridRows(gridRow, below, count);
    }

    /// <summary>在网格行 gridRow 的上方(false)/下方(true)一次性插入 count 个空行。</summary>
    private void InsertGridRows(int gridRow, bool below, int count)
    {
        if (_snapshot == null || count <= 0) return;
        int physical;
        if (gridRow >= 0 && gridRow < _snapshot.RowNumbers.Count)
            physical = _snapshot.RowNumbers[gridRow] + (below ? 1 : 0);
        else
            physical = (_snapshot.RowNumbers.Count > 0 ? _snapshot.RowNumbers[^1] : _snapshot.MaxRow) + 1;

        for (int i = 0; i < count; i++)
            _pendingStructure.Add(new StructuralOp(_snapshot.SheetName, StructuralKind.InsertRow, physical));
        _structureQueued = true;
        RemapRowKeys(physical, count, false);
        for (int i = 0; i < count; i++)
            _snapshot.InsertRow(physical);
        RebuildAfterStructure();
    }

    /// <summary>弹出“插入自定义行数”对话框，返回输入的行数；取消返回 0。</summary>
    private int PromptRowCount()
    {
        using var f = new Form
        {
            Text = "插入自定义行数",
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterParent,
            ClientSize = new Size(340, 140),
            MaximizeBox = false,
            MinimizeBox = false
        };
        f.Controls.Add(new Label { Text = "请输入要插入的行数（如 100、110）：", AutoSize = true, Location = new Point(18, 22) });
        var num = new NumericUpDown
        {
            DecimalPlaces = 0,
            Minimum = 1,
            Maximum = 200000,
            Value = 1,
            Location = new Point(18, 56),
            Width = 150
        };
        f.Controls.Add(num);
        var ok = new Button { Text = "确定", DialogResult = DialogResult.OK, Location = new Point(150, 98), Width = 80 };
        var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Location = new Point(238, 98), Width = 80 };
        f.Controls.Add(ok);
        f.Controls.Add(cancel);
        f.AcceptButton = ok;
        f.CancelButton = cancel;
        return f.ShowDialog(this) == DialogResult.OK ? (int)num.Value : 0;
    }

    /// <summary>当前选中的网格行索引（去重、升序、排除表头）。</summary>
    private List<int> GetSelectedGridRows()
        => _grid.SelectedCells.Cast<DataGridViewCell>()
            .Where(c => c.RowIndex >= 0)
            .Select(c => c.RowIndex)
            .Distinct()
            .OrderBy(r => r)
            .ToList();

    /// <summary>目标行是否被“整行”选中（行头拖拽得到的满行），用于区分列头/单元格选择。</summary>
    private bool IsRowFullySelected(int gridRow)
    {
        if (gridRow < 0 || gridRow >= _grid.Rows.Count) return false;
        foreach (DataGridViewColumn c in _grid.Columns)
            if (c.Visible && !_grid.Rows[gridRow].Cells[c.Index].Selected)
                return false;
        return true;
    }

    /// <summary>根据目标行与选中信息决定要删除的行：整行选中块优先，否则退化为目标单行。</summary>
    private List<int> ResolveRowsToDelete(int targetRow, IReadOnlyList<int> selected, bool rowBlock)
    {
        if (targetRow >= 0 && rowBlock && selected.Count > 0) return selected.ToList();
        if (targetRow >= 0) return new List<int> { targetRow };
        return selected.ToList();
    }
    
    /// <summary>一次性删除多个网格行（按物理行号降序，保证最低索引始终有效）。</summary>
    private void DeleteGridRows(IReadOnlyList<int> gridRows)
    {
        if (_snapshot == null) return;
        var physicals = gridRows
            .Where(r => r >= 0 && r < _snapshot.RowNumbers.Count)
            .Select(r => _snapshot.RowNumbers[r])
            .Distinct()
            .OrderByDescending(x => x)
            .ToList();
        if (physicals.Count == 0) return;
        foreach (int physical in physicals)
        {
            _pendingStructure.Add(new StructuralOp(_snapshot.SheetName, StructuralKind.DeleteRow, physical));
            _structureQueued = true;
            RemapRowKeys(physical, 0, true);
            _snapshot.DeleteRow(physical);
        }
        RebuildAfterStructure();
    }

    /// <summary>在网格列 gridCol 的左侧(false)/右侧(true)插入新列。</summary>
    private void InsertGridColumn(int gridCol, bool right)
    {
        if (_snapshot == null || gridCol <= 0) return;
        int physical = gridCol - 1 + (right ? 1 : 0);
        physical = Math.Clamp(physical, 0, _snapshot.ColumnCount);
        InsertColumnAtPhysical(physical, DefaultColumnName(physical));
        RebuildAfterStructure(true);
    }

    /// <summary>删除网格列 gridCol。</summary>
    private void DeleteGridColumn(int gridCol)
    {
        if (_snapshot == null || gridCol <= 0 || gridCol > _snapshot.ColumnCount) return;
        int physical = gridCol - 1;
        if (physical < 0 || physical >= _snapshot.Columns.Count) return;
        _pendingStructure.Add(new StructuralOp(_snapshot.SheetName, StructuralKind.DeleteColumn, physical));
        _structureQueued = true;
        RemapColumnKeys(physical, 0, true);
        _snapshot.DeleteColumn(physical);
        RebuildAfterStructure(true);
    }

    /// <summary>在物理列 physical 处插入新列（含结构队列与快照平移），之后由调用方刷新。</summary>
    private void InsertColumnAtPhysical(int physical, string header)
    {
        if (_snapshot == null) return;
        physical = Math.Clamp(physical, 0, _snapshot.ColumnCount);
        _pendingStructure.Add(new StructuralOp(_snapshot.SheetName, StructuralKind.InsertColumn, physical));
        _structureQueued = true;
        RemapColumnKeys(physical, +1, false);
        _snapshot.InsertColumn(physical, header);
    }

    /// <summary>提交列名修改：更新快照并排队写回 Excel 表头。</summary>
    private void CommitColumnRename(int physical, string text)
    {
        if (_snapshot == null || physical < 0 || physical >= _snapshot.Columns.Count) return;
        var meta = _snapshot.Columns[physical];
        if ((meta.HeaderRaw ?? "") == text) return;
        _snapshot.RenameColumn(physical, text);
        _pendingHeaderRename.RemoveAll(p => p.Col == physical);
        _pendingHeaderRename.Add((physical, text));
        RebuildAfterStructure(true);
        _statusLabel.Text = $"已重命名列 {meta.Letter} → {text}";
    }

    private static string DefaultColumnName(int physical)
        => "新列" + CellHelper.ColumnIndexToLetter(physical);

    /// <summary>在当前网格列 gridCol 的表头位置弹出内联编辑框。</summary>
    private void BeginHeaderEdit(int gridCol)
    {
        CommitHeaderEdit();
        if (_snapshot == null || gridCol <= 0 || gridCol >= _grid.Columns.Count) return;
        int physical = gridCol - 1;
        if (physical < 0 || physical >= _snapshot.Columns.Count) return;

        string current = _snapshot.Columns[physical].HeaderRaw ?? "";
        var rect = _grid.GetCellDisplayRectangle(gridCol, -1, false);
        var tb = new TextBox
        {
            Bounds = new Rectangle(rect.X, Math.Max(0, rect.Y), Math.Max(40, rect.Width), Math.Max(18, _grid.ColumnHeadersHeight)),
            Text = current,
            TextAlign = HorizontalAlignment.Center,
            Font = _grid.Font,
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = Color.White,
            Multiline = false
        };
        tb.KeyDown += OnHeaderEditKeyDown;
        tb.Leave += OnHeaderEditLeave;
        _grid.Controls.Add(tb);
        tb.BringToFront();
        _headerEdit = tb;
        _headerEditPhysicalCol = physical;
        tb.Focus();
        tb.SelectAll();
    }

    private void OnHeaderEditKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter) { CommitHeaderEdit(); e.Handled = true; e.SuppressKeyPress = true; }
        else if (e.KeyCode == Keys.Escape) { CancelHeaderEdit(); e.Handled = true; e.SuppressKeyPress = true; }
    }

    private void OnHeaderEditLeave(object? sender, EventArgs e) => CommitHeaderEdit();

    /// <summary>提交内联列名编辑。</summary>
    private void CommitHeaderEdit()
    {
        var tb = _headerEdit;
        if (tb == null) return;
        _headerEdit = null;
        string text = tb.Text.Trim();
        int physical = _headerEditPhysicalCol;
        _headerEditPhysicalCol = -1;
        tb.Dispose();

        if (string.IsNullOrWhiteSpace(text) || physical < 0 || physical >= _snapshot!.Columns.Count) return;
        CommitColumnRename(physical, text);
    }

    private void CancelHeaderEdit()
    {
        var tb = _headerEdit;
        if (tb == null) return;
        _headerEdit = null;
        _headerEditPhysicalCol = -1;
        tb.Dispose();
    }

    // ---------- 网格：剪贴板（Excel 式复制/剪切/粘贴/清除） ----------
    /// <summary>当前选中的数据单元格（排除行号列），按行、列排序。</summary>
    private List<(int Col, int Row, string Text)> SelectedDataCells()
    {
        var list = new List<(int Col, int Row, string Text)>();
        if (_snapshot == null) return list;
        foreach (var cell in _grid.SelectedCells.Cast<DataGridViewCell>())
        {
            if (cell.RowIndex < 0 || cell.ColumnIndex <= 0) continue;
            int col = cell.ColumnIndex - 1;
            int row = _snapshot.RowNumbers[cell.RowIndex];
            list.Add((col, row, cell.Value?.ToString() ?? ""));
        }
        return list.OrderBy(c => c.Row).ThenBy(c => c.Col).ToList();
    }

    private void CopySelection()
    {
        var cells = SelectedDataCells();
        if (cells.Count == 0) return;
        int minR = cells.Min(c => c.Row), maxR = cells.Max(c => c.Row);
        int minC = cells.Min(c => c.Col), maxC = cells.Max(c => c.Col);
        var sb = new StringBuilder();
        for (int r = minR; r <= maxR; r++)
        {
            if (r != minR) sb.Append("\r\n");
            for (int c = minC; c <= maxC; c++)
            {
                if (c != minC) sb.Append('\t');
                sb.Append(cells.FirstOrDefault(x => x.Row == r && x.Col == c).Text ?? "");
            }
        }
        try { Clipboard.SetText(sb.ToString(), TextDataFormat.UnicodeText); _statusLabel.Text = $"已复制 {cells.Count} 个单元格"; }
        catch { _statusLabel.Text = "复制失败：剪贴板暂不可用"; }
    }

    private void CutSelection()
    {
        CopySelection();
        ClearSelection();
        var cells = SelectedDataCells();
        _statusLabel.Text = $"已剪切 {cells.Count} 个单元格";
    }

    private void ClearSelection()
    {
        var cells = SelectedDataCells();
        if (cells.Count == 0) return;
        var touched = ApplyCellEdits(cells.Select(c => (c.Col, c.Row, "")));
        foreach (var (col, row) in touched)
            if (_rowToGridIndex.TryGetValue(row, out var gi) && gi >= 0 && gi < _grid.Rows.Count)
                _grid.Rows[gi].Cells[col + 1].Value = "";
        _statusLabel.Text = $"已清除 {touched.Count} 个单元格";
    }

    private void PasteFromClipboard()
    {
        if (_snapshot == null) return;
        string text;
        try { text = Clipboard.GetText(); } catch { _statusLabel.Text = "读取剪贴板失败"; return; }
        if (string.IsNullOrEmpty(text)) return;
        var block = text.Replace("\r\n", "\n").Replace('\r', '\n')
            .Split('\n')
            .Select(l => l.Split('\t').ToList())
            .ToList();
        while (block.Count > 0 && block[^1].Count == 1 && block[^1][0].Length == 0)
            block.RemoveAt(block.Count - 1);
        if (block.Count == 0) return;

        var anchor = _grid.CurrentCell;
        int anchorGridRow = anchor?.RowIndex ?? (_grid.SelectedCells.Count > 0 ? _grid.SelectedCells[0].RowIndex : -1);
        int anchorGridCol = anchor?.ColumnIndex ?? (_grid.SelectedCells.Count > 0 ? _grid.SelectedCells[0].ColumnIndex : -1);
        if (anchorGridRow < 0 || anchorGridRow >= _snapshot.RowNumbers.Count || anchorGridCol <= 0) return;
        if (anchorGridCol > _snapshot.ColumnCount) return;

        int startCol = anchorGridCol - 1;
        int startRowIndex = anchorGridRow;
        var edits = new List<(int Col, int Row, string Text)>();
        for (int r = 0; r < block.Count; r++)
        {
            int gridRow = startRowIndex + r;
            if (gridRow >= _snapshot.RowNumbers.Count) break;
            int row = _snapshot.RowNumbers[gridRow];
            var line = block[r];
            for (int c = 0; c < line.Count; c++)
            {
                int col = startCol + c;
                if (col >= _snapshot.ColumnCount) break;
                edits.Add((col, row, line[c]));
            }
        }
        var touched = ApplyCellEdits(edits);
        foreach (var (col, row) in touched)
            if (_rowToGridIndex.TryGetValue(row, out var gi) && gi >= 0 && gi < _grid.Rows.Count)
                _grid.Rows[gi].Cells[col + 1].Value = CellText(col, row);
        _statusLabel.Text = $"已粘贴 {touched.Count} 个单元格";
    }

    /// <summary>结构改动后刷新列选择器与曲线/网格。</summary>
    private void RefreshAfterStructure(bool columnsChanged = false)
    {
        if (_snapshot == null) return;
        var opts = SubCurveHelper.BuildOptions(_snapshot!)
            .Where(o => !ShouldHideColumn(o.Column))
            .ToList();
        var checkedKeys = _colsChecked.CheckedItems.Cast<CurveColumnOption>()
            .Select(o => (o.Column.ColumnIndex, o.SubIndex, o.JsonIndex, o.JsonId, o.IsJsonValue))
            .ToHashSet();
        var xb = _xCombo.SelectedItem as CurveColumnOption;
        var xKey = xb == null ? default((int, int, int, string, bool)?) : (xb.Column.ColumnIndex, xb.SubIndex, xb.JsonIndex, xb.JsonId, xb.IsJsonValue);

        _suppressRebuild = true;
        try
        {
            _colsChecked.Items.Clear();
            _xCombo.Items.Clear();
            _xCombo.Items.Add("行号");
            foreach (var o in opts)
            {
                int idx = _colsChecked.Items.Add(o);
                _colsChecked.SetItemChecked(idx, checkedKeys.Contains((o.Column.ColumnIndex, o.SubIndex, o.JsonIndex, o.JsonId, o.IsJsonValue)));
                _xCombo.Items.Add(o);
            }
            int xSel = 0;
            for (int i = 1; i < _xCombo.Items.Count; i++)
                if (_xCombo.Items[i] is CurveColumnOption oo &&
                    xKey.HasValue &&
                    (oo.Column.ColumnIndex, oo.SubIndex, oo.JsonIndex, oo.JsonId, oo.IsJsonValue) == xKey.Value) { xSel = i; break; }
            _xCombo.SelectedIndex = xSel >= 0 ? xSel : 0;
        }
        finally
        {
            _preferredListIndex = -1;
            _preferredListChecked = false;
            _suppressRebuild = false;
        }
        RebuildFromSelection();
        HighlightEditableColumn();
        UpdateTitle();
    }

    /// <summary>仅刷新列名，不丢任何未保存的改动或曲线状态。</summary>
    private void RefreshColumnNames()
    {
        if (_busyLoading) return;
        if (_wb == null || _snapshot == null || string.IsNullOrEmpty(_activeSheet)) return;

        try
        {
            var fresh = _wb.LoadSheet(_activeSheet);
            bool colsChanged = ColumnsChanged(fresh);
            bool hasUnsaved = HasUnsavedChanges();

            if (!hasUnsaved && colsChanged)
            {
                // 无未保存改动且列结构变化：安全重载整表，新列连同数据读取，曲线一并刷新。
                _snapshot = fresh;
                RefreshNumericFlagsFromData(_snapshot);
                RefreshAfterStructure(true);
                _statusLabel.Text = "已刷新列名";
            }
            else
            {
                // 有未保存改动，或列结构未变：只同步列名元信息与新列数据，不重载、不动曲线、不清改动。
                SyncColumnHeaders(fresh);
                RefreshNumericFlagsFromData(_snapshot);
                RebuildColumnOptionsPreserve();
                _statusLabel.Text = colsChanged
                    ? "已刷新列名；当前有未保存改动，如需同步新列数据请先保存后再刷新"
                    : "列名无变化。若新增了列，请先在 Excel 保存后再刷新列名";
            }
        }
        catch (Exception ex)
        {
            _statusLabel.Text = "刷新列名失败（文件可能被占用）：" + ex.Message;
        }
    }

    /// <summary>比较磁盘最新快照与当前快照的列结构（列数或表头文本）是否变化。</summary>
    private bool ColumnsChanged(SheetSnapshot fresh)
    {
        if (fresh.ColumnCount != _snapshot!.ColumnCount) return true;
        for (int c = 0; c < fresh.ColumnCount; c++)
            if (fresh.Columns[c].HeaderRaw != _snapshot.Columns[c].HeaderRaw) return true;
        return false;
    }

    /// <summary>仅同步列名元信息到当前快照；列数增加时在末尾补列（不重载数据）。返回是否发生变化。</summary>
    private bool SyncColumnHeaders(SheetSnapshot fresh)
    {
        var cur = _snapshot!;
        bool changed = false;

        int common = Math.Min(cur.ColumnCount, fresh.ColumnCount);
        for (int c = 0; c < common; c++)
        {
            var src = fresh.Columns[c];
            var dst = cur.Columns[c];
            if (dst.HeaderRaw != src.HeaderRaw)
            {
                dst.HeaderRaw = src.HeaderRaw;
                dst.Name = src.Name;
                dst.Label = src.Label;
                dst.Type = src.Type;
                dst.IsEmpty = src.IsEmpty;
                dst.IsNumericScalar = src.IsNumericScalar;
                dst.IsInteger = src.IsInteger;
                dst.NonEmptyCount = src.NonEmptyCount;
                dst.NumericCount = src.NumericCount;
                dst.TotalRows = src.TotalRows;
                dst.Width = src.Width;
                changed = true;
            }
        }

        // 建物理行号 → 数据行的映射，用于把新增列在磁盘上的数据一并拷进来
        var freshByRow = new Dictionary<int, string?[]>();
        for (int fi = 0; fi < fresh.Grid.Count; fi++)
            freshByRow[fresh.RowNumbers[fi]] = fresh.Grid[fi];

        for (int c = cur.ColumnCount; c < fresh.ColumnCount; c++)
        {
            var meta = fresh.Columns[c];
            for (int gi = 0; gi < cur.Grid.Count; gi++)
            {
                var old = cur.Grid[gi];
                var nrow = new string?[old.Length + 1];
                Array.Copy(old, nrow, old.Length);
                if (freshByRow.TryGetValue(cur.RowNumbers[gi], out var frow) && c < frow.Length)
                    nrow[^1] = frow[c];
                cur.Grid[gi] = nrow;
            }
            cur.Columns.Add(meta);
            cur.ColumnAlignments.Add(HorizontalAlign.Default);
            cur.ColumnCount++;
            changed = true;
        }

        return changed;
    }

    /// <summary>只重建列选择器与 X 轴下拉框并恢复勾选/选中，不触发曲线重建、不清空编辑状态。</summary>
    private void RebuildColumnOptionsPreserve()
    {
        var opts = SubCurveHelper.BuildOptions(_snapshot!)
            .Where(o => !ShouldHideColumn(o.Column))
            .ToList();
        var checkedKeys = _colsChecked.CheckedItems.Cast<CurveColumnOption>()
            .Select(o => (o.Column.ColumnIndex, o.SubIndex, o.JsonIndex, o.JsonId, o.IsJsonValue))
            .ToHashSet();
        var xb = _xCombo.SelectedItem as CurveColumnOption;
        var xKey = xb == null
            ? default((int, int, int, string, bool)?)
            : (xb.Column.ColumnIndex, xb.SubIndex, xb.JsonIndex, xb.JsonId, xb.IsJsonValue);

        _suppressRebuild = true;
        try
        {
            _colsChecked.Items.Clear();
            _xCombo.Items.Clear();
            _xCombo.Items.Add("行号");
            foreach (var o in opts)
            {
                int idx = _colsChecked.Items.Add(o);
                _colsChecked.SetItemChecked(idx, checkedKeys.Contains((o.Column.ColumnIndex, o.SubIndex, o.JsonIndex, o.JsonId, o.IsJsonValue)));
                _xCombo.Items.Add(o);
            }
            int xSel = 0;
            for (int i = 1; i < _xCombo.Items.Count; i++)
                if (_xCombo.Items[i] is CurveColumnOption oo &&
                    xKey.HasValue &&
                    (oo.Column.ColumnIndex, oo.SubIndex, oo.JsonIndex, oo.JsonId, oo.IsJsonValue) == xKey.Value)
                { xSel = i; break; }
            _xCombo.SelectedIndex = xSel >= 0 ? xSel : 0;
        }
        finally
        {
            _preferredListIndex = -1;
            _preferredListChecked = false;
            _suppressRebuild = false;
        }
    }

    /// <summary>
    /// 根据当前快照的实际网格数据，重新判定每列是否为标量数值列 / 整数列。
    /// 用于“刷新列名”时把新增（可能未保存到磁盘）的数值列识别出来。
    /// 数组列、JSON 列、显式非数值类型列保持不变（由子曲线逻辑处理）。
    /// </summary>
    private void RefreshNumericFlagsFromData(SheetSnapshot snap)
    {
        for (int c = 0; c < snap.Columns.Count; c++)
        {
            var col = snap.Columns[c];
            if (SubCurveHelper.IsArrayCol(col) || SubCurveHelper.IsJsonCol(col)) continue;

            bool explicitNumeric = HeaderParser.IsExplicitNumericScalar(col.HeaderRaw, col.Type);
            bool explicitInteger = HeaderParser.IsExplicitInteger(col.Type);
            int nonEmpty = 0, numeric = 0;
            bool allInteger = true;
            for (int gi = 0; gi < snap.Grid.Count; gi++)
            {
                var row = snap.Grid[gi];
                if (col.ColumnIndex < 0 || col.ColumnIndex >= row.Length) continue;
                var raw = row[col.ColumnIndex];
                if (string.IsNullOrWhiteSpace(raw)) continue;
                nonEmpty++;
                if (CellHelper.TryParseDouble(raw, out var d))
                {
                    numeric++;
                    if (d != Math.Truncate(d)) allInteger = false;
                }
            }

            bool isNumeric = explicitNumeric
                || (string.IsNullOrEmpty(col.Type) && nonEmpty > 0 && (double)numeric / nonEmpty >= 0.7);
            bool isInteger = explicitInteger
                || (string.IsNullOrEmpty(col.Type) && allInteger && isNumeric);

            col.IsNumericScalar = isNumeric;
            col.IsInteger = isInteger;
            col.NonEmptyCount = nonEmpty;
            col.NumericCount = numeric;
            col.TotalRows = snap.DataRowCount;
        }
    }

    /// <summary>把结构改动后的网格重建延迟到消息循环，避免在控件事件回调中改动网格。</summary>
    private void RebuildAfterStructure(bool columnsChanged = false)
    {
        if (IsDisposed || !IsHandleCreated) return;
        BeginInvoke(() => RefreshAfterStructure(columnsChanged));
    }

    /// <summary>行插入/删除后，平移 dirty 集合中的物理行号。delta=+1 插入，删除时用 delete=true。</summary>
    private void RemapRowKeys(int fromPhysical, int delta, bool delete)
    {
        var newDirty = new HashSet<(int, int)>();
        foreach (var (c, r) in _dirtyCells)
        {
            if (delete) { if (r == fromPhysical) continue; if (r > fromPhysical) newDirty.Add((c, r - 1)); else newDirty.Add((c, r)); }
            else newDirty.Add((c, r >= fromPhysical ? r + delta : r));
        }
        _dirtyCells.Clear();
        _dirtyCells.UnionWith(newDirty);

        var newSub = new Dictionary<(int, int), string>();
        foreach (var kv in _subEditOldText)
        {
            var (c, r) = kv.Key;
            if (delete) { if (r == fromPhysical) continue; if (r > fromPhysical) newSub[(c, r - 1)] = kv.Value; else newSub[(c, r)] = kv.Value; }
            else newSub[(c, r >= fromPhysical ? r + delta : r)] = kv.Value;
        }
        _subEditOldText.Clear();
        foreach (var kv in newSub) _subEditOldText[kv.Key] = kv.Value;
    }

    /// <summary>列插入/删除后，平移 dirty 集合中的物理列号。delta=+1 插入，删除时用 delete=true。</summary>
    private void RemapColumnKeys(int fromPhysical, int delta, bool delete)
    {
        var newDirty = new HashSet<(int, int)>();
        foreach (var (c, r) in _dirtyCells)
        {
            if (delete) { if (c == fromPhysical) continue; if (c > fromPhysical) newDirty.Add((c - 1, r)); else newDirty.Add((c, r)); }
            else newDirty.Add((c >= fromPhysical ? c + delta : c, r));
        }
        _dirtyCells.Clear();
        _dirtyCells.UnionWith(newDirty);

        var newSub = new Dictionary<(int, int), string>();
        foreach (var kv in _subEditOldText)
        {
            var (c, r) = kv.Key;
            if (delete) { if (c == fromPhysical) continue; if (c > fromPhysical) newSub[(c - 1, r)] = kv.Value; else newSub[(c, r)] = kv.Value; }
            else newSub[(c >= fromPhysical ? c + delta : c, r)] = kv.Value;
        }
        _subEditOldText.Clear();
        foreach (var kv in newSub) _subEditOldText[kv.Key] = kv.Value;

        // 列增删也会平移之前记录的表头重命名目标列
        if (delete)
        {
            _pendingHeaderRename.RemoveAll(p => p.Col == fromPhysical);
            for (int i = 0; i < _pendingHeaderRename.Count; i++)
                if (_pendingHeaderRename[i].Col > fromPhysical)
                    _pendingHeaderRename[i] = (_pendingHeaderRename[i].Col - 1, _pendingHeaderRename[i].Text);
        }
        else
        {
            for (int i = 0; i < _pendingHeaderRename.Count; i++)
                if (_pendingHeaderRename[i].Col >= fromPhysical)
                    _pendingHeaderRename[i] = (_pendingHeaderRename[i].Col + delta, _pendingHeaderRename[i].Text);
        }

        // 列增删也平移待写回的表头对齐目标列
        var newHeaderAlign = new HashSet<int>();
        foreach (var c in _pendingHeaderAlign)
        {
            if (delete) { if (c == fromPhysical) continue; newHeaderAlign.Add(c > fromPhysical ? c - 1 : c); }
            else newHeaderAlign.Add(c >= fromPhysical ? c + delta : c);
        }
        _pendingHeaderAlign.Clear();
        _pendingHeaderAlign.UnionWith(newHeaderAlign);

        var newColAlign = new HashSet<int>();
        foreach (var c in _pendingColumnAlign)
        {
            if (delete) { if (c == fromPhysical) continue; newColAlign.Add(c > fromPhysical ? c - 1 : c); }
            else newColAlign.Add(c >= fromPhysical ? c + delta : c);
        }
        _pendingColumnAlign.Clear();
        _pendingColumnAlign.UnionWith(newColAlign);
    }

    // ---------- 统计 ----------
    private void UpdateStats()
    {
        _statLabel.Text = Statistics.Summarize(_curve.Points.Select(p => p.Y)).AsText();
    }

    // ---------- 批量 ----------
    private void OnApplyValue() { double v = (double)_valUpDown.Value; _curve.ApplyToSelected(p => (p.X, v)); }

    /// <summary>
    /// 填充整列空白：扫描当前编辑列的每个连续空白段，用左右相邻的有值单元格连成直线，
    /// 对空白格做线性插值填入。与"高级拟合 直线 + 仅选中点(含空) + 过首尾两点"的思路一致，
    /// 只是这里遍历整列的所有空白段，每段以左右有值格为首尾两点。
    /// </summary>
    private void FillBlanks()
    {
        if (_activeYColumn == null || _snapshot == null) return;
        var snap = _snapshot;
        var opt = _activeYColumn;

        // 单元格拆分列（数组/JSON）：点由单元格内容拆分而来，空白必须手工在单元格填入，曲线操作不允许自动补充空白。
        if (opt.IsSubCurve)
        {
            _statusLabel.Text = "该列为单元格拆分列（数组/JSON），不允许补充空白单元格，请手工在单元格中填入内容";
            return;
        }

        // 收集每个数据行的 (行号, X, 是否有Y值, Y值)。X 无法读取的行直接跳过（无法定位）。
        var rows = new List<(int Row, double X, bool HasValue, double Y)>(snap.RowNumbers.Count);
        for (int gi = 0; gi < snap.RowNumbers.Count; gi++)
        {
            int row = snap.RowNumbers[gi];
            double x = gi + 1;
            if (_xColumn != null)
            {
                if (!SubCurveHelper.TryReadValue(snap, _xColumn, gi, out var xv)) continue;
                x = xv;
            }
            bool hasValue = SubCurveHelper.TryReadValue(snap, opt, gi, out var y);
            rows.Add((row, x, hasValue, y));
        }
        if (rows.Count == 0) return;

        var anchors = rows.Where(r => r.HasValue).ToList();
        if (anchors.Count < 2)
        {
            _statusLabel.Text = "该列至少需要 2 个有值单元格才能填充空白";
            return;
        }

        var writes = new List<(int Row, double X, double Y)>();

        // 用一段直线（通过左右/前后两个有值格）为空白行取落点。before=true 表示填充前段（该直线左端之前的空白）。
        void FillWithLine((int Row, double X, bool HasValue, double Y) pa, (int Row, double X, bool HasValue, double Y) pb, bool before)
        {
            if (Math.Abs(pb.X - pa.X) < 1e-12) return;
            foreach (var r in rows)
            {
                bool inRange = before ? r.Row < pa.Row : r.Row > pb.Row;
                if (!inRange || r.HasValue) continue;
                double y = pa.Y + (pb.Y - pa.Y) * (r.X - pa.X) / (pb.X - pa.X);
                if (!double.IsFinite(y)) continue;
                writes.Add((r.Row, r.X, y));
            }
        }

        // 前段空白：用开头两个有值格外推（首个有值格作右端、第二个作右端锚点外推）
        FillWithLine(anchors[0], anchors[1], before: true);

        // 内部空白段：相邻两个有值格之间
        for (int i = 0; i < anchors.Count - 1; i++)
        {
            var a = anchors[i];
            var b = anchors[i + 1];
            if (Math.Abs(b.X - a.X) < 1e-12) continue;
            foreach (var r in rows)
                if (r.Row > a.Row && r.Row < b.Row && !r.HasValue)
                {
                    double y = a.Y + (b.Y - a.Y) * (r.X - a.X) / (b.X - a.X);
                    if (!double.IsFinite(y)) continue;
                    writes.Add((r.Row, r.X, y));
                }
        }

        // 后段空白：用末尾两个有值格外推
        FillWithLine(anchors[^2], anchors[^1], before: false);

        if (writes.Count == 0)
        {
            _statusLabel.Text = "没有可填充的空白单元格";
            return;
        }
        ApplyColumnFill(writes);
        _statusLabel.Text = $"已填充 {writes.Count} 个空白单元格";
    }

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

    /// <summary>平滑“选中的点”：与整列平滑不同，只对选中点序列做移动平均，非选中点不动。</summary>
    private void BatchSmoothSelected()
    {
        var sel = _curve.GetSelectedPoints();
        if (sel.Count < 3)
        {
            _statusLabel.Text = "请至少选择 3 个点再平滑";
            return;
        }
        var sm = CurveMath.MovingAverage(sel.Select(p => p.Y).ToList(), 3);
        int i = 0;
        _curve.ApplyToSelected(p => (p.X, sm[i++]));
        _statusLabel.Text = $"已平滑 {sel.Count} 个选中点";
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
            if (HasSubCurveForColumn(col))
                RefreshSubSeriesForColumnRow(col, row);
            else
                ApplyPlottedCellChange(col, row, val);
            if (_activeYColumn != null && col == _activeYColumn.Column.ColumnIndex)
            {
                if (_activeYColumn.IsSubCurve)
                {
                    if (TryGetPoint(row, out var sp))
                    {
                        _editing[row] = (sp.X, sp.Y);
                        _committed[row] = (sp.X, sp.Y);
                    }
                    else { _editing.Remove(row); _committed.Remove(row); }
                }
                else if (CellHelper.TryParseDouble(val, out var y))
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

    private bool HasSubCurveForColumn(int col)
        => _checkedCols.Any(o => o.Column.ColumnIndex == col && o.IsSubCurve);

    private void RefreshSubSeriesForColumnRow(int col, int row)
    {
        for (int si = 0; si < _checkedCols.Count; si++)
        {
            var opt = _checkedCols[si];
            if (opt.Column.ColumnIndex != col || !opt.IsSubCurve) continue;

            double? y = null;
            if (_rowToGridIndex.TryGetValue(row, out var gi) &&
                gi >= 0 && gi < _snapshot!.Grid.Count &&
                SubCurveHelper.TryReadValue(_snapshot, opt, gi, out var v))
                y = v;

            if (y.HasValue)
            {
                double x = gi >= 0 ? gi + 1 : row;
                bool xEditable = false;
                var existing = _series[si].Points.FirstOrDefault(pp => pp.RowNumber == row);
                if (existing != null)
                {
                    x = existing.X;
                    xEditable = existing.XEditable;
                }
                else if (_xColumn != null && gi >= 0 &&
                         SubCurveHelper.TryReadValue(_snapshot, _xColumn, gi, out var xv))
                {
                    x = xv;
                    xEditable = true;
                }
                _curve.SetSeriesPoint(si, row, x, y.Value, xEditable, gi + 1);
            }
            else
                _curve.RemoveSeriesPoint(si, row);
        }
        _curve.Invalidate();
    }

    // ---------- 保存/刷新/导出 ----------
    private void OnSave()
    {
        if (!CommitPending(out var err) && !string.IsNullOrEmpty(err))
            MessageBox.Show(this, err, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }

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
        if (_busyLoading) return;
        if (_wb == null || string.IsNullOrEmpty(_activeSheet)) return;
        try
        {
            _wb.RefreshMeta();
            _snapshot = _wb.LoadSheet(_activeSheet);
            _pendingStructure.Clear();
            _pendingHeaderRename.Clear();
            _structureQueued = false;
            _dirtyCells.Clear();
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
        _gridBottomItem.Checked = !atRight;
        _gridRightItem.Checked = atRight;
        _chartGridSplit.Orientation = atRight ? Orientation.Vertical : Orientation.Horizontal;
        try
        {
            if (atRight) _chartGridSplit.SplitterDistance = Math.Max(200, (int)(_chartGridSplit.Width * 0.72));
            else _chartGridSplit.SplitterDistance = Math.Max(140, (int)(_chartGridSplit.Height * 0.58));
        }
        catch { }
        SaveSettings();
    }

    // ---------- 表格最大化 / 还原 ----------
    private void ToggleGridMaximize()
    {
        if (_gridMaximized) RestoreGrid();
        else MaximizeGrid();
    }

    /// <summary>
    /// 把表格（含底部工作表标签）从下方分隔面板取出，铺满主窗口中间区域，
    /// 隐藏曲线、左右功能面板，便于专注编辑数据。
    /// </summary>
    private void MaximizeGrid()
    {
        if (_gridMaximized) return;
        _gridMaximized = true;

        // 从分隔面板中剥离，改挂到主窗体，并置于工具栏下方、状态栏上方
        _chartGridSplit.Panel2.Controls.Remove(_gridPane);
        Controls.Add(_gridPane);
        _gridPane.Dock = DockStyle.Fill;
        // 置为 z-order 最前，使工具栏、状态栏先按边停靠占据上下，表格再填充剩余中间区域，
        // 避免表格顶部表头被工具栏遮住。
        Controls.SetChildIndex(_gridPane, 0);

        // 隐藏曲线与两侧功能面板，只留表格编辑区
        _chartGridSplit.Visible = false;

        UpdateLayoutMenu();
    }

    /// <summary>把表格放回原来的位置（下方/右侧分隔面板）。</summary>
    private void RestoreGrid()
    {
        if (!_gridMaximized) return;
        _gridMaximized = false;

        Controls.Remove(_gridPane);
        _chartGridSplit.Visible = true;
        _chartGridSplit.Panel2.Controls.Add(_gridPane);
        _gridPane.Dock = DockStyle.Fill;

        UpdateLayoutMenu();
    }

    private void UpdateLayoutMenu()
    {
        _gridMaxItem.Text = _gridMaximized ? "还原表格" : "表格最大化";
        _gridMaxItem.ToolTipText = _gridMaximized
            ? "把表格还原到原来的位置（F11）"
            : "把表格铺满主窗口，便于编辑数据（F11）";
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
        if (_curve.SelectedCount >= 2)
            _menu.Items.Add(BuildFitMenu());
        var adv = new ToolStripMenuItem("高级拟合...");
        adv.Click += (s, e) => OpenAdvancedFit();
        _menu.Items.Add(adv);
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
        AddBatch(menu, "曲线平滑", () => { EnsureActiveHitSeries(); BatchSmoothSelected(); });
        return menu;
    }

    private static void AddBatch(ToolStripMenuItem parent, string text, Action action)
    {
        var item = new ToolStripMenuItem(text);
        item.Click += (s, e) => action();
        parent.DropDownItems.Add(item);
    }

    private FitModel SelectedFitModel()
    {
        var models = System.Enum.GetValues<FitModel>();
        int idx = Math.Max(0, _fitTypeCombo.SelectedIndex);
        return models[Math.Min(idx, models.Length - 1)];
    }

    private int FitDegree => decimal.ToInt32(_fitDegree.Value);

    private ToolStripMenuItem BuildFitMenu()
    {
        var fit = new ToolStripMenuItem("拟合选中点");
        foreach (var m in System.Enum.GetValues<FitModel>())
            fit.DropDownItems.Add(AddFitItem(CurveFit.LabelOf(m), m));
        return fit;
    }

    private ToolStripMenuItem AddFitItem(string text, FitModel model)
    {
        var item = new ToolStripMenuItem(text);
        item.Click += (s, e) => ApplyFitModel(model);
        return item;
    }

    private void ApplyFitModel(FitModel model)
    {
        var models = System.Enum.GetValues<FitModel>();
        int idx = Array.IndexOf(models, model);
        if (idx >= 0) _fitTypeCombo.SelectedIndex = idx;
        OnFitApply();
    }

    private void OnFitPreview()
    {
        var pts = _curve.GetSelectedPoints();
        if (pts.Count < 2)
        {
            _statusLabel.Text = "请先选择至少 2 个点，再预览拟合";
            return;
        }
        var r = CurveFit.Fit(pts.Select(p => (p.X, p.Y)).ToArray(), SelectedFitModel(), FitDegree);
        if (r.Evaluate == null)
        {
            MessageBox.Show(this, r.Error, "拟合失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        SetFitOverlay(r);
        ShowFitInfo(r);
        _statusLabel.Text = $"拟合预览：{r.Label}  R²={r.R2:0.###}  RMSE={r.RMSE:0.###}";
    }

    private void OnFitApply()
    {
        var pts = _curve.GetSelectedPoints();
        if (pts.Count < 2)
        {
            _statusLabel.Text = "请先选择至少 2 个点，再应用拟合";
            return;
        }
        var r = CurveFit.Fit(pts.Select(p => (p.X, p.Y)).ToArray(), SelectedFitModel(), FitDegree);
        if (r.Evaluate == null)
        {
            MessageBox.Show(this, r.Error, "拟合失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        var ev = r.Evaluate;
        _curve.ApplyToSelected(p => (p.X, ev(p.X)));
        SetFitOverlay(r);
        ShowFitInfo(r);
        _statusLabel.Text = $"已按{r.Label}应用：{r.Formula}  R²={r.R2:0.###}  RMSE={r.RMSE:0.###}";
    }

    private void ClearFitPreview()
    {
        _curve.ClearOverlay();
        _fitInfo.Text = "";
        _statusLabel.Text = "已清除拟合预览";
    }

    private void SetFitOverlay(FitResult r)
    {
        var pts = _curve.GetSelectedPoints();
        if (pts.Count == 0) return;
        double xMin = pts.Min(p => p.X), xMax = pts.Max(p => p.X);
        if (Math.Abs(xMax - xMin) < 1e-12) xMax = xMin + 1;
        var overlay = new List<(double X, double Y)>();
        const int n = 200;
        for (int i = 0; i <= n; i++)
        {
            double x = xMin + (xMax - xMin) * i / n;
            double y;
            try { y = r.Evaluate!(x); }
            catch { y = 0; }
            if (!double.IsFinite(y)) y = 0;
            overlay.Add((x, y));
        }
        _curve.OverlayPoints = overlay;
        _curve.Invalidate();
    }

    private void ShowFitInfo(FitResult r)
    {
        _fitInfo.Text = $"{r.Label}{Environment.NewLine}{r.Formula}{Environment.NewLine}R²={r.R2:0.###}  RMSE={r.RMSE:0.###}";
    }

    private void OpenAdvancedFit()
    {
        if (_fitDialog == null || _fitDialog.IsDisposed)
        {
            _fitDialog = new FitDialog(_curve) { Owner = this };
            _fitDialog.ColumnRowsProvider = GetColumnRowsForFill;
            _fitDialog.ColumnFillApplier = ApplyColumnFill;
        }
        if (!_fitDialog.Visible)
        {
            _fitDialog.Show();
            _fitDialog.BringToFront();
        }
    }

    /// <summary>返回当前编辑列整行（含空单元格）的物理行号与 X 值：默认 X=行号，设了 X 列则取该列值。</summary>
    private IReadOnlyList<(int Row, double X)> GetColumnRowsForFill()
    {
        var list = new List<(int, double)>();
        if (_snapshot == null || _activeYColumn == null) return list;
        for (int gi = 0; gi < _snapshot.RowNumbers.Count; gi++)
        {
            int row = _snapshot.RowNumbers[gi];
            double x = gi + 1;
            if (_xColumn != null)
            {
                if (!SubCurveHelper.TryReadValue(_snapshot, _xColumn, gi, out var xv)) continue;
                x = xv;
            }
            list.Add((row, x));
        }
        return list;
    }

    /// <summary>
    /// 把“整列每格”的拟合结果写满当前编辑列（含原本为空的单元格），
    /// 同步曲线点、快照与撤销栈，一次性提交一条撤销记录。
    /// </summary>
    private void ApplyColumnFill(IReadOnlyList<(int Row, double X, double Y)> writes)
    {
        if (_activeYColumn == null || _snapshot == null) return;
        var opt = _activeYColumn;
        int col = opt.Column.ColumnIndex;
        var changed = new List<(int, int, string, string)>();
        int skippedBlank = 0;

        foreach (var (row, x, y) in writes)
        {
            if (col < 0 || col >= _snapshot.ColumnCount) continue;
            if (!_rowToGridIndex.TryGetValue(row, out var gi) || gi < 0 || gi >= _snapshot.Grid.Count) continue;

            // 单元格拆分列（数组/JSON）：该行无拆分值时跳过，不生成曲线点也不向空白单元格写入
            if (opt.IsSubCurve && !SubCurveHelper.TryReadValue(_snapshot, opt, gi, out _))
            {
                skippedBlank++;
                continue;
            }

            // 曲线点始终按拟合值更新，且与单元格存储值保持一致（整数列取整），
            // 否则仅改写“取整后发生变化”的行会让折线新旧值混杂，看起来就不是直线。
            double chartY = !opt.IsSubCurve && opt.IsInteger ? Math.Round(y) : y;
            _curve.SetSeriesPoint(_curve.ActiveSeriesIndex, row, x, chartY, _xColumn != null, gi + 1);

            string oldText = CellText(col, row);
            string newText = opt.IsSubCurve
                ? SubCurveHelper.SetValue(oldText, opt, y)
                : FormatCellValue(y, opt.IsInteger);
            if (oldText == newText) continue;

            UpdateSnapshotCell(row, col, newText);
            _dirtyCells.Add((col, row));
            _editing[row] = (x, chartY);
            _committed[row] = (x, chartY);
            changed.Add((col, row, oldText, newText));
        }

        if (skippedBlank > 0)
            _statusLabel.Text = changed.Count > 0
                ? $"已更新 {changed.Count} 个点，跳过 {skippedBlank} 个空白（数组/JSON 列不允许自动补充）"
                : $"已跳过 {skippedBlank} 个空白单元格（数组/JSON 列不允许自动补充，请手工填入）";
        if (changed.Count == 0) return;
        _undo.Add(new EditCmd(changed));
        _redo.Clear();
        _pendingUndoRows.Clear();
        UpdateGridCells(writes.Select(w => w.Row).ToArray());
        UpdateStats();
        UpdateTitle();
        _curve.Invalidate();
        if (_autoSaveCheck.Checked) { _autoSaveTimer.Stop(); _autoSaveTimer.Start(); }
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
        if (_wb != null && HasUnsavedChanges())
        {
            var r = MessageBox.Show(this, "有未保存的改动，是否保存？", "确认", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (r == DialogResult.Cancel) { e.Cancel = true; return; }
            if (r == DialogResult.Yes)
            {
                if (!CommitPending(out var err))
                {
                    MessageBox.Show(this, string.IsNullOrEmpty(err) ? "保存失败，请检查文件是否被占用。" : err,
                        "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
            }
        }
        _watcher?.Dispose();
        _autoSaveTimer.Stop();
        _reloadTimer.Stop();
        _fitDialog?.Dispose();
        CloseSplash();
        SaveSettings();
        _wb?.Dispose();
        base.OnFormClosing(e);
    }

    private void UpdateTitle()
    {
        var dirty = HasUnsavedChanges() ? "（有未保存改动）" : "";
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

    /// <summary>表头对齐映射：未设置或居中时默认居中，与全局表头样式保持一致。</summary>
    private static DataGridViewContentAlignment MapHeaderAlign(HorizontalAlign a) => a switch
    {
        HorizontalAlign.Left => DataGridViewContentAlignment.MiddleLeft,
        HorizontalAlign.Right => DataGridViewContentAlignment.MiddleRight,
        _ => DataGridViewContentAlignment.MiddleCenter
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
        // ---------- 打开 / 文件 ----------
        var pick = new ToolStripMenuItem("选择文件...");
        pick.Click += (s, e) => OnOpen();
        _openMenu.DropDownItems.Add(pick);

        var close = new ToolStripMenuItem("关闭当前文件");
        close.Enabled = _wb != null;
        close.Click += (s, e) => CloseCurrentFile();
        _openMenu.DropDownItems.Add(close);

        var openExcel = new ToolStripMenuItem("用Excel打开当前文件");
        openExcel.Enabled = _wb != null && File.Exists(_wb.Path);
        openExcel.Click += (s, e) => OpenCurrentInExcel();
        _openMenu.DropDownItems.Add(openExcel);

        _openMenu.DropDownItems.Add(new ToolStripSeparator());

        // ---------- 保存 ----------
        var save = new ToolStripMenuItem("保存");
        save.ToolTipText = "把当前改动写回 Excel 文件（Ctrl+S）";
        save.Click += (s, e) => OnSave();
        _openMenu.DropDownItems.Add(save);

        var saveAs = new ToolStripMenuItem("另存为");
        saveAs.ToolTipText = "复制一份并另存为新文件";
        saveAs.Click += (s, e) => OnSaveAs();
        _openMenu.DropDownItems.Add(saveAs);

        _openMenu.DropDownItems.Add(_autoSaveCheck);

        var exportPng = new ToolStripMenuItem("导出PNG");
        exportPng.ToolTipText = "把当前曲线图导出为 PNG 图片";
        exportPng.Click += (s, e) => OnExport();
        _openMenu.DropDownItems.Add(exportPng);

        _openMenu.DropDownItems.Add(new ToolStripSeparator());

        // ---------- 最近文件 ----------
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
    {
        string name = Path.GetFileName(path);
        if (name.Contains("~$", StringComparison.Ordinal) ||
            name.Contains("111", StringComparison.Ordinal))
            return false;
        return name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
               name.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase);
    }

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
        Tip(_randUpDown, "随机扰动幅度（±，配合“随机扰动(选中点)”按钮）");
        Tip(_statLabel, "当前编辑列的统计信息（最大/平均/总和/标准差等）");
    }

    private void Tip(Control c, string t) => _tip.SetToolTip(c, t);

    private void OnSelectionChangedUi()
    {
        EnsureActiveSeriesSynced();
        _selInfo.Text = "选中: " + _curve.SelectedCount;
        _curve.ClearOverlay();
        _fitInfo.Text = "";
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
                                // 目标滚动列不能是冻结列（行号列 Frozen=true），否则抛 InvalidOperationException
                                int frozenCount = 0;
                                foreach (DataGridViewColumn c in _grid.Columns)
                                {
                                    if (!c.Frozen) break;
                                    frozenCount++;
                                }
                                int scrollCol = Math.Max(frozenCount, col - 1);
                                if (scrollCol > _grid.Columns.Count - 1)
                                    scrollCol = _grid.Columns.Count - 1;
                                _grid.FirstDisplayedScrollingColumnIndex = scrollCol;
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
