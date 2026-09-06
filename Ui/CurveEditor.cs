using System.Drawing.Drawing2D;
using GameCurve.Models;
using GameCurve.Services;

namespace GameCurve.Ui;

/// <summary>一条可显示的曲线序列（含该列颜色、是否可编辑）。</summary>
public sealed class CurveSeriesView
{
    public string Name { get; set; } = "";
    public Color Color { get; set; } = Color.SteelBlue;
    public List<CurvePoint> Points { get; } = new();
    public bool IsEditable { get; set; }
    public bool Visible { get; set; } = true;
}

/// <summary>
/// 可交互曲线编辑器：网格坐标、多曲线叠加、鼠标拖点、框选/多选/全选、键盘微调、缩放平移。
/// </summary>
public sealed class CurveEditor : Control
{
    private readonly List<CurveSeriesView> _series = new();
    private readonly HashSet<int> _selected = new();

    // 护眼暗色画布底
    private static readonly Color CanvasBack = Color.FromArgb(30, 34, 42);

    private double _xMin, _xMax, _yMin, _yMax;
    private bool _hasView;

    private enum DragMode { None, Move, Marquee, Pan }
    private DragMode _drag = DragMode.None;
    private int _grabIndex = -1;
    private PointF _dragStartWorld;
    private readonly Dictionary<int, (double X, double Y)> _dragInitial = new();
    private bool _dragged;
    private bool _spaceDown;
    private Point _mouseDown;
    private int _hoverIndex = -1;
    private bool _additiveSelect;
    private Rectangle _marquee;

    /// <summary>叠加绘制的拟合曲线（按 X 升序的折线）。</summary>
    public List<(double X, double Y)>? OverlayPoints { get; set; }

    public event Action<IReadOnlyList<int>>? PointsChanged;
    public event Action? EditCommitted;
    public event Action? SelectionChanged;
    public event Action<string>? HoverChanged;

    // 临时诊断
    public bool ShowSpline { get; set; } = true;
    public bool ShowGrid { get; set; } = true;
    public bool ShowPoints { get; set; } = true;
    public bool ShowLabels { get; set; } = true;
    public double KeyboardStep { get; set; } = 1;
    public string XAxisLabel { get; set; } = "";
    public string YAxisLabel { get; set; } = "";

    public static readonly Color[] Palette =
    {
        Color.FromArgb(49, 110, 244), // 蓝
        Color.FromArgb(242, 130, 40), // 橙
        Color.FromArgb(46, 160, 120), // 绿
        Color.FromArgb(220, 70, 90),  // 红
        Color.FromArgb(150, 90, 210), // 紫
        Color.FromArgb(24, 160, 180), // 青
        Color.FromArgb(220, 170, 40), // 金
        Color.FromArgb(110, 130, 160) // 灰蓝
    };

    public CurveEditor()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                 ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                 ControlStyles.Selectable, true);
        BackColor = Color.FromArgb(38, 42, 48);
        TabStop = true;
    }

    public int ActiveSeriesIndex { get; private set; } = -1;
    public int SeriesCount => _series.Count;
    public CurveSeriesView? ActiveSeries => ActiveSeriesIndex >= 0 && ActiveSeriesIndex < _series.Count ? _series[ActiveSeriesIndex] : null;

    public IReadOnlyList<CurvePoint> Points => ActiveSeries != null ? ActiveSeries.Points : new List<CurvePoint>();
    public IReadOnlyList<int> SelectedRows => _selected.Select(i => ActiveSeries?.Points[i].RowNumber ?? 0).Where(r => r != 0).ToArray();
    public int SelectedCount => _selected.Count;

    public int HoverRow => (_hoverIndex >= 0 && ActiveSeries != null && _hoverIndex < ActiveSeries.Points.Count) ? ActiveSeries.Points[_hoverIndex].RowNumber : 0;

    public void SetSeries(IReadOnlyList<CurveSeriesView> series, int activeIndex)
    {
        _series.Clear();
        foreach (var s in series) _series.Add(s);
        ActiveSeriesIndex = activeIndex;
        _selected.Clear();
        _hasView = false;
        AutoFitView();
        Invalidate();
    }

    public void SetActiveSeries(int index, bool autoFit = true)
    {
        ActiveSeriesIndex = index;
        _selected.Clear();
        if (autoFit)
        {
            _hasView = false;
            AutoFitView();
        }
        Invalidate();
        SelectionChanged?.Invoke();
    }

    public void ProgrammaticSetYByRow(int row, double y)
    {
        var list = ActiveSeries?.Points;
        if (list == null) return;
        for (int i = 0; i < list.Count; i++)
        {
            if (list[i].RowNumber == row)
            {
                list[i].Y = y;
                Invalidate();
                return;
            }
        }
    }

    /// <summary>设置某条曲线指定行的点；不存在则新增。</summary>
    public void SetSeriesPoint(int seriesIndex, int row, double x, double y, bool xEditable)
    {
        if (seriesIndex < 0 || seriesIndex >= _series.Count) return;
        var list = _series[seriesIndex].Points;
        for (int i = 0; i < list.Count; i++)
            if (list[i].RowNumber == row)
            {
                list[i].X = x; list[i].Y = y; list[i].XEditable = xEditable;
                Invalidate();
                return;
            }
        // 新增点按 X 升序插入，避免“接到曲线尾部”
        var pt = new CurvePoint(x, y, row, xEditable);
        int insertAt = list.Count;
        for (int i = 0; i < list.Count; i++)
            if (list[i].X > x) { insertAt = i; break; }
        list.Insert(insertAt, pt);
        Invalidate();
    }

    /// <summary>移除某条曲线指定行的点（单元格非数值时）。</summary>
    public void RemoveSeriesPoint(int seriesIndex, int row)
    {
        if (seriesIndex < 0 || seriesIndex >= _series.Count) return;
        var list = _series[seriesIndex].Points;
        for (int i = 0; i < list.Count; i++)
            if (list[i].RowNumber == row)
            {
                list.RemoveAt(i);
                Invalidate();
                return;
            }
    }

    public void ApplyToSelected(Func<CurvePoint, (double X, double Y)> transform, bool commit = true)
    {
        var list = ActiveSeries?.Points;
        if (list == null) return;
        var affected = new List<int>();
        foreach (var idx in _selected.OrderBy(i => i))
        {
            var p = list[idx];
            var (nx, ny) = transform(p);
            p.X = nx;
            p.Y = ny;
            affected.Add(p.RowNumber);
        }
        if (affected.Count > 0)
        {
            Invalidate();
            PointsChanged?.Invoke(affected);
            if (commit) EditCommitted?.Invoke();
        }
    }

    public void ApplyToAll(Func<CurvePoint, (double X, double Y)> transform, bool commit = true)
    {
        var list = ActiveSeries?.Points;
        if (list == null || list.Count == 0) return;
        var affected = new List<int>(list.Count);
        for (int i = 0; i < list.Count; i++)
        {
            var p = list[i];
            var (nx, ny) = transform(p);
            p.X = nx;
            p.Y = ny;
            affected.Add(p.RowNumber);
        }
        Invalidate();
        PointsChanged?.Invoke(affected);
        if (commit) EditCommitted?.Invoke();
    }

    /// <summary>返回当前编辑列上被选中点的数据（按列表索引升序）。</summary>
    public IReadOnlyList<CurvePoint> GetSelectedPoints()
    {
        var list = ActiveSeries?.Points;
        if (list == null) return Array.Empty<CurvePoint>();
        return _selected.OrderBy(i => i).Select(i => list[i]).ToArray();
    }

    public void ClearSelection()
    {
        if (_selected.Count == 0) return;
        _selected.Clear();
        Invalidate();
        SelectionChanged?.Invoke();
    }

    /// <summary>清除叠加的拟合曲线。</summary>
    public void ClearOverlay()
    {
        OverlayPoints = null;
        Invalidate();
    }

    public void SelectAll()
    {
        var list = ActiveSeries?.Points;
        if (list == null) return;
        _selected.Clear();
        for (int i = 0; i < list.Count; i++) _selected.Add(i);
        Invalidate();
        SelectionChanged?.Invoke();
    }

    /// <summary>按物理行号选中当前编辑列上对应的点。</summary>
    public void SelectPointByRow(int row)
    {
        _selected.Clear();
        var list = ActiveSeries?.Points;
        if (list != null)
            for (int i = 0; i < list.Count; i++)
                if (list[i].RowNumber == row) { _selected.Add(i); break; }
        Invalidate();
        SelectionChanged?.Invoke();
    }

    /// <summary>按物理行号批量选中当前编辑列的点。</summary>
    public void SelectPointsByRows(IEnumerable<int> rows, bool additive = false)
    {
        var list = ActiveSeries?.Points;
        if (list == null) return;
        if (!additive) _selected.Clear();
        foreach (var row in rows)
            for (int i = 0; i < list.Count; i++)
                if (list[i].RowNumber == row) { _selected.Add(i); break; }
        Invalidate();
        SelectionChanged?.Invoke();
    }

    // ---------- 坐标变换 ----------
    private const int MarginL = 66, MarginR = 26, MarginT = 16, MarginB = 44;
    private int PlotW => Math.Max(2, ClientSize.Width - MarginL - MarginR);
    private int PlotH => Math.Max(2, ClientSize.Height - MarginT - MarginB);
    private Rectangle PlotRect => new(MarginL, MarginT, PlotW, PlotH);

    private PointF WorldToScreen(double x, double y)
    {
        double dx = (x - _xMin) / Math.Max(1e-12, _xMax - _xMin) * PlotW;
        double dy = (y - _yMin) / Math.Max(1e-12, _yMax - _yMin);
        double sx = MarginL + dx;
        double sy = MarginT + (1 - dy) * PlotH;
        // NaN/Infinity 兜底
        if (!double.IsFinite(sx)) sx = MarginL;
        if (!double.IsFinite(sy)) sy = MarginT;
        // 远离视图的点钳制到可控范围，避免 GDI 将超大坐标转 int 时溢出
        const double Limit = 1e6;
        sx = Math.Clamp(sx, -Limit, Limit);
        sy = Math.Clamp(sy, -Limit, Limit);
        return new PointF((float)sx, (float)sy);
    }

    private (double X, double Y) ScreenToWorld(PointF p)
    {
        double x = _xMin + (p.X - MarginL) / Math.Max(1, PlotW) * (_xMax - _xMin);
        double y = _yMin + (1 - (p.Y - MarginT) / Math.Max(1, PlotH)) * (_yMax - _yMin);
        return (x, y);
    }

    public void AutoFitView()
    {
        var pts = _series.Where(s => s.Visible).SelectMany(s => s.Points).ToList();
        if (pts.Count == 0)
        {
            _xMin = 0; _xMax = 100; _yMin = 0; _yMax = 100;
            _hasView = true;
            Invalidate();
            return;
        }
        double xMin = pts.Min(p => p.X), xMax = pts.Max(p => p.X);
        double yMin = pts.Min(p => p.Y), yMax = pts.Max(p => p.Y);
        if (xMax <= xMin) { double c = xMin == 0 ? 1 : Math.Abs(xMin) * 0.1; xMin -= c; xMax += c; }
        if (yMax <= yMin) { double c = yMin == 0 ? 1 : Math.Abs(yMin) * 0.1; yMin -= c; yMax += c; }
        _xMin = xMin - (xMax - xMin) * 0.03; _xMax = xMax + (xMax - xMin) * 0.03;
        _yMin = yMin - (yMax - yMin) * 0.08; _yMax = yMax + (yMax - yMin) * 0.08;
        _hasView = true;
        Invalidate();
    }

    public void SaveBitmap(string path)
    {
        using var bmp = new Bitmap(ClientSize.Width, ClientSize.Height);
        DrawToBitmap(bmp, new Rectangle(0, 0, ClientSize.Width, ClientSize.Height));
        bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
    }

    // ---------- 曲线可见性 / 命中 ----------
    public bool IsSeriesVisible(int index) => index >= 0 && index < _series.Count && _series[index].Visible;
    public string GetSeriesName(int index) => index >= 0 && index < _series.Count ? _series[index].Name : "";
    public void SetSeriesVisible(int index, bool visible)
    {
        if (index < 0 || index >= _series.Count) return;
        _series[index].Visible = visible;
        Invalidate();
    }

    /// <summary>命中某条曲线（其数据点在光标附近），返回曲线索引；无命中返回 -1。</summary>
    public int HitTestAnySeries(Point location)
    {
        for (int si = 0; si < _series.Count; si++)
        {
            var ser = _series[si];
            if (!ser.Visible) continue;
            foreach (var p in ser.Points)
            {
                var sp = WorldToScreen(p.X, p.Y);
                double d = Math.Sqrt((sp.X - location.X) * (sp.X - location.X) + (sp.Y - location.Y) * (sp.Y - location.Y));
                if (d < 14) return si;
            }
        }
        return -1;
    }

    /// <summary>命中任意可见曲线，返回曲线索引和点索引。</summary>
    private bool HitTestAnySeriesDetailed(Point location, out int seriesIndex, out int pointIndex)
    {
        seriesIndex = -1;
        pointIndex = -1;
        double bestDist = 12;
        for (int si = 0; si < _series.Count; si++)
        {
            var ser = _series[si];
            if (!ser.Visible) continue;
            for (int pi = 0; pi < ser.Points.Count; pi++)
            {
                var sp = WorldToScreen(ser.Points[pi].X, ser.Points[pi].Y);
                double d = Math.Sqrt((sp.X - location.X) * (sp.X - location.X) + (sp.Y - location.Y) * (sp.Y - location.Y));
                if (d < bestDist)
                {
                    bestDist = d;
                    seriesIndex = si;
                    pointIndex = pi;
                }
            }
        }
        return seriesIndex >= 0;
    }

    // ---------- 绘制 ----------
    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

        g.Clear(BackColor);
        if (!_hasView) AutoFitView();

        // 绘图区域背景
        var plot = PlotRect;
        using (var bg = new SolidBrush(CanvasBack))
            g.FillRectangle(bg, plot);

        if (_series.Count == 0)
        {
            using var f = new Font("Microsoft YaHei UI", 11f);
            TextRenderer.DrawText(g, "请选择工作表与数值列", f, ClientRectangle, Color.Gray,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            return;
        }

        DrawGridAndAxes(g, plot);
        var oldClip = g.Clip;
        g.SetClip(plot); // 放大/平移后裁剪到绘图区，避免曲线跑到坐标轴外
        for (int si = 0; si < _series.Count; si++)
            DrawSeries(g, plot, _series[si], si == ActiveSeriesIndex);
        if (_drag == DragMode.Marquee)
            DrawMarquee(g);
        DrawOverlay(g);
        g.Clip = oldClip;

        // 悬停信息
        if (_hoverIndex >= 0 && ActiveSeries != null && _hoverIndex < ActiveSeries.Points.Count)
            DrawHover(g, plot, ActiveSeries.Points[_hoverIndex]);
    }

    private void DrawGridAndAxes(Graphics g, Rectangle plot)
    {
        var major = Color.FromArgb(70, 78, 88);
        var minor = Color.FromArgb(52, 60, 68);
        var axisColor = Color.FromArgb(150, 160, 170);
        var textColor = Color.FromArgb(185, 195, 205);
        using var majorPen = new Pen(major, 1f);
        using var minorPen = new Pen(minor, 1f);
        using var axisPen = new Pen(axisColor, 1.2f);
        using var font = new Font("Microsoft YaHei UI", 8f);
        using var labelBrush = new SolidBrush(textColor);

        var xTicks = BuildTicks(_xMin, _xMax, 7);
        var yTicks = BuildTicks(_yMin, _yMax, 6);

        if (ShowGrid)
        {
            foreach (var v in xTicks.Ticks)
            {
                var p = WorldToScreen(v, 0);
                g.DrawLine(minorPen, p.X, plot.Top, p.X, plot.Bottom);
            }
            foreach (var v in yTicks.Ticks)
            {
                var p = WorldToScreen(0, v);
                g.DrawLine(minorPen, plot.Left, p.Y, plot.Right, p.Y);
            }
        }

        // 边框
        g.DrawRectangle(axisPen, plot);

        if (ShowLabels)
        {
            foreach (var v in xTicks.Ticks)
            {
                var p = WorldToScreen(v, 0);
                if (p.X < plot.Left - 2 || p.X > plot.Right + 2) continue;
                g.DrawString(FormatTick(v), font, labelBrush, p.X - 20, plot.Bottom + 4);
            }
            foreach (var v in yTicks.Ticks)
            {
                var p = WorldToScreen(0, v);
                if (p.Y < plot.Top - 2 || p.Y > plot.Bottom + 2) continue;
                g.DrawString(FormatTick(v), font, labelBrush, plot.Left - 52, p.Y - 8);
            }

            // 轴标题
            if (!string.IsNullOrEmpty(XAxisLabel))
            {
                var sz = g.MeasureString(XAxisLabel, font);
                g.DrawString(XAxisLabel, font, labelBrush, plot.Left + plot.Width / 2f - sz.Width / 2, plot.Bottom + 22);
            }
            if (!string.IsNullOrEmpty(YAxisLabel))
            {
                var sz = g.MeasureString(YAxisLabel, font);
                var old = g.Transform;
                g.TranslateTransform(10, plot.Top + plot.Height / 2f + sz.Width / 2);
                g.RotateTransform(-90);
                g.DrawString(YAxisLabel, font, labelBrush, 0, 0);
                g.Transform = old;
            }
        }
    }

    private void DrawSeries(Graphics g, Rectangle plot, CurveSeriesView series, bool editable)
    {
        if (!series.Visible) return;
        if (series.Points.Count == 0) return;
        var pts = series.Points;

        if (ShowSpline && pts.Count >= 2)
        {
            using var pen = new Pen(series.Color, editable ? 2.4f : 1.8f);
            pen.LineJoin = LineJoin.Round;
            if (pts.Count > 400)
            {
                // 大数据量直接连点，避免样条插值开销
                var screen = pts.Select(p => WorldToScreen(p.X, p.Y)).ToArray();
                g.DrawLines(pen, screen);
            }
            else
            {
                int segs = pts.Count > 120 ? 4 : 22;
                var bez = CurveMath.CatmullRom(pts, segs);
                var screen = bez.Select(b => WorldToScreen(b.X, b.Y)).ToArray();
                using var path = new GraphicsPath();
                if (screen.Length > 0)
                {
                    path.AddLines(screen);
                    g.DrawPath(pen, path);
                }
            }
        }

        if (ShowPoints)
        {
            for (int i = 0; i < pts.Count; i++)
            {
                var sp = WorldToScreen(pts[i].X, pts[i].Y);
                if (sp.X < plot.Left - 10 || sp.X > plot.Right + 10 || sp.Y < plot.Top - 10 || sp.Y > plot.Bottom + 10) continue;
                bool sel = editable && _selected.Contains(i);
                bool hover = editable && _hoverIndex == i;
                float r = sel ? 6.5f : (hover ? 5.5f : 4.2f);
                using var fill = new SolidBrush(sel ? Color.FromArgb(255, 150, 0) : series.Color);
                g.FillEllipse(fill, sp.X - r, sp.Y - r, r * 2, r * 2);
                if (sel)
                {
                    using var ring = new Pen(Color.White, 1.6f);
                    g.DrawEllipse(ring, sp.X - r - 1.5f, sp.Y - r - 1.5f, (r + 1.5f) * 2, (r + 1.5f) * 2);
                }
            }
        }
    }

    private void DrawMarquee(Graphics g)
    {
        using var pen = new Pen(Color.FromArgb(49, 110, 244), 1f) { DashStyle = DashStyle.Dash };
        using var brush = new SolidBrush(Color.FromArgb(40, 49, 110, 244));
        var rect = Normalize(_marquee);
        g.FillRectangle(brush, rect);
        g.DrawRectangle(pen, rect);
    }

    private void DrawOverlay(Graphics g)
    {
        if (OverlayPoints == null || OverlayPoints.Count < 2) return;
        var pts = OverlayPoints.Select(p => WorldToScreen(p.X, p.Y)).ToArray();
        using var pen = new Pen(Color.FromArgb(120, 210, 255), 1.6f) { DashStyle = DashStyle.Dash };
        g.DrawLines(pen, pts);
    }

    private void DrawHover(Graphics g, Rectangle plot, CurvePoint p)
    {
        var sp = WorldToScreen(p.X, p.Y);
        string txt = $"行:{p.RowNumber}  X:{p.X:0.###}  Y:{p.Y:0.###}";
        using var font = new Font("Microsoft YaHei UI", 8f);
        var sz = g.MeasureString(txt, font);
        int bx = (int)Math.Min(Math.Max(sp.X + 12, MarginL), ClientSize.Width - sz.Width - 8);
        int by = (int)Math.Max(4, sp.Y - sz.Height - 6);
        var rect = new Rectangle(bx, by, (int)sz.Width + 10, (int)sz.Height + 6);
        using var bg = new SolidBrush(Color.FromArgb(235, 245, 250, 255));
        g.FillRectangle(bg, rect);
        g.DrawRectangle(Pens.Gray, rect);
        using var b = new SolidBrush(Color.FromArgb(40, 46, 54));
        g.DrawString(txt, font, b, bx + 5, by + 2);
    }

    // ---------- 交互 ----------
    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        Focus();
        _mouseDown = e.Location;

        if (e.Button == MouseButtons.Middle)
        {
            _drag = DragMode.Pan;
            _lastPan = e.X;
            _lastPanY = e.Y;
            Capture = true;
            Cursor = Cursors.SizeAll;
            return;
        }
        if (e.Button != MouseButtons.Left) return;

        // 按住空格时，鼠标按在画布任意位置都拖拽所有选中点（不用精准点中）
        if (_spaceDown && _selected.Count > 0 && ActiveSeries != null)
        {
            _grabIndex = _selected.First();
            _drag = DragMode.Move;
            _dragged = false;
            _dragInitial.Clear();
            var list = ActiveSeries.Points;
            for (int i = 0; i < list.Count; i++) _dragInitial[i] = (list[i].X, list[i].Y);
            _dragStartWorld = new PointF(e.X, e.Y);
            Invalidate();
            return;
        }

        // 点选时命中任意可见曲线；命中其它曲线时先切换为当前编辑列
        bool hitAny = HitTestAnySeriesDetailed(e.Location, out var hitSeries, out var hitPoint);
        if (hitAny && hitSeries != ActiveSeriesIndex)
            SetActiveSeries(hitSeries, autoFit: false);
        int hit = hitAny ? hitPoint : -1;
        bool ctrl = (ModifierKeys & Keys.Control) != 0;
        if (hit >= 0)
        {
            if (ctrl)
            {
                if (!_selected.Add(hit)) _selected.Remove(hit);
                Invalidate();
                SelectionChanged?.Invoke();
                return;
            }
            if (!_selected.Contains(hit))
            {
                _selected.Clear();
                _selected.Add(hit);
                SelectionChanged?.Invoke();
            }
            _grabIndex = hit;
            _drag = DragMode.Move;
            _dragged = false;
            _dragInitial.Clear();
            var list = ActiveSeries!.Points;
            for (int i = 0; i < list.Count; i++) _dragInitial[i] = (list[i].X, list[i].Y);
            var sp = WorldToScreen(ActiveSeries!.Points[hit].X, ActiveSeries.Points[hit].Y);
            _dragStartWorld = new PointF(sp.X, sp.Y);
            Invalidate();
        }
        else
        {
            _additiveSelect = ctrl;
            if (!ctrl)
            {
                _selected.Clear();
                SelectionChanged?.Invoke();
            }
            _drag = DragMode.Marquee;
            _marquee = new Rectangle(e.Location, Size.Empty);
        }
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        if (_drag == DragMode.Move && ActiveSeries != null)
        {
            float dx = e.X - _dragStartWorld.X;
            float dy = e.Y - _dragStartWorld.Y;
            if (Math.Abs(dx) > 0.5 || Math.Abs(dy) > 0.5) _dragged = true;
            if (_dragged)
            {
                var wx = dx / Math.Max(1, PlotW) * (_xMax - _xMin);
                var wy = -dy / Math.Max(1, PlotH) * (_yMax - _yMin);
                var list = ActiveSeries.Points;
                var affected = new List<int>();
                foreach (var idx in _selected)
                {
                    var p = list[idx];
                    var init = _dragInitial[idx];
                    if (p.XEditable) p.X = init.X + wx;
                    p.Y = init.Y + wy;
                    affected.Add(p.RowNumber);
                }
                Invalidate();
                PointsChanged?.Invoke(affected);
            }
        }
        else if (_drag == DragMode.Marquee)
        {
            _marquee = new Rectangle(_mouseDown.X, _mouseDown.Y, e.X - _mouseDown.X, e.Y - _mouseDown.Y);
            Invalidate();
        }
        else if (_drag == DragMode.Pan)
        {
            double wx = (e.X - _lastPan) / (double)Math.Max(1, PlotW) * (_xMax - _xMin);
            double wy = (e.Y - _lastPanY) / (double)Math.Max(1, PlotH) * (_yMax - _yMin);
            _xMin -= wx; _xMax -= wx;
            _yMin += wy; _yMax += wy;
            _lastPan = e.X; _lastPanY = e.Y;
            Invalidate();
        }
        else
        {
            UpdateHover(e.Location);
        }
    }

    private int _lastPan, _lastPanY;

    protected override void OnMouseUp(MouseEventArgs e)
    {
        base.OnMouseUp(e);
        Capture = false;
        if (_drag == DragMode.Pan) Cursor = null;
        if (_drag == DragMode.Move && _dragged)
        {
            EditCommitted?.Invoke();
        }
        else if (_drag == DragMode.Marquee)
        {
            bool moved = Math.Abs(e.X - _mouseDown.X) > 4 || Math.Abs(e.Y - _mouseDown.Y) > 4;
            if (!moved)
            {
                if (!_additiveSelect) _selected.Clear();
            }
            else
            {
                var rect = Normalize(_marquee);
                int bestSeries = -1;
                var bestHits = new List<int>();
                for (int si = 0; si < _series.Count; si++)
                {
                    var ser = _series[si];
                    if (!ser.Visible) continue;
                    var hits = new List<int>();
                    for (int i = 0; i < ser.Points.Count; i++)
                    {
                        var sp = WorldToScreen(ser.Points[i].X, ser.Points[i].Y);
                        if (rect.Contains((int)sp.X, (int)sp.Y))
                            hits.Add(i);
                    }
                    if (hits.Count > bestHits.Count)
                    {
                        bestHits = hits;
                        bestSeries = si;
                    }
                }

                if (bestSeries >= 0 && bestHits.Count > 0)
                {
                    if (bestSeries != ActiveSeriesIndex) SetActiveSeries(bestSeries, autoFit: false);
                    if (!_additiveSelect) _selected.Clear();
                    foreach (var i in bestHits) _selected.Add(i);
                    SelectionChanged?.Invoke();
                }
            }
            Invalidate();
            SelectionChanged?.Invoke();
        }
        _drag = DragMode.None;
        _dragged = false;
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        base.OnMouseWheel(e);
        if (_series.Count == 0) return;
        var before = ScreenToWorld(e.Location);
        double factor = e.Delta > 0 ? 0.86 : 1.16;
        double xr = _xMax - _xMin, yr = _yMax - _yMin;
        _xMin = before.X - (before.X - _xMin) * factor;
        _xMax = before.X + (_xMax - before.X) * factor;
        _yMin = before.Y - (before.Y - _yMin) * factor;
        _yMax = before.Y + (_yMax - before.Y) * factor;
        Invalidate();
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        base.OnMouseLeave(e);
        if (_hoverIndex >= 0)
        {
            _hoverIndex = -1;
            HoverChanged?.Invoke("");
            Invalidate();
        }
    }

    protected override void OnMouseEnter(EventArgs e)
    {
        base.OnMouseEnter(e);
        // 鼠标进入画布即让画布获得焦点，表格等其他控件自然失去焦点
        if (!ContainsFocus) Focus();
    }

    protected override void OnDoubleClick(EventArgs e)
    {
        base.OnDoubleClick(e);
        AutoFitView();
    }

    protected override bool IsInputKey(Keys keyData)
    {
        switch (keyData)
        {
            case Keys.Up: case Keys.Down: case Keys.Left: case Keys.Right:
            case Keys.Home: case Keys.End: case Keys.PageUp: case Keys.PageDown:
            case Keys.Space:
                return true;
        }
        return base.IsInputKey(keyData);
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (e.KeyCode == Keys.Space) _spaceDown = true;
        if (e.Control && e.KeyCode == Keys.A)
        {
            SelectAll();
            e.Handled = true;
            return;
        }
        if (e.KeyCode == Keys.Escape)
        {
            ClearSelection();
            e.Handled = true;
            return;
        }

        double step = KeyboardStep;
        if ((e.Modifiers & Keys.Shift) != 0) step *= 10;
        if ((e.Modifiers & Keys.Control) != 0) step /= 10;

        bool changed = false;
        var list = ActiveSeries?.Points;
        if (list == null || _selected.Count == 0) return;
        var affected = new List<int>();
        foreach (var idx in _selected)
        {
            var p = list[idx];
            switch (e.KeyCode)
            {
                case Keys.Up: p.Y += step; changed = true; break;
                case Keys.Down: p.Y -= step; changed = true; break;
                case Keys.Right: if (p.XEditable) p.X += step; changed = true; break;
                case Keys.Left: if (p.XEditable) p.X -= step; changed = true; break;
            }
            affected.Add(p.RowNumber);
        }
        if (changed)
        {
            Invalidate();
            PointsChanged?.Invoke(affected);
            EditCommitted?.Invoke();
        }
        e.Handled = changed;
    }

    protected override void OnKeyUp(KeyEventArgs e)
    {
        base.OnKeyUp(e);
        if (e.KeyCode == Keys.Space) _spaceDown = false;
    }

    private void UpdateHover(Point location)
    {
        int hit = HitTest(location);
        if (hit != _hoverIndex)
        {
            _hoverIndex = hit;
            if (hit >= 0)
            {
                var p = ActiveSeries!.Points[hit];
                HoverChanged?.Invoke($"行:{p.RowNumber}  X:{p.X:0.###}  Y:{p.Y:0.###}");
            }
            else
            {
                HoverChanged?.Invoke("");
            }
            Invalidate();
        }
    }

    private int HitTest(Point location)
    {
        var list = ActiveSeries?.Points;
        if (list == null) return -1;
        int best = -1; double bestDist = 12;
        for (int i = 0; i < list.Count; i++)
        {
            var sp = WorldToScreen(list[i].X, list[i].Y);
            double d = Math.Sqrt((sp.X - location.X) * (sp.X - location.X) + (sp.Y - location.Y) * (sp.Y - location.Y));
            if (d < bestDist) { bestDist = d; best = i; }
        }
        return best;
    }

    private static Rectangle Normalize(Rectangle r)
    {
        int x = Math.Min(r.X, r.Right), y = Math.Min(r.Y, r.Bottom);
        int w = Math.Abs(r.Width), h = Math.Abs(r.Height);
        return new Rectangle(x, y, w, h);
    }

    private static (double Min, double Max, List<double> Ticks) BuildTicks(double min, double max, int target)
    {
        var ticks = new List<double>();
        if (double.IsNaN(min) || double.IsNaN(max)) return (0, 1, ticks);
        if (max <= min) { double c = min == 0 ? 1 : Math.Abs(min) * 0.1; min -= c; max += c; }
        double range = max - min;
        double rough = range / target;
        double mag = Math.Pow(10, Math.Floor(Math.Log10(rough)));
        double norm = rough / mag;
        double step = mag * (norm < 1.5 ? 1 : norm < 3.5 ? 2 : norm < 7.5 ? 5 : 10);
        double start = Math.Floor(min / step) * step;
        for (double v = start; v <= max + step * 0.001; v += step)
            ticks.Add(v);
        return (min, max, ticks);
    }

    private static string FormatTick(double v)
    {
        if (v == 0) return "0";
        if (Math.Abs(v) >= 1e6 || Math.Abs(v) < 1e-4) return v.ToString("0.###E+0", System.Globalization.CultureInfo.InvariantCulture);
        if (v == Math.Truncate(v)) return v.ToString("0", System.Globalization.CultureInfo.InvariantCulture);
        return v.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
    }
}
