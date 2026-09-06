using GameCurve.Models;
using GameCurve.Services;

namespace GameCurve.Ui;

/// <summary>
/// 高级拟合对话框：支持任意参数固定/覆盖、多项式正则化、拟合范围、
/// 吸附或按比例混合、锚定首尾点，并实时在主图上预览拟合曲线。
/// </summary>
public sealed class FitDialog : Form
{
    private readonly CurveEditor _editor;

    private readonly ComboBox _typeCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly NumericUpDown _degreeNum = new() { DecimalPlaces = 0, Minimum = 2, Maximum = CurveFit.MaxPolynomialDegree, Value = 2, Width = 70 };
    private readonly TrackBar _smoothBar = new() { Minimum = 0, Maximum = 100, Width = 250, TickFrequency = 25 };
    private readonly Label _smoothLabel = new() { Text = "平滑度: 0", AutoSize = true };
    private readonly RadioButton _selRadio = new() { Text = "仅选中点", AutoSize = true, Checked = true };
    private readonly RadioButton _allRadio = new() { Text = "整列全部点", AutoSize = true };
    private readonly RadioButton _snapRadio = new() { Text = "直接吸附", AutoSize = true, Checked = true };
    private readonly RadioButton _blendRadio = new() { Text = "按比例混合", AutoSize = true };
    private readonly NumericUpDown _blendPct = new() { DecimalPlaces = 0, Minimum = 0, Maximum = 100, Value = 50, Width = 60 };
    private readonly CheckBox _anchorCheck = new() { Text = "保持首尾点不动", AutoSize = true };
    private readonly TableLayoutPanel _paramTable = new() { AutoSize = false, ColumnCount = 3, Dock = DockStyle.Top };
    private readonly Label _infoLabel = new() { AutoSize = false, Height = 92, Font = new Font("Microsoft YaHei UI", 8f), ForeColor = Color.FromArgb(60, 66, 74) };

    private readonly List<TextBox> _paramBoxes = new();

    public FitDialog(CurveEditor editor)
    {
        _editor = editor;
        Text = "高级拟合控制";
        FormBorderStyle = FormBorderStyle.SizableToolWindow;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(430, 690);
        MaximizeBox = false;
        MinimizeBox = false;
        Font = new Font("Microsoft YaHei UI", 9f);
        BuildUi();
        RebuildParams();
        FormClosing += (s, e) => _editor.ClearOverlay();
    }

    private void BuildUi()
    {
        int y = 12;
        Controls.Add(Section("拟合类型", 12, y)); y += 24;

        _typeCombo.SetBounds(12, y, 190, 26);
        Controls.Add(_typeCombo);
        Controls.Add(new Label { Text = "多项式次数:", AutoSize = true, Location = new Point(214, y + 4) });
        _degreeNum.SetBounds(300, y, 70, 26);
        Controls.Add(_degreeNum);
        y += 34;

        Controls.Add(Section("高阶多项式正则化", 12, y)); y += 24;
        _smoothBar.SetBounds(12, y, 250, 30);
        Controls.Add(_smoothBar);
        _smoothLabel.SetBounds(270, y + 4, 120, 20);
        Controls.Add(_smoothLabel);
        y += 40;

        Controls.Add(Section("拟合范围", 12, y)); y += 24;
        _selRadio.SetBounds(12, y, 110, 22); Controls.Add(_selRadio);
        _allRadio.SetBounds(130, y, 130, 22); Controls.Add(_allRadio);
        y += 30;

        Controls.Add(Section("应用方式", 12, y)); y += 24;
        _snapRadio.SetBounds(12, y, 100, 22); Controls.Add(_snapRadio);
        _blendRadio.SetBounds(120, y, 120, 22); Controls.Add(_blendRadio);
        _blendPct.SetBounds(250, y, 60, 26); _blendPct.Enabled = false; Controls.Add(_blendPct);
        Controls.Add(new Label { Text = "% 向拟合线移动", AutoSize = true, Location = new Point(316, y + 4) });
        y += 30;

        _anchorCheck.SetBounds(12, y, 160, 22); Controls.Add(_anchorCheck);
        y += 30;

        Controls.Add(Section("参数（留空=自动，填值=固定）", 12, y)); y += 24;
        _paramTable.Dock = DockStyle.Top;
        _paramTable.AutoSize = true;
        _paramTable.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        _paramTable.Width = 404;
        _paramTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 76));
        _paramTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170));
        _paramTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        var paramScroll = new Panel { AutoScroll = true, Location = new Point(12, y), Size = new Size(404, 200) };
        paramScroll.Controls.Add(_paramTable);
        Controls.Add(paramScroll);
        y += 208;

        _infoLabel.SetBounds(12, y, 404, 92);
        Controls.Add(_infoLabel);
        y += 100;

        var preview = MakeButton("预览", OnPreview, 88);
        var apply = MakeButton("应用", OnApply, 88);
        var close = MakeButton("关闭", Close, 88);
        preview.SetBounds(12, y, 88, 30);
        apply.SetBounds(106, y, 88, 30);
        close.SetBounds(200, y, 88, 30);
        Controls.Add(preview);
        Controls.Add(apply);
        Controls.Add(close);

        _typeCombo.Items.AddRange(Enum.GetValues<FitModel>().Select(CurveFit.LabelOf).Cast<object>().ToArray());
        _typeCombo.SelectedIndex = 0;
        _smoothBar.Enabled = false;

        _typeCombo.SelectedIndexChanged += (s, e) => RebuildParams();
        _degreeNum.ValueChanged += (s, e) => RebuildParams();
        _smoothBar.Scroll += (s, e) => _smoothLabel.Text = "平滑度: " + _smoothBar.Value;
        _blendRadio.CheckedChanged += (s, e) => _blendPct.Enabled = _blendRadio.Checked;
    }

    private static Label Section(string text, int x, int y)
        => new() { Text = text, AutoSize = true, Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold), ForeColor = Color.FromArgb(49, 110, 244), Location = new Point(x, y) };

    private static Button MakeButton(string text, Action onClick, int width)
    {
        var b = new Button { Text = text, Width = width };
        b.Click += (s, e) => onClick();
        return b;
    }

    private FitModel SelectedModel()
    {
        var models = Enum.GetValues<FitModel>();
        int idx = Math.Max(0, _typeCombo.SelectedIndex);
        return models[Math.Min(idx, models.Length - 1)];
    }

    private void RebuildParams()
    {
        int degree = decimal.ToInt32(_degreeNum.Value);
        var model = SelectedModel();
        bool isPoly = model == FitModel.Polynomial;
        _degreeNum.Enabled = isPoly;
        _smoothBar.Enabled = isPoly;
        if (!isPoly) _smoothLabel.Text = "平滑度: 0";

        _paramTable.RowStyles.Clear();
        _paramTable.Controls.Clear();
        _paramBoxes.Clear();
        var names = CurveFit.ParamNames(model, degree);
        for (int i = 0; i < names.Length; i++)
        {
            _paramTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            _paramTable.Controls.Add(new Label { Text = names[i] + " =", TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill }, 0, i);
            var box = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(2) };
            box.PlaceholderText = "自动";
            _paramBoxes.Add(box);
            _paramTable.Controls.Add(box, 1, i);
        }
        _paramTable.PerformLayout();
    }

    private FitOptions BuildOptions()
    {
        int degree = decimal.ToInt32(_degreeNum.Value);
        var model = SelectedModel();
        var fixedParams = new double?[_paramBoxes.Count];
        for (int i = 0; i < _paramBoxes.Count; i++)
        {
            string text = _paramBoxes[i].Text.Trim();
            if (text.Length == 0) { fixedParams[i] = null; continue; }
            if (double.TryParse(text, out var v)) fixedParams[i] = v;
        }
        return new FitOptions
        {
            Model = model,
            PolynomialDegree = degree,
            Smoothing = _smoothBar.Value / 100.0,
            FixedParams = fixedParams
        };
    }

    private IReadOnlyList<CurvePoint> GetData()
        => _allRadio.Checked ? _editor.Points : _editor.GetSelectedPoints();

    private void OnPreview()
    {
        var pts = GetData();
        if (pts.Count < 2)
        {
            MessageBox.Show(this, "没有足够的数据点（至少 2 个）", "拟合", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        var r = CurveFit.Fit(pts.Select(p => (p.X, p.Y)).ToArray(), BuildOptions());
        if (r.Evaluate == null)
        {
            MessageBox.Show(this, r.Error, "拟合失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        SetOverlay(r, pts);
        UpdateInfo(r);
    }

    private void OnApply()
    {
        var pts = GetData();
        if (pts.Count < 2)
        {
            MessageBox.Show(this, "没有足够的数据点（至少 2 个）", "拟合", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        var r = CurveFit.Fit(pts.Select(p => (p.X, p.Y)).ToArray(), BuildOptions());
        if (r.Evaluate == null)
        {
            MessageBox.Show(this, r.Error, "拟合失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        ApplyTransform(r, pts);
        SetOverlay(r, pts);
        UpdateInfo(r);
    }

    private void ApplyTransform(FitResult r, IReadOnlyList<CurvePoint> pts)
    {
        if (r.Evaluate == null) return;
        var ev = r.Evaluate;
        double pct = (double)_blendPct.Value / 100.0;
        bool blend = _blendRadio.Checked;
        bool anchor = _anchorCheck.Checked;
        double minX = pts.Min(p => p.X), maxX = pts.Max(p => p.X);

        (double X, double Y) Map(CurvePoint p)
        {
            if (anchor && (Math.Abs(p.X - minX) < 1e-12 || Math.Abs(p.X - maxX) < 1e-12))
                return (p.X, p.Y);
            double fy = ev(p.X);
            double ny = blend ? p.Y + pct * (fy - p.Y) : fy;
            return (p.X, ny);
        }

        if (_allRadio.Checked) _editor.ApplyToAll(Map);
        else _editor.ApplyToSelected(Map);
    }

    private void SetOverlay(FitResult r, IReadOnlyList<CurvePoint> pts)
    {
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
        _editor.OverlayPoints = overlay;
        _editor.Invalidate();
    }

    private void UpdateInfo(FitResult r)
    {
        string ps = string.Join("  ", r.Parameters.Select(p => $"{p.Name}={p.Value:0.####}"));
        _infoLabel.Text = $"{r.Label}{Environment.NewLine}{r.Formula}{Environment.NewLine}R²={r.R2:0.###}  RMSE={r.RMSE:0.###}{Environment.NewLine}{ps}";
    }
}
