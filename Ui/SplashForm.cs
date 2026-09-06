namespace GameCurve.Ui;

/// <summary>
/// 无边框启动/加载画面：标题 + 动态文字 + 跑马灯进度条。
/// 用于打开工作簿或切换工作表等耗时操作期间，避免界面看起来无响应。
/// </summary>
public sealed class SplashForm : Form
{
    private readonly Label _title = new()
    {
        Text = "GameCurve",
        Font = new Font("Microsoft YaHei UI", 17f, FontStyle.Bold),
        ForeColor = Color.FromArgb(49, 110, 244),
        TextAlign = ContentAlignment.MiddleCenter,
        Dock = DockStyle.Top,
        Height = 60,
        BackColor = Color.FromArgb(246, 250, 255)
    };

    private readonly Label _subtitle = new()
    {
        Text = "正在加载...",
        Font = new Font("Microsoft YaHei UI", 10f),
        ForeColor = Color.FromArgb(80, 86, 94),
        TextAlign = ContentAlignment.MiddleCenter,
        Dock = DockStyle.Fill,
        BackColor = Color.FromArgb(246, 250, 255)
    };

    private readonly ProgressBar _progress = new()
    {
        Dock = DockStyle.Bottom,
        Height = 16,
        Style = ProgressBarStyle.Marquee,
        MarqueeAnimationSpeed = 30
    };

    public SplashForm()
    {
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(440, 190);
        BackColor = Color.FromArgb(246, 250, 255);
        ShowInTaskbar = false;
        TopMost = true;
        Controls.Add(_progress);
        Controls.Add(_title);
        Controls.Add(_subtitle);
    }

    public void SetText(string text) => _subtitle.Text = text;

    protected override bool ShowWithoutActivation => true;
}
