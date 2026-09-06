using GameCurve.Excel;
using GameCurve.Models;
using GameCurve.Ui;

namespace GameCurve;

/// <summary>
/// 无界面自测：验证工作簿读取、数值列识别、曲线点生成与单元格回写。
/// </summary>
internal static class TestHarness
{
    public static int Run(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        if (args.Length > 1 && args[1] == "--render")
            return Render(args);
        if (args.Length > 1 && args[1] == "--structure")
            return StructureTest(args);

        string file = args.Length > 1 && File.Exists(args[1])
            ? args[1]
            : @"C:\Users\50690\Desktop\game7\doc\excel\ShopSystem@商城系统.xlsm";

        Console.WriteLine("测试文件: " + file);
        var wb = new WorkbookModel();
        wb.Open(file);
        Console.WriteLine("工作表: " + string.Join(" | ", wb.SheetNames));

        string? firstSheetWithData = null;
        foreach (var sheet in wb.SheetNames)
        {
            var snap = wb.LoadSheet(sheet);
            var nums = snap.Columns.Where(c => c.IsNumericScalar && c.NumericCount > 0).ToList();
            var withData = nums.FirstOrDefault(c => c.NumericCount >= 3);
            if (withData != null && firstSheetWithData == null)
            {
                firstSheetWithData = sheet;
                var vals = snap.GetNumericColumn(withData.ColumnIndex);
                Console.WriteLine($"[{sheet}] 数值列 '{withData.DisplayName}' 类型={withData.Type} 整数={withData.IsInteger} 点数={vals.Count}");
                Console.WriteLine("  样例: " + string.Join(", ", vals.Take(6).Select(v => $"{v.Item1}:{v.Item2}")));
            }
        }

        if (firstSheetWithData == null)
        {
            Console.WriteLine("!! 未找到可编辑数值列");
            return 2;
        }

        // 在副本上做回写测试，保护源文件
        string copy = Path.Combine(Path.GetTempPath(), "gc_selftest_" + Guid.NewGuid().ToString("N")[..8] + ".xlsm");
        File.Copy(file, copy, true);
        var wb2 = new WorkbookModel();
        wb2.Open(copy);
        var snap2 = wb2.LoadSheet(firstSheetWithData);
        var num = snap2.Columns
            .Where(c => c.IsNumericScalar && c.NumericCount >= 3)
            .OrderBy(c => c.IsInteger ? 1 : 0)
            .ThenByDescending(c => c.NumericCount)
            .First();
        var pairs = snap2.GetNumericColumn(num.ColumnIndex);
        var (row, oldVal) = pairs[0];
        double newVal = (num.IsInteger ? Math.Floor(oldVal) : oldVal) + 1.25;
        string cellRef = CellHelper.ToCellReference(num.ColumnIndex, row);

        _ = wb2.TryWriteCells(new[] { new CellWrite(firstSheetWithData, cellRef, newVal, num.IsInteger, 4) }, out var err);
        Console.WriteLine($"写回 {firstSheetWithData}!{cellRef}: {oldVal} -> {newVal}  err={err ?? "无"}");

        wb2.RefreshMeta();
        var snap3 = wb2.LoadSheet(firstSheetWithData);
        var readBack = snap3.GetNumericColumn(num.ColumnIndex).FirstOrDefault(v => v.Item1 == row).Item2;
        double expected = num.IsInteger ? Math.Round(newVal) : Math.Round(newVal, 4);
        bool ok = Math.Abs(readBack - expected) < 1e-3;
        Console.WriteLine($"重读 {cellRef} = {readBack}  -> {(ok ? "OK" : "失败")}");

        // 校验宏是否保留
        using var zip = System.IO.Compression.ZipFile.OpenRead(copy);
        bool hasVba = zip.Entries.Any(e => e.FullName.Contains("vba", StringComparison.OrdinalIgnoreCase));
        Console.WriteLine("宏保留: " + (hasVba ? "是" : "否"));

        try { File.Delete(copy); } catch { }
        return ok ? 0 : 3;
    }

    private static int StructureTest(string[] args)
    {
        string file = args.Length > 2 && File.Exists(args[2])
            ? args[2]
            : @"C:\Users\50690\Desktop\github\GameCurve\test\excel\ShopSystem@商城系统.xlsm";
        Console.WriteLine("结构测试文件: " + file);

        string copy = Path.Combine(Path.GetTempPath(), "gc_struct_" + Guid.NewGuid().ToString("N")[..8] + ".xlsm");
        File.Copy(file, copy, true);
        var wb = new WorkbookModel();
        wb.Open(copy);
        string sheet = wb.SheetNames.First();
        var snap = wb.LoadSheet(sheet);

        // 1) 插入行：验证原行整体下移，新行为空（选一个在该行有值的列做探针）
        int contentGi = Enumerable.Range(0, snap.Grid.Count).FirstOrDefault(r => snap.Grid[r].Any(x => !string.IsNullOrEmpty(x)), 0);
        int insertRow = snap.RowNumbers[contentGi];
        int gi = contentGi;
        int probeCol = Enumerable.Range(0, snap.ColumnCount)
            .FirstOrDefault(c => !string.IsNullOrEmpty(snap.Grid[gi][c]), 0);
        string probeCell = CellHelper.ToCellReference(probeCol, insertRow);
        string beforeVal = snap.Grid[gi][probeCol] ?? "";
        bool ok = wb.TryApplyStructure(new[] { new StructuralOp(sheet, StructuralKind.InsertRow, insertRow) }, out var err);
        Console.WriteLine("插入行[" + insertRow + "]: " + (ok ? "成功" : "失败 " + err));
        wb.RefreshMeta();
        var s2 = wb.LoadSheet(sheet);
        int idxAfter = s2.RowNumbers.IndexOf(insertRow + 1);
        int idxNew = s2.RowNumbers.IndexOf(insertRow);
        bool rowShifted = s2.RowNumbers.Count == snap.RowNumbers.Count + 1
            && s2.RowNumbers.Contains(insertRow)
            && s2.Grid[s2.RowNumbers.IndexOf(insertRow + 1)][probeCol] == beforeVal
            && (s2.Grid[s2.RowNumbers.IndexOf(insertRow)][probeCol] ?? "") == "";
        Console.WriteLine("插入行后单元格从 " + probeCell + " 下移: " + (rowShifted ? "OK" : "失败"));

        // 2) 删除刚插入的空行，恢复
        bool okDel = wb.TryApplyStructure(new[] { new StructuralOp(sheet, StructuralKind.DeleteRow, insertRow) }, out var errDel);
        wb.RefreshMeta();
        var s3 = wb.LoadSheet(sheet);
        bool restored = s3.RowNumbers.Count == snap.RowNumbers.Count
            && s3.Grid[s3.RowNumbers.IndexOf(insertRow)][probeCol] == beforeVal;
        Console.WriteLine("删除空行后恢复: " + (okDel && restored ? "OK" : "失败"));

        // 3) 插入列 + 删除列：净列为 0，值右移后还原（用有值的行做探针）
        int colIdx = Math.Min(2, snap.ColumnCount - 1);
        int probeGi = Enumerable.Range(0, snap.Grid.Count).FirstOrDefault(r => !string.IsNullOrEmpty(snap.Grid[r][colIdx]), 0);
        string beforeColVal = snap.Grid[probeGi][colIdx] ?? "";
        bool okC = wb.TryApplyStructure(new[]
        {
            new StructuralOp(sheet, StructuralKind.InsertColumn, colIdx),
            new StructuralOp(sheet, StructuralKind.DeleteColumn, colIdx)
        }, out var errC);
        wb.RefreshMeta();
        var s4 = wb.LoadSheet(sheet);
        bool colRestored = okC && s4.ColumnCount == snap.ColumnCount
            && s4.Grid[probeGi][colIdx] == beforeColVal;
        Console.WriteLine("插入列+删除列净恢复: " + (colRestored ? "OK" : "失败"));

        // 4) 内存快照行列操作（不写磁盘）
        var orig = wb.LoadSheet(sheet);
        int rRow = orig.RowNumbers[2];
        int origRowCount = orig.RowNumbers.Count;
        orig.InsertRow(rRow);
        bool memRow = orig.RowNumbers.Count == origRowCount + 1
            && orig.RowNumbers[2] == rRow
            && orig.RowNumbers[3] == rRow + 1;
        orig.DeleteRow(rRow);
        bool memRowBack = orig.RowNumbers.Count == origRowCount && orig.RowNumbers[2] == rRow;

        int origCols = orig.ColumnCount;
        orig.InsertColumn(1, "NewB");
        bool memCol = orig.ColumnCount == origCols + 1
            && orig.Columns[1].ColumnIndex == 1
            && orig.Columns[2].ColumnIndex == 2
            && orig.Grid[0].Length == origCols + 1;
        orig.DeleteColumn(1);
        bool memColBack = orig.ColumnCount == origCols
            && orig.Columns[0].ColumnIndex == 0
            && orig.Grid[0].Length == origCols;
        orig.RenameColumn(0, "renamed:A");
        bool memRename = orig.Columns[0].HeaderRaw == "renamed:A" && !orig.Columns[0].IsEmpty;
        Console.WriteLine($"内存行操作: {(memRow && memRowBack ? "OK" : "失败")}; 内存列操作: {(memCol && memColBack ? "OK" : "失败")}; 重命名: {(memRename ? "OK" : "失败")}");

        using var zip = System.IO.Compression.ZipFile.OpenRead(copy);
        bool hasVba = zip.Entries.Any(e => e.FullName.Contains("vba", StringComparison.OrdinalIgnoreCase));
        Console.WriteLine("宏保留: " + (hasVba ? "是" : "否"));

        try { File.Delete(copy); } catch { }
        return rowShifted && restored && colRestored && memRow && memRowBack && memCol && memColBack && memRename ? 0 : 3;
    }

    private static int Render(string[] args)
    {
        string file = args.Length > 2 && File.Exists(args[2]) ? args[2] : @"C:\Users\50690\Desktop\game7\doc\excel\Base@全局数据.xlsm";
        string outPng = args.Length > 3 ? args[3] : Path.Combine(Path.GetTempPath(), "gc_render.png");

        var wb = new WorkbookModel();
        wb.Open(file);

        string? bestSheet = null;
        ColumnMeta? bestCol = null;
        int bestCount = 0;
        foreach (var sheet in wb.SheetNames)
        {
            var snap = wb.LoadSheet(sheet);
            foreach (var col in snap.Columns.Where(c => c.IsNumericScalar))
            {
                if (col.Name.Equals("ID", StringComparison.OrdinalIgnoreCase) || col.Name.Contains("序号")) continue;
                int cnt = snap.GetNumericColumn(col.ColumnIndex).Count;
                if (cnt > bestCount) { bestCount = cnt; bestSheet = sheet; bestCol = col; }
            }
        }
        if (bestCol == null) { Console.WriteLine("!! 无可渲染数值列"); return 2; }

        var sn = wb.LoadSheet(bestSheet!);
        var view = new CurveSeriesView { Name = bestCol.DisplayName, Color = CurveEditor.Palette[0], IsEditable = true };
        foreach (var (row, val) in sn.GetNumericColumn(bestCol.ColumnIndex))
            view.Points.Add(new CurvePoint(row, val, row, false));

        using var editor = new CurveEditor { Width = 1100, Height = 620, ShowSpline = true };
        editor.SetSeries(new List<CurveSeriesView> { view }, 0);
        editor.XAxisLabel = "行号";
        editor.YAxisLabel = bestCol.DisplayName;
        using var bmp = new Bitmap(1100, 620);
        editor.DrawToBitmap(bmp, new Rectangle(0, 0, 1100, 620));
        bmp.Save(outPng, System.Drawing.Imaging.ImageFormat.Png);
        Console.WriteLine($"已渲染 [{bestSheet}] '{bestCol.DisplayName}' ({bestCount} 点) -> {outPng}");
        return 0;
    }
}
