using GameCurve.Excel;
using GameCurve.Models;
using GameCurve.Services;
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
        if (args.Length > 1 && args[1] == "--subcurve")
            return SubCurveTest(args);
        if (args.Length > 1 && args[1] == "--columnorder")
            return ColumnOrderTest(args);
        if (args.Length > 1 && args[1] == "--columnwidth")
            return ColumnWidthTest(args);
        if (args.Length > 1 && args[1] == "--sheetorder")
            return SheetOrderTest(args);
        if (args.Length > 1 && args[1] == "--addsheet")
            return AddSheetTest(args);
        if (args.Length > 1 && args[1] == "--renamesheet")
            return RenameSheetTest(args);

        string file = args.Length > 1 && File.Exists(args[1])
            ? args[1]
            : @"C:\Users\50690\Desktop\game7\doc\excel\ShopSystem@商城系统.xlsm";

        Console.WriteLine("测试文件: " + file);
        var wb = new WorkbookModel();
        wb.Open(file);
        Console.WriteLine("工作表: " + string.Join(" | ", wb.SheetNames));

        // 临时验证：打印每个工作表的曲线选项显示名
        foreach (var s in wb.SheetNames)
        {
            var sn = wb.LoadSheet(s);
            Console.WriteLine($"[{s}] 曲线选项:");
            foreach (var o in SubCurveHelper.BuildOptions(sn).Take(14))
                Console.WriteLine("   " + o.DisplayName);
        }

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

    /// <summary>
    /// 子曲线（单元格拆分列）约定：空白/空值单元格不生成曲线点，曲线操作中的 SetValue
    /// 不可把空白单元格自动改写为 JSON，必须由用户手工在单元格填入内容。
    /// </summary>
    private static int SubCurveTest(string[] args)
    {
        string file = args.Length > 2 && File.Exists(args[2])
            ? args[2]
            : @"C:\Users\50690\Desktop\github\GameCurve\test\excel\ShopSystem@商城系统.xlsm";
        Console.WriteLine("子曲线测试文件: " + file);

        var wb = new WorkbookModel();
        wb.Open(file);

        // 找到第一个含子曲线（数组/JSON）的工作表
        ColumnMeta? subCol = null;
        CurveColumnOption? sub = null;
        SheetSnapshot? subSnap = null;
        foreach (var sheet in wb.SheetNames)
        {
            var snap = wb.LoadSheet(sheet);
            var opt = SubCurveHelper.BuildOptions(snap).FirstOrDefault(o => o.IsSubCurve);
            if (opt != null) { subSnap = snap; subCol = opt.Column; sub = opt; break; }
        }

        if (sub == null || subSnap == null)
        {
            Console.WriteLine("!! 未找到子曲线列");
            return 2;
        }

        int col = sub.Column.ColumnIndex;
        bool ok = true;

        // 1) 空白单元格：TryReadValue 应返回 false，且 SetValue("", ...) 不应改写原格
        string setBlank = SubCurveHelper.SetValue("", sub, 3.0);
        bool blankNoWrite = setBlank == "";
        Console.WriteLine("子曲线 SetValue(空白格) 不改写: " + (blankNoWrite ? "OK" : "失败"));
        ok &= blankNoWrite;

        // 2) 找一个确实有值的行：TryReadValue 为真，且 SetValue 能改写原格
        int goodGi = Enumerable.Range(0, subSnap.Grid.Count)
            .FirstOrDefault(gi => SubCurveHelper.TryReadValue(subSnap, sub, gi, out _), -1);
        bool goodVal = goodGi >= 0;
        double v0 = 0;
        if (goodGi >= 0)
            goodVal = SubCurveHelper.TryReadValue(subSnap, sub, goodGi, out v0);
        bool goodWrite = goodVal;
        if (goodVal)
        {
            string raw = subSnap.Grid[goodGi][col] ?? "";
            string after = SubCurveHelper.SetValue(raw, sub, v0 + 1.0);
            goodWrite = after != raw;
        }
        Console.WriteLine("子曲线有值格可读取/改写: " + (goodVal && goodWrite ? "OK" : "失败"));
        ok &= goodVal && goodWrite;

        // 3) 找空白行：TryReadValue 为 false（不生成曲线点）
        int blankGi = Enumerable.Range(0, subSnap.Grid.Count)
            .FirstOrDefault(gi => string.IsNullOrWhiteSpace(subSnap.Grid[gi][col] ?? ""), -1);
        bool blankNoPoint = blankGi >= 0 && !SubCurveHelper.TryReadValue(subSnap, sub, blankGi, out _);
        Console.WriteLine("子曲线空白行不生成点: " + (blankNoPoint ? "OK" : "失败"));
        ok &= blankNoPoint;

        wb.Dispose();
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

        // 5) 批量插入多行：一次插入 count 个空行，验证行数净增 count、原数据下移 count 行
        int count = 7;
        int batchInsert = insertRow; // 在原来有数据的行位置插入
        bool okBatch = wb.TryApplyStructure(
            Enumerable.Range(0, count).Select(_ => new StructuralOp(sheet, StructuralKind.InsertRow, batchInsert)).ToList(),
            out var errBatch);
        wb.RefreshMeta();
        var s5 = wb.LoadSheet(sheet);
        bool batchOk = okBatch
            && s5.RowNumbers.Count == snap.RowNumbers.Count + count
            && s5.Grid[s5.RowNumbers.IndexOf(batchInsert + count)][probeCol] == beforeVal
            && Enumerable.Range(0, count).All(i =>
                s5.RowNumbers.Contains(batchInsert + i) &&
                (s5.Grid[s5.RowNumbers.IndexOf(batchInsert + i)][probeCol] ?? "") == "");
        Console.WriteLine($"批量插入 {count} 行: {(batchOk ? "OK" : "失败")}");

        using var zip = System.IO.Compression.ZipFile.OpenRead(copy);
        bool hasVba = zip.Entries.Any(e => e.FullName.Contains("vba", StringComparison.OrdinalIgnoreCase));
        Console.WriteLine("宏保留: " + (hasVba ? "是" : "否"));

        try { File.Delete(copy); } catch { }
        return rowShifted && restored && colRestored && memRow && memRowBack && memCol && memColBack && memRename && batchOk ? 0 : 3;
    }

    private static int ColumnOrderTest(string[] args)
    {
        string file = args.Length > 2 && File.Exists(args[2])
            ? args[2]
            : @"C:\Users\50690\Desktop\github\GameCurve\test\excel\Base@全局数据.xlsm";
        Console.WriteLine("列顺序测试文件: " + file);

        string copy = Path.Combine(Path.GetTempPath(), "gc_colorder_" + Guid.NewGuid().ToString("N")[..8] + ".xlsm");
        File.Copy(file, copy, true);
        bool allOk = true;
        WorkbookModel? wb2 = null;
        try
        {
            var wb = new WorkbookModel();
            wb.Open(copy);
            string sheet = wb.SheetNames.First();
            int colCount = wb.LoadSheet(sheet).ColumnCount;
            var order = Enumerable.Range(0, colCount).Reverse().ToList(); // 倒序
            bool writeOk = wb.TryWriteColumnOrder(sheet, order, out var err);
            Console.WriteLine("写列顺序: " + (writeOk ? "OK" : "失败 " + err));
            if (!writeOk) return 3;
            wb.Dispose();

            wb2 = new WorkbookModel();
            wb2.Open(copy);
            bool hiddenFiltered = !wb2.SheetNames.Contains("__GameCurve__", StringComparer.OrdinalIgnoreCase);
            Console.WriteLine("内部配置表已隐藏: " + (hiddenFiltered ? "OK" : "失败"));
            var read = wb2.GetColumnOrder(sheet);
            bool orderOk = read != null && read.SequenceEqual(order);
            Console.WriteLine("重读列顺序: " + (orderOk ? "OK" : "失败 " + (read == null ? "null" : string.Join(",", read))));
            allOk &= hiddenFiltered && orderOk;

            using var zip = System.IO.Compression.ZipFile.OpenRead(copy);
            bool hasVba = zip.Entries.Any(e => e.FullName.Contains("vba", StringComparison.OrdinalIgnoreCase));
            Console.WriteLine("宏保留: " + (hasVba ? "是" : "否"));
            allOk &= hasVba;
        }
        finally
        {
            wb2?.Dispose();
            try { File.Delete(copy); } catch { }
        }
        return allOk ? 0 : 3;
    }

    private static int ColumnWidthTest(string[] args)
    {
        string file = args.Length > 2 && File.Exists(args[2])
            ? args[2]
            : @"C:\Users\50690\Desktop\github\GameCurve\test\excel\ShopSystem@商城系统.xlsm";
        Console.WriteLine("列宽测试文件: " + file);

        string copy = Path.Combine(Path.GetTempPath(), "gc_colwidth_" + Guid.NewGuid().ToString("N")[..8] + ".xlsm");
        File.Copy(file, copy, true);
        bool allOk = true;
        WorkbookModel? wb2 = null;
        try
        {
            var wb = new WorkbookModel();
            wb.Open(copy);
            string sheet = wb.SheetNames.First();
            var snap = wb.LoadSheet(sheet);
            Console.WriteLine("读取原始列宽（字符单位）:");
            for (int c = 0; c < snap.ColumnCount; c++)
                Console.WriteLine($"  col{c} ({CellHelper.ColumnIndexToLetter(c)}) = {snap.Columns[c].Width?.ToString("0.##") ?? "null"}");

            // 写回一组新的列宽（每列互不相同，覆盖所有列）
            var widths = new List<(int Col, double Width)>();
            for (int c = 0; c < snap.ColumnCount; c++)
                widths.Add((c, 10.0 + c * 1.5));
            bool writeOk = wb.TryWriteColumnWidths(sheet, widths, out var err);
            Console.WriteLine("写列宽: " + (writeOk ? "OK" : "失败 " + err));
            if (!writeOk) return 3;
            wb.Dispose();

            wb2 = new WorkbookModel();
            wb2.Open(copy);
            var snap2 = wb2.LoadSheet(sheet);
            bool allMatch = true;
            for (int c = 0; c < snap2.ColumnCount; c++)
            {
                double expect = 10.0 + c * 1.5;
                double? actual = snap2.Columns[c].Width;
                if (actual == null || Math.Abs(actual.Value - expect) > 1e-6)
                {
                    string actualStr = actual.HasValue ? actual.Value.ToString("0.##") : "null";
                    Console.WriteLine($"  col{c} 期望 {expect.ToString("0.##")} 实际 {actualStr} -> 不一致");
                    allMatch = false;
                }
            }
            Console.WriteLine("重读列宽一致: " + (allMatch ? "OK" : "失败"));
            allOk &= allMatch;

            // 只写部分列，验证其它列被保留
            var partial = new List<(int, double)> { (0, 42.0) };
            bool partialOk = wb2.TryWriteColumnWidths(sheet, partial, out var errP);
            Console.WriteLine("写部分列宽: " + (partialOk ? "OK" : "失败 " + errP));
            wb2.Dispose();

            wb2 = new WorkbookModel();
            wb2.Open(copy);
            var snap3 = wb2.LoadSheet(sheet);
            bool partialPreserved = Math.Abs((snap3.Columns[0].Width ?? 0) - 42.0) < 1e-6
                && Math.Abs((snap3.Columns[snap3.ColumnCount - 1].Width ?? 0) - (10.0 + (snap3.ColumnCount - 1) * 1.5)) < 1e-6;
            Console.WriteLine("部分写回时其它列保留: " + (partialPreserved ? "OK" : "失败"));
            allOk &= partialPreserved;

            using var zip = System.IO.Compression.ZipFile.OpenRead(copy);
            bool hasVba = zip.Entries.Any(e => e.FullName.Contains("vba", StringComparison.OrdinalIgnoreCase));
            Console.WriteLine("宏保留: " + (hasVba ? "是" : "否"));
            allOk &= hasVba;
        }
        finally
        {
            wb2?.Dispose();
            try { File.Delete(copy); } catch { }
        }
        return allOk ? 0 : 3;
    }

    private static int SheetOrderTest(string[] args)
    {
        string file = args.Length > 2 && File.Exists(args[2])
            ? args[2]
            : @"C:\Users\50690\Desktop\github\GameCurve\test\excel\Base@全局数据.xlsm";
        Console.WriteLine("工作表顺序测试文件: " + file);

        string copy = Path.Combine(Path.GetTempPath(), "gc_sheetorder_" + Guid.NewGuid().ToString("N")[..8] + ".xlsm");
        File.Copy(file, copy, true);
        bool allOk = true;
        WorkbookModel? wb2 = null;
        try
        {
            var wb = new WorkbookModel();
            wb.Open(copy);
            var names = wb.SheetNames.ToList();
            var order = names.AsEnumerable().Reverse().ToList(); // 倒序
            bool writeOk = wb.TryWriteSheetOrder(order, out var err);
            Console.WriteLine("写工作表顺序: " + (writeOk ? "OK" : "失败 " + err));
            if (!writeOk) return 3;
            // 再写一个列顺序，验证不会把工作表顺序覆盖掉
            string first = names.First();
            int colCount = wb.LoadSheet(first).ColumnCount;
            var colOrder = Enumerable.Range(0, colCount).Reverse().ToList();
            bool colOk = wb.TryWriteColumnOrder(first, colOrder, out var errCol);
            Console.WriteLine("写列顺序（应保留工作表顺序）: " + (colOk ? "OK" : "失败 " + errCol));
            wb.Dispose();

            wb2 = new WorkbookModel();
            wb2.Open(copy);
            var read = wb2.GetSheetDisplayOrder();
            bool orderOk = read != null && read.SequenceEqual(order);
            Console.WriteLine("重读工作表顺序: " + (orderOk ? "OK" : "失败 " + (read == null ? "null" : string.Join(",", read))));
            allOk &= orderOk;

            using var zip = System.IO.Compression.ZipFile.OpenRead(copy);
            bool hasVba = zip.Entries.Any(e => e.FullName.Contains("vba", StringComparison.OrdinalIgnoreCase));
            Console.WriteLine("宏保留: " + (hasVba ? "是" : "否"));
            allOk &= hasVba;
        }
        finally
        {
            wb2?.Dispose();
            try { File.Delete(copy); } catch { }
        }
        return allOk ? 0 : 3;
    }

    private static int AddSheetTest(string[] args)
    {
        string file = args.Length > 2 && File.Exists(args[2])
            ? args[2]
            : @"C:\Users\50690\Desktop\github\GameCurve\test\excel\ShopSystem@商城系统.xlsm";
        Console.WriteLine("新建工作表测试文件: " + file);

        string copy = Path.Combine(Path.GetTempPath(), "gc_addsheet_" + Guid.NewGuid().ToString("N")[..8] + ".xlsm");
        File.Copy(file, copy, true);
        bool allOk = true;
        WorkbookModel? wb2 = null;
        try
        {
            var wb = new WorkbookModel();
            wb.Open(copy);
            int before = wb.SheetNames.Count;
            bool addOk = wb.TryAddSheet("新表99", out var err);
            Console.WriteLine("新建工作表: " + (addOk ? "OK" : "失败 " + err));
            if (!addOk) return 3;
            bool exists = wb.SheetNames.Contains("新表99", StringComparer.OrdinalIgnoreCase);
            Console.WriteLine(("新建后存在于列表: " + (exists ? "OK" : "失败")));
            var snap = wb.LoadSheet("新表99");
            bool usable = snap.ColumnCount >= 1 && snap.DataRowCount >= 1;
            Console.WriteLine("新表可加载且有列/行: " + (usable ? "OK" : "失败"));
            wb.Dispose();

            wb2 = new WorkbookModel();
            wb2.Open(copy);
            bool exists2 = wb2.SheetNames.Count == before + 1
                && wb2.SheetNames.Contains("新表99", StringComparer.OrdinalIgnoreCase);
            Console.WriteLine("关闭后重读仍存在: " + (exists2 ? "OK" : "失败"));
            allOk &= exists && usable && exists2;

            using var zip = System.IO.Compression.ZipFile.OpenRead(copy);
            bool hasVba = zip.Entries.Any(e => e.FullName.Contains("vba", StringComparison.OrdinalIgnoreCase));
            Console.WriteLine("宏保留: " + (hasVba ? "是" : "否"));
            allOk &= hasVba;
        }
        finally
        {
            wb2?.Dispose();
            try { File.Delete(copy); } catch { }
        }
        return allOk ? 0 : 3;
    }

    private static int RenameSheetTest(string[] args)
    {
        string file = args.Length > 2 && File.Exists(args[2])
            ? args[2]
            : @"C:\Users\50690\Desktop\github\GameCurve\test\excel\ShopSystem@商城系统.xlsm";
        Console.WriteLine("重命名工作表测试文件: " + file);

        string copy = Path.Combine(Path.GetTempPath(), "gc_renamesheet_" + Guid.NewGuid().ToString("N")[..8] + ".xlsm");
        File.Copy(file, copy, true);
        bool allOk = true;
        WorkbookModel? wb2 = null;
        try
        {
            var wb = new WorkbookModel();
            wb.Open(copy);
            string oldName = wb.SheetNames.First();
            string newName = "重命名测试表";
            // 先写入列顺序与工作表顺序，验证改名会同步迁移两种顺序
            int colCount = wb.LoadSheet(oldName).ColumnCount;
            var colOrder = Enumerable.Range(0, colCount).Reverse().ToList();
            bool colOk = wb.TryWriteColumnOrder(oldName, colOrder, out var errCol);
            var sheetOrder = wb.SheetNames.ToList();
            bool orderOk = wb.TryWriteSheetOrder(sheetOrder, out var errOrder);
            Console.WriteLine("预写列顺序: " + (colOk ? "OK" : "失败 " + errCol));
            Console.WriteLine("预写工作表顺序: " + (orderOk ? "OK" : "失败 " + errOrder));

            bool renameOk = wb.TryRenameSheet(oldName, newName, out var errR);
            Console.WriteLine("重命名: " + (renameOk ? "OK" : "失败 " + errR));
            if (!renameOk) return 3;
            bool namesOk = !wb.SheetNames.Contains(oldName, StringComparer.OrdinalIgnoreCase)
                && wb.SheetNames.Contains(newName, StringComparer.OrdinalIgnoreCase);
            var colAfter = wb.GetColumnOrder(newName);
            bool colMigrated = colAfter != null && colAfter.SequenceEqual(colOrder);
            var sheetOrderAfter = wb.GetSheetDisplayOrder();
            bool orderMigrated = sheetOrderAfter != null
                && sheetOrderAfter.Contains(newName, StringComparer.OrdinalIgnoreCase)
                && !sheetOrderAfter.Contains(oldName, StringComparer.OrdinalIgnoreCase);
            Console.WriteLine("列表已更新: " + (namesOk ? "OK" : "失败"));
            Console.WriteLine("列顺序键迁移: " + (colMigrated ? "OK" : "失败"));
            Console.WriteLine("工作表顺序迁移: " + (orderMigrated ? "OK" : "失败"));
            wb.Dispose();

            wb2 = new WorkbookModel();
            wb2.Open(copy);
            bool persistent = wb2.SheetNames.Contains(newName, StringComparer.OrdinalIgnoreCase)
                && !wb2.SheetNames.Contains(oldName, StringComparer.OrdinalIgnoreCase);
            Console.WriteLine("关闭后重读: " + (persistent ? "OK" : "失败"));

            // 重名与非法字符校验
            var dupOk = wb2.TryRenameSheet(newName, wb2.SheetNames.First(n => !string.Equals(n, newName, StringComparison.OrdinalIgnoreCase)), out var errDup);
            Console.WriteLine("重名校验: " + (!dupOk && errDup != null ? "OK" : "失败"));
            var badOk = wb2.TryRenameSheet(newName, "bad/name", out var errBad);
            Console.WriteLine("非法字符校验: " + (!badOk && errBad != null ? "OK" : "失败"));
            allOk &= namesOk && colMigrated && orderMigrated && persistent && !dupOk && !badOk;

            using var zip = System.IO.Compression.ZipFile.OpenRead(copy);
            bool hasVba = zip.Entries.Any(e => e.FullName.Contains("vba", StringComparison.OrdinalIgnoreCase));
            Console.WriteLine("宏保留: " + (hasVba ? "是" : "否"));
            allOk &= hasVba;
        }
        finally
        {
            wb2?.Dispose();
            try { File.Delete(copy); } catch { }
        }
        return allOk ? 0 : 3;
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
