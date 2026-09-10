using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Tools.Excel;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using static 激光报表.ToolsHelper;

namespace 激光报表
{
    internal class ExcelHelper
    {
        /// <summary>
        /// 根据合同编号从订单二维数组中筛选匹配的数据行，返回结果包含标题行。
        /// 兼容从 Excel 读取的 1‑based 数组和普通 0‑based 数组。
        /// </summary>
        /// <param name="contractNumber">要匹配的合同编号。</param>
        /// <param name="orderArray">
        /// 包含订单数据的二维数组。第一行为标题行，其余行为数据行。
        /// 标题行中必须包含名称为“合同编号”的列。
        /// </param>
        /// <returns>
        /// 一个列表，其中第一行为标题行（object[]），后续每行为匹配的数据行。
        /// 若未找到“合同编号”列，则返回空列表。
        /// 若找到该列但无匹配数据，返回仅含标题行的列表。
        /// </returns>
        /// <exception cref="ArgumentNullException">参数为 null 时抛出。</exception>
        public static List<object[]> GetFilteredRowsByContractNumber(string contractNumber, object[,] orderArray)
        {
            if (contractNumber == null)
                throw new ArgumentNullException(nameof(contractNumber));
            if (orderArray == null)
                throw new ArgumentNullException(nameof(orderArray));

            // 获取数组维度边界（适配 0‑based 或 1‑based）
            int rowLower = orderArray.GetLowerBound(0);
            int rowUpper = orderArray.GetUpperBound(0);
            int colLower = orderArray.GetLowerBound(1);
            int colUpper = orderArray.GetUpperBound(1);

            int rowCount = rowUpper - rowLower + 1;
            int colCount = colUpper - colLower + 1;

            // 无数据时返回空列表（至少需要标题行）
            if (rowCount == 0)
                return new List<object[]>();

            // 1. 查找“合同编号”所在列索引
            int contractColIndex = -1;
            for (int j = colLower; j <= colUpper; j++)
            {
                if (orderArray[rowLower, j]?.ToString() == "合同编号")
                {
                    contractColIndex = j;
                    break;
                }
            }

            // 若未找到该列，无法筛选，返回空列表
            if (contractColIndex == -1)
                return new List<object[]>();

            // 2. 构建结果列表，首先加入标题行
            var result = new List<object[]>();

            // 复制标题行（转为 0‑based 的 object[]）
            object[] headerRow = new object[colCount];
            for (int j = colLower; j <= colUpper; j++)
            {
                headerRow[j - colLower] = orderArray[rowLower, j];
            }
            result.Add(headerRow);

            // 3. 遍历数据行，筛选匹配合同编号的行
            for (int i = rowLower + 1; i <= rowUpper; i++)
            {
                object cellValue = orderArray[i, contractColIndex];
                string cellString = cellValue?.ToString();

                if (cellString == contractNumber)
                {
                    object[] dataRow = new object[colCount];
                    for (int j = colLower; j <= colUpper; j++)
                    {
                        dataRow[j - colLower] = orderArray[i, j];
                    }
                    result.Add(dataRow);
                }
            }
            return result;
        }

        /// <summary>
        /// 根据合同编号和物料编码组合从订单二维数组中筛选匹配的数据行。
        /// 只要数据行的合同编号和物料编码同时等于列表中某一组合，即视为匹配。
        /// 结果第一行为标题行，后续为匹配的数据行。
        /// </summary>
        /// <param name="filters">合同与物料编码的组合条件列表。不能为 <c>null</c>，可以为空列表（此时只返回标题行）。</param>
        /// <param name="orderArray">
        /// 包含订单数据的二维数组。第一行为标题行，必须包含“合同编号”和“物料编码”列。
        /// 兼容 0‑based 和 1‑based 数组（如 Excel Range.Value）。
        /// </param>
        /// <returns>
        /// 包含标题行及所有匹配数据行的列表。若列标题缺失则返回空列表；
        /// 若列存在但无匹配，返回仅含标题行的列表。
        /// </returns>
        /// <exception cref="ArgumentNullException">任一参数为 <c>null</c> 时抛出。</exception>
        public static List<object[]> GetFilteredRowsByContractAndMaterial(
            List<ContractMaterialData> filters, object[,] orderArray)
        {
            if (filters == null)
                throw new ArgumentNullException(nameof(filters));
            if (orderArray == null)
                throw new ArgumentNullException(nameof(orderArray));
            // 获取数组维度边界（适配 0‑based 或 1‑based）
            int rowLower = orderArray.GetLowerBound(0);
            int rowUpper = orderArray.GetUpperBound(0);
            int colLower = orderArray.GetLowerBound(1);
            int colUpper = orderArray.GetUpperBound(1);
            int rowCount = rowUpper - rowLower + 1;
            int colCount = colUpper - colLower + 1;
            if (rowCount == 0)
                return new List<object[]>();
            // 查找“合同编号”列索引
            int contractColIndex = -1;
            int materialColIndex = -1;

            for (int j = colLower; j <= colUpper; j++)
            {
                string header = orderArray[rowLower, j]?.ToString();
                if (header == "合同编号")
                    contractColIndex = j;
                else if (header == "物料编码")
                    materialColIndex = j;
            }

            // 任一必需列缺失，无法筛选，返回空列表
            if (contractColIndex == -1 || materialColIndex == -1)
                return new List<object[]>();

            // 构建高效查找结构：将所有组合转换为 HashSet（Tuple 或 ValueTuple）
            var conditionSet = new HashSet<(string Contract, string Material)>();
            foreach (var item in filters)
            {
                // 跳过 null 条目，避免空引用
                if (item != null)
                    conditionSet.Add((item.ContractNumber, item.MaterialCode));
            }

            // 结果列表，首先加入标题行
            var result = new List<object[]>();

            // 复制标题行
            object[] headerRow = new object[colCount];
            for (int j = colLower; j <= colUpper; j++)
            {
                headerRow[j - colLower] = orderArray[rowLower, j];
            }
            result.Add(headerRow);

            // 遍历数据行
            for (int i = rowLower + 1; i <= rowUpper; i++)
            {
                string contractValue = orderArray[i, contractColIndex]?.ToString();
                string materialValue = orderArray[i, materialColIndex]?.ToString();

                // 检查是否满足任一组合条件
                if (conditionSet.Contains((contractValue, materialValue)))
                {
                    object[] row = new object[colCount];
                    for (int j = colLower; j <= colUpper; j++)
                    {
                        row[j - colLower] = orderArray[i, j];
                    }
                    result.Add(row);
                }
            }
            return result;
        }

        /// <summary>
        /// 打包
        /// </summary>
        /// <param name="range"></param>
        /// <param name="packSheet"></param>
        public static void CopyPack(Range range, Worksheet packSheet)
        {
            // 1. 单行检查
            if (range.Rows.Count != 1)
            {
                System.Windows.Forms.MessageBox.Show("请选择单行区域！", "提示",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }

            Worksheet sourceSheet = range.Worksheet;
            int sourceRow = range.Row;

            // 标准化换行符，兼容不同环境
            string Normalize(string s)
            {
                if (string.IsNullOrEmpty(s)) return s;
                return s.Replace("\r\n", "\n").Replace("\r", "\n").Trim();
            }

            // 2. 读取源表标题行（固定第4行）并构建标题->列号字典
            object[,] sourceHeaderValues = sourceSheet.Rows[4].Value as object[,];
            var sourceHeaders = new System.Collections.Generic.Dictionary<string, int>();
            if (sourceHeaderValues != null)
            {
                int srcLB = sourceHeaderValues.GetLowerBound(1);
                int srcCount = sourceHeaderValues.GetLength(1);
                for (int c = srcLB; c < srcLB + srcCount; c++)
                {
                    string header = Normalize(sourceHeaderValues[1, c]?.ToString());
                    if (!string.IsNullOrEmpty(header) && !sourceHeaders.ContainsKey(header))
                        sourceHeaders[header] = c;
                }
            }

            // 3. 确保源数据区域为整行（防止单单元格导致 Value 不是数组）
            if (range.Count == 1 || range.Columns.Count < sourceSheet.UsedRange.Columns.Count)
                range = range.EntireRow;

            object[,] sourceRowValues = range.Value as object[,];
            if (sourceRowValues == null)
            {
                System.Windows.Forms.MessageBox.Show("源行无数据！", "提示");
                return;
            }

            // 4. 读取 packSheet 标题行（固定第1行）并构建标题->列号字典
            object[,] packHeaderValues = packSheet.Rows[1].Value as object[,];
            var packHeaders = new System.Collections.Generic.Dictionary<string, int>();
            if (packHeaderValues != null)
            {
                int pLB = packHeaderValues.GetLowerBound(1);
                int pCount = packHeaderValues.GetLength(1);
                for (int c = pLB; c < pLB + pCount; c++)
                {
                    string header = Normalize(packHeaderValues[1, c]?.ToString());
                    if (!string.IsNullOrEmpty(header) && !packHeaders.ContainsKey(header))
                        packHeaders[header] = c;
                }
            }

            // 5. 将 packSheet 全部数据读入内存（用于定位空行、生成柜号、计算毛重）
            Range usedRange = packSheet.UsedRange;
            int packFirstRow = usedRange.Row;
            object[,] packData = null;
            int packRows = usedRange.Rows.Count;
            int packCols = usedRange.Columns.Count;

            if (packRows == 1 && packCols == 1)
            {
                object singleVal = usedRange.Value;
                if (singleVal != null && !string.IsNullOrEmpty(singleVal.ToString()))
                {
                    packData = new object[1, 1];
                    packData[1, 1] = singleVal;
                }
            }
            else if (packRows > 0 && packCols > 0)
            {
                packData = usedRange.Value as object[,];
            }

            // 6. 定位目标写入行：从下向上找最后一个有数据的行，目标行为其下一行
            int lastDataRow = 0;
            if (packData != null)
            {
                int d1Lo = packData.GetLowerBound(0);
                int d1Hi = packData.GetUpperBound(0);
                int d2Lo = packData.GetLowerBound(1);
                int d2Hi = packData.GetUpperBound(1);
                for (int r = d1Hi; r >= d1Lo; r--)
                {
                    bool hasData = false;
                    for (int c = d2Lo; c <= d2Hi; c++)
                    {
                        if (packData[r, c] != null)
                        {
                            hasData = true;
                            break;
                        }
                    }
                    if (hasData)
                    {
                        lastDataRow = packFirstRow + (r - d1Lo);
                        break;
                    }
                }
            }
            int targetRow = lastDataRow > 0 ? lastDataRow + 1 : 2;  // 至少为第2行（第1行是标题）

            // 7. 按列映射复制数据（源“未打包数” → 目标“打包\n数量”）
            var colMap = new System.Collections.Generic.Dictionary<string, string>
    {
        { Normalize("日期"), Normalize("日期") },
        { Normalize("班"), Normalize("班") },
        { Normalize("机台\n编号"), Normalize("机台\n编号") },
        { Normalize("计划单号"), Normalize("计划单号") },
        { Normalize("合同编号"), Normalize("合同编号") },
        { Normalize("物料编码"), Normalize("物料编码") },
        { Normalize("未打包数"), Normalize("打包\n数量") }
    };

            string contractNo = null;
            foreach (var kv in colMap)
            {
                if (sourceHeaders.TryGetValue(kv.Key, out int srcCol) &&
                    packHeaders.TryGetValue(kv.Value, out int dstCol))
                {
                    object val = sourceRowValues[1, srcCol];
                    packSheet.Cells[targetRow, dstCol].Value = val;
                    if (kv.Key == Normalize("合同编号"))
                        contractNo = val?.ToString();
                }
            }

            if (string.IsNullOrEmpty(contractNo))
            {
                System.Windows.Forms.MessageBox.Show("未找到合同编号，无法生成柜号。", "提示");
                return;
            }

            // 8. 生成柜号（前缀和序号均为两位数字，不足补零）
            if (!packHeaders.TryGetValue(Normalize("合同编号"), out int contractColIdx) ||
                !packHeaders.TryGetValue(Normalize("柜号"), out int cabinetColIdx))
            {
                System.Windows.Forms.MessageBox.Show("packSheet 缺少“合同编号”或“柜号”列。", "错误");
                return;
            }

            string newCabinetNo = "01#-01";  // 默认柜号
            if (packData != null)
            {
                int d1Lo = packData.GetLowerBound(0);
                int d1Hi = packData.GetUpperBound(0);
                var regex = new System.Text.RegularExpressions.Regex(@"^(\d+)#-(\d+)$");
                long maxFullNumber = -1;
                string bestPrefix = "01";
                int bestSeq = 0;
                for (int r = d1Lo; r <= d1Hi; r++)
                {
                    int excelRow = packFirstRow + (r - d1Lo);
                    if (excelRow == 1) continue;   // 跳过标题行

                    string cNo = packData[r, contractColIdx]?.ToString();
                    if (cNo != contractNo) continue;

                    string cab = packData[r, cabinetColIdx]?.ToString();
                    var m = regex.Match(cab ?? "");
                    if (!m.Success) continue;

                    string pre = m.Groups[1].Value;
                    string seqStr = m.Groups[2].Value;
                    int seq = int.Parse(seqStr);

                    // 组合成完整数值进行比较，例如 06 + 25 → 0625
                    long fullNum = long.Parse(pre + seqStr);

                    if (fullNum > maxFullNumber)
                    {
                        maxFullNumber = fullNum;
                        bestPrefix = pre;
                        bestSeq = seq;
                    }
                }
                if (maxFullNumber >= 0)
                {
                    int newSeq = bestSeq + 1;
                    // 统一格式化为两位数，不足两位补0
                    newCabinetNo = $"{bestPrefix.PadLeft(2, '0')}#-{newSeq:D2}";
                }
            }
            packSheet.Cells[targetRow, cabinetColIdx].Value = newCabinetNo;
            // 9. 计算毛重合计并弹窗
            if (packHeaders.TryGetValue(Normalize("毛重"), out int grossColIdx))
            {
                var prefixRegex = new System.Text.RegularExpressions.Regex(@"^(\d+)#-");
                string newPrefix = prefixRegex.Match(newCabinetNo).Groups[1].Value;
                if (string.IsNullOrEmpty(newPrefix)) newPrefix = "01";

                double totalGross = 0;
                if (packData != null)
                {
                    int d1Lo = packData.GetLowerBound(0);
                    int d1Hi = packData.GetUpperBound(0);
                    for (int r = d1Lo; r <= d1Hi; r++)
                    {
                        int excelRow = packFirstRow + (r - d1Lo);
                        if (excelRow == 1) continue;

                        string cNo = packData[r, contractColIdx]?.ToString();
                        if (cNo != contractNo) continue;

                        string cab = packData[r, cabinetColIdx]?.ToString();
                        string cabPrefix = prefixRegex.Match(cab ?? "").Groups[1].Value;
                        if (cabPrefix == newPrefix)
                        {
                            if (double.TryParse(packData[r, grossColIdx]?.ToString(), out double gross))
                                totalGross += gross;
                        }
                    }
                }
                // 最后跳转显示 packSheet 工作表
                packSheet.Activate();
                System.Windows.Forms.MessageBox.Show(
            $"柜号：{newCabinetNo}\n毛重合计：{totalGross}",
            "毛重合计", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
        }
        /// <summary>
        /// 处理生产日报中的张数计算（含O列内容自动更新）
        /// 更新规则：
        ///   - 在 O 列合并区域第一个单元格的文本中，
        ///     若存在“张”字，则在其前面：有数字→替换为新张数，无数字→插入新张数；
        ///     若存在“/”，则处理其前面的数字（有数字→数字+张数后替换，无数字→直接插入张数）。
        /// </summary>
        /// <param name="range">当前选中的单元格区域</param>
        public static void ProcessProductionDaily(Range range)
        {
            // 1. 判断工作表名称是否为“生产日报”
            if (range.Worksheet.Name != "生产日报")
                return;

            // 2. 弹出输入框，获取整数张数
            object input = range.Application.InputBox(
                Prompt: "请输入张数：",
                Title: "张数输入",
                Type: 1);   // Type: 1 表示数字

            if (input is bool && (bool)input == false)
                return;

            int quantity = Convert.ToInt32(input);

            // 3. 获取 range 所在行 O 列的单元格，并得到其合并区域
            int row = range.Row;
            Range cellO = range.Worksheet.Cells[row, 15]; // O 列为第15列
            Range mergeArea = cellO.MergeArea;

            // --- 处理 O 列合并单元格中的文本值 ---
            Range firstCellO = mergeArea.Cells[1, 1];
            object originalValue = firstCellO.Value;

            if (originalValue != null)
            {
                string text = originalValue.ToString();
                string modifiedText = text;

                // 4a. 处理“张”字：若前面有数字则替换，否则插入
                int zhangIndex = modifiedText.IndexOf("张");
                if (zhangIndex >= 0)
                {
                    // 从“张”前一个字符开始，向前查找连续的数字
                    int digitEnd = zhangIndex - 1;
                    int digitStart = -1;
                    for (int i = digitEnd; i >= 0; i--)
                    {
                        if (char.IsDigit(modifiedText[i]))
                        {
                            digitStart = i;
                        }
                        else
                        {
                            break;
                        }
                    }

                    if (digitStart >= 0)
                    {
                        // 存在前面的数字，替换之
                        int digitLength = digitEnd - digitStart + 1;
                        modifiedText = modifiedText.Remove(digitStart, digitLength)
                                                   .Insert(digitStart, quantity.ToString());
                    }
                    else
                    {
                        // 前面没有数字，直接插入
                        modifiedText = modifiedText.Insert(zhangIndex, quantity.ToString());
                    }
                }

                // 仅当文本发生改变时写回合并单元格
                if (modifiedText != text)
                {
                    firstCellO.Value = modifiedText;
                }
            }

            // 5. 遍历合并区域中的每一个单元格，在对应行的 J 列写入公式
            foreach (Range cell in mergeArea.Cells)
            {
                int currentRow = cell.Row;
                Range cellJ = cell.Worksheet.Cells[currentRow, 10]; // J 列为第10列

                // 公式：张数 * Q列当前行号 - K列当前行号
                cellJ.Formula = $@"={quantity}*Q{currentRow}-K{currentRow}";
            }
        }
        public static void PrintLabel(Range range)
        {
            if (range == null)
                return;
            Microsoft.Office.Interop.Excel.Application app = null;
            Workbook workbook = null;
            Worksheet sourceSheet = null;
            Worksheet labelSheet = null;
            try
            {
                app = range.Application;
                sourceSheet = range.Worksheet as Worksheet;
                if (sourceSheet == null || sourceSheet.Name != "打包记录")
                {
                    app.StatusBar = "请选择【打包记录】工作表中的记录。";
                    return;
                }
                workbook = sourceSheet.Parent as Workbook;
                labelSheet = workbook.Worksheets["标签"] as Worksheet;

                if (labelSheet == null)
                {
                    app.StatusBar = "未找到【标签】工作表。";
                    return;
                }

                foreach (Range rowRange in range.Rows)
                {
                    try
                    {
                        int row = rowRange.Row;
                        bool hasEmptyCell = false;
                        // 检查第2~16列
                        for (int col = 2; col <= 16; col++)
                        {
                            Range cell = null;

                            try
                            {
                                cell = sourceSheet.Cells[row, col] as Range;

                                string value = Convert.ToString(cell?.Value2)?.Trim();

                                if (string.IsNullOrEmpty(value))
                                {
                                    hasEmptyCell = true;
                                    break;
                                }
                            }
                            finally
                            {
                                if (cell != null)
                                    Marshal.ReleaseComObject(cell);
                            }
                        }

                        if (hasEmptyCell)
                        {
                            //MessageBox.Show($"第 {row} 行存在空值，已跳过。");
                            continue;
                        }

                        // 读取I列(第9列)
                        Range sourceCell = null;
                        Range targetCell = null;

                        try
                        {
                            sourceCell = sourceSheet.Cells[row, 9] as Range;
                            targetCell = labelSheet.Range["D2"];

                            targetCell.Value2 = sourceCell?.Value2;
                        }
                        finally
                        {
                            if (sourceCell != null)
                                Marshal.ReleaseComObject(sourceCell);

                            if (targetCell != null)
                                Marshal.ReleaseComObject(targetCell);
                        }

                        // 重新计算
                        app.Calculate();
                        Thread.Sleep(1000);
                        try
                        {
                            // 打印2份
                            labelSheet.PrintOut(
                                Copies: 2,
                                Collate: true
                            );

                            //app.StatusBar = $"第 {row} 行标签打印成功。";
                        }
                        catch
                        {
                            //app.StatusBar = $"第 {row} 行标签打印已取消。";
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(rowRange);
                    }
                }
            }
            catch (Exception ex)
            {
                if (app != null)
                {
                    app.StatusBar = $"打印失败：{ex.Message}";
                }
            }
            finally
            {
                if (labelSheet != null)
                    Marshal.ReleaseComObject(labelSheet);

                if (sourceSheet != null)
                    Marshal.ReleaseComObject(sourceSheet);

                if (workbook != null)
                    Marshal.ReleaseComObject(workbook);

                // 不释放 app（Excel主实例）
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // Helper to release COM objects safely
        private static void ReleaseComObjectSafe(object comObj)
        {
            if (comObj == null) return;
            try
            {
                if (Marshal.IsComObject(comObj))
                {
                    while (Marshal.ReleaseComObject(comObj) > 0) { }
                }
            }
            catch
            {
                // swallow exceptions on release
            }
            finally
            {
                comObj = null;
            }
        }

        // Helper to force GC to finalize released COM objects
        private static void ForceRelease()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public static string StartNesting(System.Data.DataTable dataTable, string contractNumber, string folder)
        {
            // 1. 路径定义为普通字符串（不含双引号字符）
            string runAsDatePath = @"C:\Program Files\Siemens\Solid Edge 2D Nesting 2025\RunAsDate.exe";
            string nestingExePath = @"C:\Program Files\Siemens\Solid Edge 2D Nesting 2025\se2dnest.exe";

            // 2. 时间字符串（移除前导/尾随空格，后续会统一加引号）
            string time = "10/10/2025 12:00:00";

            // 3. 导出 CSV 文件
            string csvPath = FileHelper.ExportDataTableToCsv(dataTable,
                contractNumber + "-" + DateTime.Now.ToString("yyyyMMdd-HHmmss"), folder);

            // 4. 为可能包含空格的路径和时间值加上双引号
            //string quotedTime = $"\"{time}\"";
            string quotedNestingExe = $"\"{nestingExePath}\"";
            string quotedCsv = $"\"{csvPath}\"";

            // 5. 构造 RunAsDate 的参数：
            //    格式：RunAsDate.exe "时间" "目标程序路径" 目标程序参数
            string arguments = $"{time} {quotedNestingExe} -i {quotedCsv} -a -c 36000:99";
            //MessageBox.Show(arguments);
            // 6. 启动进程
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = runAsDatePath,          // 可直接使用带空格的路径，不需额外引号
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = true               // 建议隐藏黑色窗口
            };
            Process process = Process.Start(startInfo);
            // 如果需要等待排样完成，可取消下一行注释
            // process.WaitForExit();
            return csvPath;
        }

        /// <summary>
        /// 跳转订单
        /// </summary>
        public static void SkipOrder(Workbook wb, Worksheet orderSheet, Range selectionRange)
        {
            if (wb == null) return;
            Worksheet sourceSheet = null;
            Range contractNumberRange = null;
            Range materialNumberRange = null;
            Range materialHeader = null;
            Range findRange = null;
            try
            {
                sourceSheet = selectionRange.Worksheet as Worksheet;
                int activeRow = wb.Application.ActiveCell.Row;
                contractNumberRange = sourceSheet.Cells.Find(
                    What: "合同编号",
                    LookAt: Microsoft.Office.Interop.Excel.XlLookAt.xlWhole,
                    LookIn: Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                    SearchOrder: Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                    SearchDirection: Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext,
                    MatchCase: false,
                    SearchFormat: false);
                materialNumberRange = sourceSheet.Cells.Find(
                    What: "物料编码",
                    LookAt: Microsoft.Office.Interop.Excel.XlLookAt.xlWhole,
                    LookIn: Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                    SearchOrder: Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                    SearchDirection: Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext,
                    MatchCase: false,
                    SearchFormat: false);

                if (contractNumberRange == null || materialNumberRange == null)
                    throw new Exception("未找到合同编号或物料编码列。");

                string contractNumber = sourceSheet.Cells[activeRow, contractNumberRange.Column].value2?.ToString();
                string materialNumber = sourceSheet.Cells[activeRow, materialNumberRange.Column].value2?.ToString();

                if (string.IsNullOrWhiteSpace(contractNumber) || string.IsNullOrWhiteSpace(materialNumber)) return;
                
                if (orderSheet.AutoFilterMode)
                {
                    if (orderSheet.FilterMode)
                    {
                        orderSheet.AutoFilter.ShowAllData();
                    }
                }

                materialHeader = orderSheet.Cells.Find(
                    What: "物料编码",
                    LookAt: Microsoft.Office.Interop.Excel.XlLookAt.xlWhole);
                if (materialHeader == null) return;

                long materialNumberColumn = materialHeader.Column;

                findRange = orderSheet.Cells.Find(
                    What: contractNumber,
                    LookAt: Microsoft.Office.Interop.Excel.XlLookAt.xlWhole);
                if (findRange == null) return;

                string findAddress = findRange.Address;
                do
                {
                    string cellValue = (orderSheet.Cells[findRange.Row, materialNumberColumn].value2?.ToString()) ?? "";
                    if (string.Equals(materialNumber ?? "", cellValue))
                    {
                        orderSheet.Activate();
                        wb.Application.Goto(orderSheet.Rows[findRange.Row], Type.Missing);
                        return;
                    }
                    Range next = orderSheet.Cells.FindNext(findRange);
                    ReleaseComObjectSafe(findRange);
                    findRange = next;
                } while (findRange != null && findAddress != findRange.Address);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"SkipOrder 出错：{ex.Message}", "错误");
            }
            finally
            {
                ReleaseComObjectSafe(findRange);
                ReleaseComObjectSafe(materialHeader);
                ReleaseComObjectSafe(materialNumberRange);
                ReleaseComObjectSafe(contractNumberRange);
                ReleaseComObjectSafe(sourceSheet);
                ReleaseComObjectSafe(orderSheet);
                ForceRelease();
            }
        }

        /// <summary>
        /// 计算打包数量
        /// </summary>
        public static void CalcPackage(Microsoft.Office.Interop.Excel.Application app)
        {
            Range selectedRange = null;
            Worksheet worksheet = null;
            Worksheet ordersheet = null;
            Worksheet packagesheet = null;
            Range findResult = null;

            try
            {
                selectedRange = app.Selection as Range;

                if (selectedRange == null || (selectedRange.Worksheet != null && selectedRange.Worksheet.Name != "生产日报"))
                    return;

                worksheet = app.Worksheets["生产日报"] as Worksheet;
                ordersheet = app.Worksheets["订单数量"] as Worksheet;
                packagesheet = app.Worksheets["打包记录"] as Worksheet;
                if (worksheet == null || ordersheet == null)
                {
                    MessageBox.Show("未找到“生产日报”或“订单数量”工作表。", "错误");
                    return;
                }
                // 生产日报列
                findResult = worksheet.Cells.Find("日期");
                if (findResult == null) { MessageBox.Show("未找到“日期”列。"); return; }
                int dateColumn = findResult.Column;
                ReleaseComObjectSafe(findResult);

                findResult = worksheet.Cells.Find("班");
                if (findResult == null) { MessageBox.Show("未找到“班”列。"); return; }
                int shiftColumn = findResult.Column;
                ReleaseComObjectSafe(findResult);

                findResult = worksheet.Cells.Find("合同编号");
                if (findResult == null) { MessageBox.Show("未找到“合同编号”列。"); return; }
                int contractNumberColumn = findResult.Column;
                ReleaseComObjectSafe(findResult);

                findResult = worksheet.Cells.Find("物料编码");
                if (findResult == null) { MessageBox.Show("未找到“物料编码”列。"); return; }
                int materialNumberColumn = findResult.Column;
                ReleaseComObjectSafe(findResult);

                findResult = worksheet.Cells.Find("机台*编号");
                if (findResult == null) { MessageBox.Show("未找到“机台*编号”列。"); return; }
                int machineNumberColumn = findResult.Column;
                ReleaseComObjectSafe(findResult);

                findResult = worksheet.Cells.Find("合格*数量");
                if (findResult == null) { MessageBox.Show("未找到“合格*数量”列。"); return; }
                int quantityColumn = findResult.Column;
                ReleaseComObjectSafe(findResult);

                findResult = worksheet.Cells.Find("未打包数");
                if (findResult == null) { MessageBox.Show("未找到“未打包数”列。"); return; }
                int orderPackageQtyColumn = findResult.Column;
                ReleaseComObjectSafe(findResult);

                findResult = worksheet.Cells.Find("单张*数量");
                if (findResult == null) { MessageBox.Show("未找到“单张*数量”列。"); return; }
                int danShuColumn = findResult.Column;
                ReleaseComObjectSafe(findResult);

                findResult = worksheet.Cells.Find("名称");
                if (findResult == null) { MessageBox.Show("未找到“名称”列。"); return; }
                int PNameColumn = findResult.Column;
                ReleaseComObjectSafe(findResult);

                findResult = worksheet.Cells.Find("规格");
                if (findResult == null) { MessageBox.Show("未找到“规格”列。"); return; }
                int PSpecColumn = findResult.Column;
                ReleaseComObjectSafe(findResult);

                // 订单数量表列
                findResult = ordersheet.Cells.Find("合同编号");
                if (findResult == null) { MessageBox.Show("订单表中未找到“合同编号”列。"); return; }
                int orderContractNumberColumn = findResult.Column;
                ReleaseComObjectSafe(findResult);

                findResult = ordersheet.Cells.Find("物料编码");
                if (findResult == null) { MessageBox.Show("订单表中未找到“物料编码”列。"); return; }
                int orderMaterialNumberColumn = findResult.Column;
                ReleaseComObjectSafe(findResult);

                findResult = ordersheet.Cells.Find("待生产数量");
                if (findResult == null) { MessageBox.Show("订单表中未找到“待生产数量”列。"); return; }
                int orderQuantityColumn = findResult.Column;
                ReleaseComObjectSafe(findResult);

                findResult = ordersheet.Cells.Find("码托*数量");
                if (findResult == null) { MessageBox.Show("订单表中未找到“码托*数量”列。"); return; }
                int packageQtyColumn = findResult.Column;
                ReleaseComObjectSafe(findResult);

                // 收集选中行号（去重）
                var rowNumbers = new HashSet<int>();
                foreach (Range area in selectedRange.Areas)
                {
                    foreach (Range cell in area.Cells)
                    {
                        rowNumbers.Add(cell.Row);
                        ReleaseComObjectSafe(cell);
                    }
                    ReleaseComObjectSafe(area);
                }

                // 构建DataTable
                System.Data.DataTable data = new System.Data.DataTable("CalcPackage");
                data.Columns.Add("序号", typeof(int));
                data.Columns.Add("机台编号", typeof(string));
                data.Columns.Add("合同编号", typeof(string));
                data.Columns.Add("物料编码", typeof(string));
                data.Columns.Add("零件名称", typeof(string));
                data.Columns.Add("零件规格", typeof(string));
                data.Columns.Add("打包数量", typeof(int));
                data.Columns.Add("截至张数", typeof(int));
                data.Columns.Add("备注", typeof(string));
                data.Columns.Add("完成张数", typeof(int));

                int count = 0;
                int beginQty = 0;
                string input = Interaction.InputBox("请输入起始张数");
                if (string.IsNullOrEmpty(input) || !int.TryParse(input, out beginQty))
                    return;

                // 为提高性能，将需要的工作表列数据一次读入数组（按行索引）
                var colValues = new Dictionary<int, object[]>();
                int[] neededCols = { contractNumberColumn, materialNumberColumn, orderPackageQtyColumn,
                                 danShuColumn, machineNumberColumn, PNameColumn, PSpecColumn };
                foreach (int col in neededCols)
                {
                    Range usedRange = null;
                    Range colRange = null;
                    try
                    {
                        usedRange = worksheet.UsedRange;
                        int lastRow = usedRange.Row + usedRange.Rows.Count - 1;
                        colRange = worksheet.Range[worksheet.Cells[1, col], worksheet.Cells[lastRow, col]];
                        object[,] vals = colRange.Value2 as object[,];
                        if (vals == null)
                        {
                            object single = colRange.Value2;
                            vals = new object[lastRow + 1, 1 + 1];
                            vals[1, 1] = single;
                        }
                        colValues[col] = new object[lastRow + 1]; // 索引即行号，0号元素不用
                        for (int r = 1; r <= lastRow; r++)
                        {
                            object cellVal = null;
                            try { cellVal = vals[r, 1]; } catch { cellVal = null; }
                            colValues[col][r] = cellVal;
                        }
                    }
                    finally
                    {
                        ReleaseComObjectSafe(colRange);
                        ReleaseComObjectSafe(usedRange);
                    }
                }

                // 缓存订单表的整列Range引用，用于SumIfs
                Range orderQtyRange = ordersheet.Columns[orderQuantityColumn];
                Range orderContractRange = ordersheet.Columns[orderContractNumberColumn];
                Range orderMaterialRange = ordersheet.Columns[orderMaterialNumberColumn];
                Range orderPackageRange = ordersheet.Columns[packageQtyColumn];

                foreach (int rowIndex in rowNumbers)
                {
                    count++;
                    object contractObj = colValues[contractNumberColumn][rowIndex];
                    object materialObj = colValues[materialNumberColumn][rowIndex];
                    object qtyObj = colValues[orderPackageQtyColumn][rowIndex];
                    object singleQtyObj = colValues[danShuColumn][rowIndex];
                    object machineObj = colValues[machineNumberColumn][rowIndex];
                    object nameObj = colValues[PNameColumn][rowIndex];
                    object specObj = colValues[PSpecColumn][rowIndex];

                    string contractNumber = contractObj?.ToString() ?? "";
                    string materialNumber = materialObj?.ToString() ?? "";

                    int qty = 0, singleQty = 0;
                    if (!int.TryParse(qtyObj?.ToString(), out qty)) qty = 0;
                    if (!int.TryParse(singleQtyObj?.ToString(), out singleQty)) singleQty = 0;

                    int orderQty = 0;
                    int packageQty = 0;
                    try
                    {
                        object sumResult = app.WorksheetFunction.SumIfs(
                            orderQtyRange, orderContractRange, contractNumber,
                            orderMaterialRange, materialNumber);
                        orderQty = Convert.ToInt32(sumResult);
                    }
                    catch { orderQty = 0; }

                    try
                    {
                        object sumResult = app.WorksheetFunction.SumIfs(
                            orderPackageRange, orderContractRange, contractNumber,
                            orderMaterialRange, materialNumber);
                        packageQty = Convert.ToInt32(sumResult);
                    }
                    catch { packageQty = 0; }

                    int remaining = orderQty + qty;
                    int pending = qty;
                    int totalNum = beginQty;

                    do
                    {
                        int PQty = 0;
                        string note = "";
                        if (packageQty == 0 || orderQty == 0 || singleQty == 0)
                        {
                            note = "数据缺失";
                            remaining = 0;
                            totalNum = 0;
                        }
                        else
                        {
                            int target = (remaining <= packageQty) ? remaining : packageQty;
                            int need = target - pending;
                            if (need < 0) need = 0;
                            int addNum = (need + singleQty - 1) / singleQty;
                            totalNum += addNum;
                            int totalAvailable = pending + addNum * singleQty;
                            bool isLastBatch = (remaining <= packageQty);

                            if (isLastBatch)
                            {
                                PQty = totalAvailable;
                                if (PQty < remaining) PQty = remaining;
                                if (PQty > remaining + singleQty) PQty = remaining + singleQty;
                                note = "订单量完成";
                            }
                            else
                            {
                                PQty = totalAvailable;
                            }
                            remaining -= PQty;
                            pending = totalAvailable - PQty;
                        }

                        data.Rows.Add(count,
                            machineObj?.ToString() ?? "",
                            contractNumber,
                            materialNumber,
                            nameObj?.ToString() ?? "",
                            specObj?.ToString() ?? "",
                            PQty,
                            totalNum,
                            note,
                            totalNum - beginQty);

                    } while (remaining > 0);
                }

                //截至张数排序
                // 使用 DefaultView.Sort 对 DataTable 按“截至张数”列进行排序，然后生成新的 DataTable
                // 注意：列名必须与 data 中的列名完全相同，这里为 "截至张数"
                try
                {
                    data.DefaultView.Sort = "[截至张数] ASC";
                    data = data.DefaultView.ToTable();
                }
                catch
                {
                    // 如果排序失败（例如列名不匹配），忽略并继续使用原始顺序
                }

                // 将DataTable写入新工作表（一次性数组写入）
                Worksheet newSheet = app.Worksheets.Add(After: app.Worksheets[app.Worksheets.Count]);
                try
                {
                    newSheet.Name = "打包计算_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                    int rowCount = data.Rows.Count;
                    int colCount = data.Columns.Count;
                    object[,] outputArray = new object[rowCount + 1, colCount]; // 包含标题行

                    for (int c = 0; c < colCount; c++)
                    {
                        outputArray[0, c] = data.Columns[c].ColumnName;
                    }
                    for (int r = 0; r < rowCount; r++)
                    {
                        for (int c = 0; c < colCount; c++)
                        {
                            object val = data.Rows[r][c];
                            outputArray[r + 1, c] = (val == DBNull.Value) ? "" : val;
                        }
                    }

                    Range writeRange = newSheet.Range[newSheet.Cells[1, 1], newSheet.Cells[rowCount + 1, colCount]];
                    try
                    {
                        writeRange.Value2 = outputArray;
                    }
                    finally
                    {
                        ReleaseComObjectSafe(writeRange);
                    }
                }
                finally
                {
                    newSheet.Columns.AutoFit();
                    ReleaseComObjectSafe(newSheet);
                }

                // 释放列范围引用
                ReleaseComObjectSafe(orderQtyRange);
                ReleaseComObjectSafe(orderContractRange);
                ReleaseComObjectSafe(orderMaterialRange);
                ReleaseComObjectSafe(orderPackageRange);

            }
            catch (Exception ex)
            {
                MessageBox.Show($"CalcPackage 发生错误：{ex.Message}", "错误");
            }
            finally
            {
                ReleaseComObjectSafe(findResult);
                ReleaseComObjectSafe(packagesheet);
                ReleaseComObjectSafe(ordersheet);
                ReleaseComObjectSafe(worksheet);
                ReleaseComObjectSafe(selectedRange);
                ForceRelease();
            }
        }

    }
}
