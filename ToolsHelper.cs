
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using 报表工具;
using Excel = Microsoft.Office.Interop.Excel;
namespace 激光报表
{
    internal class ToolsHelper
    {
        public static DataTable CreateCsvTable(List<Product> products)
        {
            if (products == null || products.Count == 0)
                throw new ArgumentException("产品列表不能为空", nameof(products));
            DataTable dataTable = new System.Data.DataTable();
            var columns = new[]
            {
                new DataColumn("Type",            typeof(string)),
                new DataColumn("Name",            typeof(string)),
                new DataColumn("Value1",          typeof(string)),
                new DataColumn("Value2",          typeof(string)),
                new DataColumn("Quantity",        typeof(int)),
                new DataColumn("Rotation",        typeof(int)),
                new DataColumn("Priority",        typeof(int)),
                new DataColumn("IsSheet",         typeof(int)),
                new DataColumn("MaterialNumber",  typeof(string)),
                new DataColumn("T",               typeof(double))
            };
            dataTable.Columns.AddRange(columns);
            foreach (var product in products)
            {
                DataRow newRow = dataTable.NewRow();
                newRow["Type"] = string.Equals(product.Path, "Rectangle", StringComparison.OrdinalIgnoreCase) ? "Rectangle" : "File";
                newRow["Name"] = $"{product.ContractNumber}-{product.Code}-{product.Name}";
                newRow["Value1"] = string.Equals(product.Path, "Rectangle", StringComparison.OrdinalIgnoreCase) ? product.Length.ToString() : product.Path;
                newRow["Value2"] = string.Equals(product.Path, "Rectangle", StringComparison.OrdinalIgnoreCase) ? product.Width.ToString() : string.Empty;
                newRow["Quantity"] = product.Quantity;
                newRow["Rotation"] = product.Rotation;
                newRow["Priority"] = 2;
                newRow["IsSheet"] = 0;
                newRow["MaterialNumber"] = product.Code;
                newRow["T"] = product.Thickness;
                dataTable.Rows.Add(newRow);
            }
            return dataTable;
        }

        /// <summary>
        /// 将工件数据写入DataTable
        /// </summary>
        /// <param name="filteredData">工件信息</param>
        /// <param name="productSheet">产品表对象</param>
        /// <returns></returns>
        public static DataTable CteateCsvTable(List<object[]> filteredData, Excel.Worksheet productSheet)
        {
            DataTable dataTable = new System.Data.DataTable();
            dataTable.Columns.Add("Type", typeof(string));
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Value1", typeof(string));
            dataTable.Columns.Add("Value2", typeof(string));
            dataTable.Columns.Add("Quantity", typeof(int));
            dataTable.Columns.Add("Rotation", typeof(int));
            dataTable.Columns.Add("Priority", typeof(int));
            dataTable.Columns.Add("IsSheet", typeof(int));
            dataTable.Columns.Add("MaterialNumber", typeof(string));
            dataTable.Columns.Add("T", typeof(double));
            // 获取标题行
            object[] headers = filteredData[0];
            int contractCol = Array.FindIndex(headers, h => h?.ToString() == "合同编号");
            int contractCol1 = Array.FindIndex(headers, h => h?.ToString() == "物料编码");
            int contractCol2 = Array.FindIndex(headers, h => h?.ToString() == "名称");
            int contractCol3 = Array.FindIndex(headers, h => h?.ToString() == "规格");
            int contractCol4 = Array.FindIndex(headers, h => h?.ToString() == "数量");
            int contractCol5 = Array.FindIndex(headers, h => h?.ToString() == "待生产数量");
            //写入工件数据到dataTable
            for (int i = 1; i < filteredData.Count; i++)
            {
                object[] row = filteredData[i];
                string path = ToolsHelper.XLookupByColumnName(productSheet, row[contractCol1]?.ToString(), "物料编码", "图纸目录");
                string strRotation = ToolsHelper.XLookupByColumnName(productSheet, row[contractCol1]?.ToString(), "物料编码", "旋转");
                string t = ToolsHelper.XLookupByColumnName(productSheet, row[contractCol1]?.ToString(), "物料编码", "厚度\nmm");

                int rotation = -1;
                int quantity = int.Parse(row[contractCol5]?.ToString());
                if (strRotation != null && strRotation != "") rotation = int.Parse(ToolsHelper.XLookupByColumnName(productSheet, row[contractCol1]?.ToString(), "物料编码", "旋转"));
                if (quantity < 0) quantity = 0;
                if (path == null || path == string.Empty) MessageBox.Show("在查找" + row[contractCol1]?.ToString() + "的图纸目录时返回空。");

                //判断是否是矩形
                if (string.Equals(path, "Rectangle", StringComparison.OrdinalIgnoreCase))
                {
                    string value1 = ToolsHelper.XLookupByColumnName(productSheet, row[contractCol1]?.ToString(), "物料编码", "长度mm");
                    string value2 = ToolsHelper.XLookupByColumnName(productSheet, row[contractCol1]?.ToString(), "物料编码", "宽度mm");
                    //dataTable.Rows.Add("Rectangle", row[contractCol2]?.ToString(), value1, value2, quantity, rotation, 2, 0, row[contractCol1]?.ToString(), t);
                    DataRow newRow = dataTable.NewRow();
                    newRow["Type"] = "Rectangle";
                    newRow["Name"] = row[contractCol]?.ToString() + "-" + row[contractCol1]?.ToString() + "-" + row[contractCol2]?.ToString();
                    newRow["Value1"] = value1;
                    newRow["Value2"] = value2;
                    newRow["Quantity"] = quantity;
                    newRow["Rotation"] = rotation;
                    newRow["Priority"] = 2;
                    newRow["IsSheet"] = 0;
                    newRow["MaterialNumber"] = row[contractCol1]?.ToString();
                    newRow["T"] = t;
                    dataTable.Rows.Add(newRow);
                }
                else
                {
                    //dataTable.Rows.Add("File", row[contractCol2]?.ToString(), path, "", quantity, rotation, 2, 0, row[contractCol1]?.ToString(), t);
                    DataRow newRow = dataTable.NewRow();
                    newRow["Type"] = "File";
                    newRow["Name"] = row[contractCol]?.ToString() + "-" + row[contractCol1]?.ToString() + "-" + row[contractCol2]?.ToString();
                    newRow["Value1"] = path;
                    newRow["Value2"] = "";
                    newRow["Quantity"] = quantity;
                    newRow["Rotation"] = rotation;
                    newRow["Priority"] = 2;
                    newRow["IsSheet"] = 0;
                    newRow["MaterialNumber"] = row[contractCol1]?.ToString();
                    newRow["T"] = t;
                    dataTable.Rows.Add(newRow);
                }
            }
            return dataTable;
        }

        public class ContractMaterialData
        {
            public string ContractNumber { get; set; }
            public string MaterialCode { get; set; }
            public string MaterialName { get; set; }
            public string MaterialT { get; set; }

        }

        /// <summary>
        /// 创建包含指定模式和数据行的 DataTable。
        /// 写入板材数据
        /// </summary>
        /// <param name="width">Value2 列的固定宽度值</param>
        /// <param name="maxLength">Value1 列的最大长度（包含）</param>
        /// <param name="minLength">Value1 列的起始最小长度</param>
        /// <param name="interval">Value1 值的递增间隔</param>
        /// <returns>填充好数据的 DataTable</returns>
        public static DataTable CreateRectangleTable(int maxLength, int minLength, int maxWidth, int minWidth, int interval)
        {
            // 参数校验（按需）
            if (interval <= 0)
                throw new ArgumentException("间隔必须大于零", nameof(interval));
            DataTable table = new DataTable();

            // 定义列
            table.Columns.Add("Type", typeof(string));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Value1", typeof(String));
            table.Columns.Add("Value2", typeof(String));
            table.Columns.Add("Quantity", typeof(int));
            table.Columns.Add("Priority", typeof(int));
            table.Columns.Add("IsSheet", typeof(int));

            // 按间隔递增生成行
            //for (int val = minLength; val <= maxLength; val += interval)
            //{
            //    DataRow row = table.NewRow();
            //    row["Type"] = "Rectangle";
            //    row["Name"] = string.Empty;   // 或 DBNull.Value，这里用空字符串
            //    row["Value1"] = val;
            //    row["Value2"] = width;
            //    row["Quantity"] = 10000;
            //    row["Priority"] = 2;
            //    row["IsSheet"] = 1;
            //    table.Rows.Add(row);
            //}

            // 按间隔递增生成行
            for (int length = minLength; length <= maxLength; length += interval)
            {
                for (int width = minWidth; width <= maxWidth; width += interval)
                {
                    DataRow row = table.NewRow();
                    row["Type"] = "Rectangle";
                    row["Name"] = string.Empty;
                    row["Value1"] = length;
                    row["Value2"] = width;
                    row["Quantity"] = 10000;
                    row["Priority"] = 2;
                    row["IsSheet"] = 1;
                    table.Rows.Add(row);
                }
            }
            return table;
        }

        /// <summary>
        /// 在 range 所在工作表查找内容为 searchString 的单元格（作为列标题），
        /// 并返回 range 第一个单元格所在行、该列交叉处的文本。
        /// </summary>
        /// <param name="range">用于定位行的 Range（若含多个区域，仅使用第一个区域的第一个单元格）</param>
        /// <param name="searchString">要搜索的列标题字符串</param>
        /// <returns>交叉单元格的显示文本，若未找到列标题则返回空字符串</returns>
        public static string GetCellValueByColumnHeader(Excel.Range range, string searchString)
        {
            if (range == null)
                throw new ArgumentNullException(nameof(range));
            if (string.IsNullOrEmpty(searchString))
                return string.Empty;

            // 1. 获取工作表并锁定目标行（Range.Row 直接返回第一个区域的首行）
            Excel.Worksheet worksheet = range.Worksheet;
            int targetRow = range.Row;   // 多区域时自动取第一个区域的第一个单元格的行

            // 2. 在工作表内搜索列标题
            Excel.Range foundCell = null;
            Excel.Range allCells = null;
            try
            {
                allCells = worksheet.Cells;
                foundCell = allCells.Find(
                    searchString,
                    LookIn: Excel.XlFindLookIn.xlValues,   // 按值查找
                    LookAt: Excel.XlLookAt.xlWhole,         // 精确匹配整个单元格
                    MatchCase: false);                     // 不区分大小写（可按需调整）

                if (foundCell == null)
                {
                    // 没有找到匹配的列标题
                    return string.Empty;
                }

                int targetColumn = foundCell.Column;

                // 3. 取出目标行、目标列交叉处的值
                Excel.Range resultCell = worksheet.Cells[targetRow, targetColumn];
                string result = resultCell.Text ?? string.Empty;

                // 释放临时 Range 对象
                Marshal.ReleaseComObject(resultCell);
                return result;
            }
            finally
            {
                // 释放 COM 对象，防止内存泄漏
                if (foundCell != null) Marshal.ReleaseComObject(foundCell);
                if (allCells != null) Marshal.ReleaseComObject(allCells);
            }
        }

        /// <summary>
        /// 在Excel工作表中按列名执行查找，支持单条件或多条件匹配。
        /// 从最后一行向前扫描，返回最靠后的匹配行中 <paramref name="matchColumnName"/> 列的值。
        /// </summary>
        /// <param name="sheet">Excel工作表</param>
        /// <param name="lookupValue">主查找值（null 视为空字符串）</param>
        /// <param name="lookupColumnName">主查找列标题</param>
        /// <param name="matchColumnName">返回值所在列标题</param>
        /// <param name="extraLookupColumnNames">（可选）额外查找列标题数组</param>
        /// <param name="extraLookupValues">（可选）与额外列对应的查找值数组</param>
        /// <returns>找到则返回单元格字符串（null单元格返回 String.Empty），未找到返回 null</returns>
        public static string XLookupByColumnName(
            Excel.Worksheet sheet,
            string lookupValue,
            string lookupColumnName,
            string matchColumnName,
            string[] extraLookupColumnNames = null,
            string[] extraLookupValues = null)
        {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (string.IsNullOrEmpty(lookupColumnName)) throw new ArgumentException("查找列名不能为空");
            if (string.IsNullOrEmpty(matchColumnName)) throw new ArgumentException("匹配列名不能为空");

            // 将主查找值 null 统一处理为 string.Empty，保持与原先一致
            if (lookupValue == null) lookupValue = string.Empty;

            // ========== 单条件（原逻辑，未改动） ==========
            if (extraLookupColumnNames == null || extraLookupColumnNames.Length == 0)
            {
                int lookupColIndex = GetColumnIndexByHeader(sheet, lookupColumnName);
                int matchColIndex = GetColumnIndexByHeader(sheet, matchColumnName);

                if (lookupColIndex == -1)
                    throw new ArgumentException($"未找到列名 '{lookupColumnName}'");
                if (matchColIndex == -1)
                    throw new ArgumentException($"未找到列名 '{matchColumnName}'");

                int lastRow = GetLastDataRow(sheet);
                if (lastRow < 2) return null;

                Excel.Range lookupRange = sheet.Range[sheet.Cells[2, lookupColIndex], sheet.Cells[lastRow, lookupColIndex]];
                Excel.Range matchRange = sheet.Range[sheet.Cells[2, matchColIndex], sheet.Cells[lastRow, matchColIndex]];

                object[,] lookupArray = lookupRange.Value2 as object[,];
                object[,] matchArray = matchRange.Value2 as object[,];

                if (lookupArray == null || matchArray == null) return null;

                int rowCount = lookupArray.GetLength(0);
                if (rowCount != matchArray.GetLength(0))
                    throw new InvalidOperationException("查找列与匹配列的行数不一致");

                var dict = new Dictionary<string, object>(rowCount, StringComparer.OrdinalIgnoreCase);
                for (int i = rowCount; i >= 1; i--)
                {
                    object keyObj = lookupArray[i, 1];
                    string keyStr = keyObj == null ? string.Empty : Convert.ToString(keyObj, CultureInfo.InvariantCulture);
                    object value = matchArray[i, 1];
                    if (!dict.ContainsKey(keyStr))
                        dict.Add(keyStr, value);
                }

                if (dict.TryGetValue(lookupValue, out object matchedValue))
                    return matchedValue?.ToString() ?? string.Empty;
                return null;
            }

            // ========== 多条件（新增） ==========
            // 校验额外参数
            if (extraLookupValues == null)
                throw new ArgumentException("提供 extraLookupColumnNames 时必须同时提供 extraLookupValues");
            if (extraLookupColumnNames.Length != extraLookupValues.Length)
                throw new ArgumentException("extraLookupColumnNames 与 extraLookupValues 数量必须相等");

            // 组合全部列名与查找值
            var allColNames = new List<string> { lookupColumnName };
            allColNames.AddRange(extraLookupColumnNames);

            var allLookupValues = new List<string> { lookupValue };
            allLookupValues.AddRange(extraLookupValues.Select(v => v ?? string.Empty));

            // 获取所有列的索引
            int matchColIndexMulti = GetColumnIndexByHeader(sheet, matchColumnName);
            if (matchColIndexMulti == -1)
                throw new ArgumentException($"未找到列名 '{matchColumnName}'");

            var colIndexes = new List<int>(allColNames.Count);
            foreach (var colName in allColNames)
            {
                int idx = GetColumnIndexByHeader(sheet, colName);
                if (idx == -1)
                    throw new ArgumentException($"未找到列名 '{colName}'");
                colIndexes.Add(idx);
            }

            int lastRowMulti = GetLastDataRow(sheet);
            if (lastRowMulti < 2) return null;

            // 读取所有条件列 + 匹配列到数组
            var conditionArrays = new List<object[,]>(colIndexes.Count);
            foreach (int colIdx in colIndexes)
            {
                Excel.Range rng = sheet.Range[sheet.Cells[2, colIdx], sheet.Cells[lastRowMulti, colIdx]];
                object[,] arr = rng.Value2 as object[,];
                if (arr == null) return null;
                conditionArrays.Add(arr);
            }

            Excel.Range matchRangeMulti = sheet.Range[sheet.Cells[2, matchColIndexMulti], sheet.Cells[lastRowMulti, matchColIndexMulti]];
            object[,] matchArrayMulti = matchRangeMulti.Value2 as object[,];
            if (matchArrayMulti == null) return null;

            int rowCountMulti = conditionArrays[0].GetLength(0);

            // 从最后一行向前扫描，找到第一个（即最靠后的）完全匹配行
            for (int i = rowCountMulti; i >= 1; i--)
            {
                bool allMatch = true;
                for (int c = 0; c < conditionArrays.Count; c++)
                {
                    object cellValue = conditionArrays[c][i, 1];
                    string cellStr = cellValue == null ? string.Empty : Convert.ToString(cellValue, CultureInfo.InvariantCulture);
                    if (!string.Equals(cellStr, allLookupValues[c], StringComparison.OrdinalIgnoreCase))
                    {
                        allMatch = false;
                        break;
                    }
                }
                if (allMatch)
                {
                    object result = matchArrayMulti[i, 1];
                    return result?.ToString() ?? string.Empty;
                }
            }

            return null;
        }

        /// <summary>
        /// 根据标题文本获取列号（1基）
        /// </summary>
        private static int GetColumnIndexByHeader(Excel.Worksheet sheet, string headerText, int column = 1)
        {
            Excel.Range headerRow = sheet.Rows[1];
            Excel.Range found = headerRow.Find(
                What: headerText,
                LookIn: Excel.XlFindLookIn.xlValues,
                LookAt: Excel.XlLookAt.xlWhole,
                SearchOrder: Excel.XlSearchOrder.xlByColumns,
                SearchDirection: Excel.XlSearchDirection.xlNext,
                MatchCase: false,
                SearchFormat: false);

            if (found != null)
                return found.Column;

            // 备用：遍历第一行
            Excel.Range used = sheet.UsedRange;
            int colCount = used.Columns.Count;
            for (int col = 1; col <= colCount; col++)
            {
                Excel.Range cell = sheet.Cells[column, col];
                if (cell.Value2 != null && string.Equals(cell.Value2.ToString(), headerText, StringComparison.OrdinalIgnoreCase))
                    return col;
            }
            return -1;
        }

        /// <summary>
        /// 获取工作表数据最后一行（包含数据的最大行号）
        /// </summary>
        private static int GetLastDataRow(Excel.Worksheet sheet)
        {
            Excel.Range used = sheet.UsedRange;
            return used.Rows.Count + used.Row - 1;
        }
        /// <summary>
        /// 获取 DataTable 中列名"T"的所有不重复值，并在返回数组的第一个位置插入 "all"。
        /// </summary>
        /// <param name="table">源数据表</param>
        /// <returns>首位为 "all" 的唯一值数组，若列不存在则返回只包含 "all" 的数组</returns>
        public static object[] GetUniqueTValues(DataTable table)
        {
            if (table == null)
                throw new ArgumentNullException(nameof(table));
            // 不区分大小写查找列名"T"
            DataColumn column = table.Columns
                                    .Cast<DataColumn>()
                                    .FirstOrDefault(c => c.ColumnName.Equals("T", StringComparison.OrdinalIgnoreCase));
            // 提取唯一值（列不存在时使用空序列）
            object[] uniqueValues = column != null
                ? table.AsEnumerable()
                       .Select(row => row[column])
                       .Distinct()
                       .ToArray()
                : Array.Empty<object>();
            // 在数组开头插入 "all"
            object[] result = new object[uniqueValues.Length + 1];
            result[0] = "All";
            Array.Copy(uniqueValues, 0, result, 1, uniqueValues.Length);
            return result;
        }

        /// <summary>
        /// 若 matchValue 为 "All"（忽略大小写），直接返回原表；
        /// 否则比较每一行的 "T" 列，不相等则将 "Quantity" 列置为 0。
        /// 直接修改原表并返回。
        /// 筛选厚度
        /// </summary>
        /// <param name="table">包含 "T" 和 "Quantity" 列的数据表</param>
        /// <param name="matchValue">用来比较的目标值；若为 "All" 则不修改</param>
        /// <returns>修改后的原 DataTable</returns>
        public static DataTable ApplyQuantityRule(DataTable table, string matchValue)
        {
            if (table == null)
                throw new ArgumentNullException(nameof(table));

            // 如果 matchValue 是 "All"（不区分大小写），不修改直接返回
            if (string.Equals(matchValue, "All", StringComparison.OrdinalIgnoreCase))
                return table;

            // 查找列（不区分大小写）
            DataColumn tCol = table.Columns
                                  .Cast<DataColumn>()
                                  .FirstOrDefault(c => c.ColumnName.Equals("T", StringComparison.OrdinalIgnoreCase));
            DataColumn qtyCol = table.Columns
                                     .Cast<DataColumn>()
                                     .FirstOrDefault(c => c.ColumnName.Equals("Quantity", StringComparison.OrdinalIgnoreCase));

            if (tCol == null)
                throw new ArgumentException("DataTable 中缺少列 'T'。");
            if (qtyCol == null)
                throw new ArgumentException("DataTable 中缺少列 'Quantity'。");

            foreach (DataRow row in table.Rows)
            {
                string currentValue = row[tCol]?.ToString() ?? string.Empty;

                // 不匹配则 Quantity 置 0
                if (!string.Equals(currentValue, matchValue, StringComparison.OrdinalIgnoreCase))
                {
                    row[qtyCol] = 0;
                }
            }
            return table;
        }
        /// <summary>
        /// 获取选中区域的产品信息列表。每个产品包含合同编号、物料编码、名称、规格、图纸目录、待生产数量和旋转等属性。
        /// </summary>
        /// <param name="selectedRange"></param>
        /// <returns></returns>
        public static List<Product> GetProducts(Excel.Range selectedRange)
        {
            var products = new List<Product>();
            Excel.Workbook workbook = selectedRange.Worksheet.Parent as Excel.Workbook;
            Excel.Worksheet productSheet = workbook.Worksheets["产品资料"];
            Excel.Worksheet orderSheet = workbook.Worksheets["订单数量"];
            Excel.Worksheet worksheet = workbook.Worksheets["生产日报"];
            Excel.Worksheet ws = selectedRange.Worksheet;
            Excel.Range worksheetUsedRange;
            // 第一步：在工作表的 UsedRange 中查找“合同编号”和“物料编码”的列号
            Excel.Range usedRange = ws.UsedRange;
            int totalRows = usedRange.Rows.Count;
            int totalCols = usedRange.Columns.Count;
            int contractCol = -1;
            int materialCol = -1;
            // 如果当前工作表是“生产日报”，则使用传入的 worksheet 的 UsedRange，否则使用当前工作表的 UsedRange
            if (ws.Name == worksheet.Name)
            {
                worksheetUsedRange = usedRange;
            }
            else worksheetUsedRange = worksheet.UsedRange;

            object[,] data = worksheetUsedRange.Value2;

            for (int r = 1; r <= totalRows; r++)
            {
                for (int c = 1; c <= totalCols; c++)
                {
                    Excel.Range cell = usedRange.Cells[r, c];
                    string cellValue = cell.Value2?.ToString().Trim();
                    if (cellValue == "合同编号")
                        contractCol = c;
                    else if (cellValue == "物料编码")
                        materialCol = c;
                }
                if (contractCol != -1 && materialCol != -1) break;
            }
            if (contractCol == -1 || materialCol == -1)
            {
                System.Windows.Forms.MessageBox.Show("未找到列标题：“合同编号”或“物料编码”");
                return products;
            }
            // 第二步：收集选中区域中所有有效的行号（去重，并限制在 UsedRange 范围内）
            HashSet<int> selectedRows = new HashSet<int>();
            int usedFirstRow = usedRange.Row;
            int usedLastRow = usedRange.Row + usedRange.Rows.Count - 1;
            foreach (Excel.Range area in selectedRange.Areas)
            {
                int areaFirstRow = area.Row;
                int areaLastRow = area.Row + area.Rows.Count - 1;

                int startRow = System.Math.Max(areaFirstRow, usedFirstRow);
                int endRow = System.Math.Min(areaLastRow, usedLastRow);

                if (startRow <= endRow)
                {
                    for (int row = startRow; row <= endRow; row++)
                        selectedRows.Add(row);
                }
            }
            foreach (int rowNum in selectedRows)
            {
                Excel.Range contractCell = ws.Cells[rowNum, contractCol];
                Excel.Range materialCell = ws.Cells[rowNum, materialCol];

                if (contractCell.EntireRow.Hidden == false)
                {
                    string contractVal = contractCell.Value2?.ToString() ?? "";
                    string materialVal = materialCell.Value2?.ToString() ?? "";
                    //string machineQuantity;
                    //string machineNumber;

                    //int ToPackQty;
                    int quantityVal = int.Parse(ToolsHelper.XLookupByColumnName(
                            orderSheet, materialVal, "物料编码", "待生产数量", new[] { "合同编号" }, new[] { contractVal }));
                    //if (quantityVal > 0) {

                    //    int MachineColumn;
                    //    int ContractNumberColumn;
                    //    int CodeColumn;
                    //    int machineNumberColumn;

                    //    MachineColumn = GetColumnIndexByHeader(worksheet, "机台\n编号",4);
                    //    ContractNumberColumn = GetColumnIndexByHeader(worksheet, "合同编号",4);
                    //    CodeColumn = GetColumnIndexByHeader(worksheet, "物料编码", 4);
                    //    machineNumberColumn = GetColumnIndexByHeader(worksheet, "未打包数", 4);

                    //    //for (int row = data.GetLength(0); row >= 1; row--) {
                    //    //    //
                    //    //    if (contractVal == data[row, ContractNumberColumn].ToString() && materialVal == data[row,CodeColumn].ToString())
                    //    //    {
                    //    //        machineNumber = data[row, machineNumberColumn].ToString();
                    //    //    }
                    //    //}
                    //}else //ToPackQty = 0;
                    products.Add(new Product
                    {
                        ContractNumber = contractVal,
                        Name = ToolsHelper.XLookupByColumnName(productSheet, materialVal, "物料编码", "零件名称"),
                        Code = materialVal,
                        Specification = ToolsHelper.XLookupByColumnName(productSheet, materialVal, "物料编码", "零件规格"),
                        Path = ToolsHelper.XLookupByColumnName(productSheet, materialVal, "物料编码", "图纸目录"),
                        Quantity = quantityVal < 0 ? 0 : quantityVal,
                        Rotation = int.Parse(ToolsHelper.XLookupByColumnName(productSheet, materialVal, "物料编码", "旋转")),
                        Length = double.Parse(ToolsHelper.XLookupByColumnName(productSheet, materialVal, "物料编码", "长度mm")),
                        Thickness = double.Parse(ToolsHelper.XLookupByColumnName(productSheet, materialVal, "物料编码", "厚度\nmm")),
                        Width = double.Parse(ToolsHelper.XLookupByColumnName(productSheet, materialVal, "物料编码", "宽度mm")),
                        Area = double.Parse(ToolsHelper.XLookupByColumnName(productSheet, materialVal, "物料编码", "面积㎡")),
                        ToPackQty = 0
                    });
                }
            }
            return products;
        }

        /// <summary>
        /// 根据传入的选中区域，提取每一行中“合同编号”和“物料编码”列的数据
        /// </summary>
        /// <param name="selectedRange">用户选中的 Excel 区域（可以是多区域、整行、整列等）</param>
        /// <returns>包含合同编号和物料编码的列表，若未找到标题则返回空列表</returns>
        public static List<ContractMaterialData> GetContractAndMaterialFromSelection(Excel.Range selectedRange)
        {
            var result = new List<ContractMaterialData>();

            if (selectedRange == null)
                return result;

            Excel.Worksheet ws = selectedRange.Worksheet;
            if (ws == null)
                return result;

            // 第一步：在工作表的 UsedRange 中查找“合同编号”和“物料编码”的列号
            Excel.Range usedRange = ws.UsedRange;
            int totalRows = usedRange.Rows.Count;
            int totalCols = usedRange.Columns.Count;
            int contractCol = -1;
            int materialCol = -1;

            for (int r = 1; r <= totalRows; r++)
            {
                for (int c = 1; c <= totalCols; c++)
                {
                    Excel.Range cell = usedRange.Cells[r, c];
                    string cellValue = cell.Value2?.ToString().Trim();
                    if (cellValue == "合同编号")
                        contractCol = c;
                    else if (cellValue == "物料编码")
                        materialCol = c;
                }
                if (contractCol != -1 && materialCol != -1)
                    break;
            }

            if (contractCol == -1 || materialCol == -1)
            {
                System.Windows.Forms.MessageBox.Show("未找到列标题：“合同编号”或“物料编码”");
                return result;
            }

            // 第二步：收集选中区域中所有有效的行号（去重，并限制在 UsedRange 范围内）
            HashSet<int> selectedRows = new HashSet<int>();
            int usedFirstRow = usedRange.Row;
            int usedLastRow = usedRange.Row + usedRange.Rows.Count - 1;

            foreach (Excel.Range area in selectedRange.Areas)
            {
                int areaFirstRow = area.Row;
                int areaLastRow = area.Row + area.Rows.Count - 1;

                int startRow = System.Math.Max(areaFirstRow, usedFirstRow);
                int endRow = System.Math.Min(areaLastRow, usedLastRow);

                if (startRow <= endRow)
                {
                    for (int row = startRow; row <= endRow; row++)
                        selectedRows.Add(row);
                }
            }

            // 第三步：提取每行的合同编号和物料编码
            foreach (int rowNum in selectedRows)
            {
                Excel.Range contractCell = ws.Cells[rowNum, contractCol];
                Excel.Range materialCell = ws.Cells[rowNum, materialCol];

                if (contractCell.EntireRow.Hidden == false)
                {
                    string contractVal = contractCell.Value2?.ToString() ?? "";
                    string materialVal = materialCell.Value2?.ToString() ?? "";
                    result.Add(new ContractMaterialData
                    {
                        ContractNumber = contractVal,
                        MaterialCode = materialVal
                    });
                }
            }
            return result;
        }
    }
}

