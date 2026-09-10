using Microsoft.Office.Interop.Excel;

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using 激光报表;

namespace 报表工具
{
    public partial class ToolsRibbon
    {
        Workbook wb;
        Worksheet productSheet;
        Worksheet orderSheet;
        private void ToolsRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            checkBox1.Checked = false;
        }
        private void UnifiedClickHandler(object sender, RibbonControlEventArgs e)
        {
            RibbonControl ctrl = sender as RibbonControl;
            wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (wb == null || wb.Name != "激光报表.xlsx") {
                MessageBox.Show("请确保已打开并激活名为 '激光报表.xlsx' 的工作簿。", "提示");
                return;
            }
            productSheet = wb.Worksheets["产品资料"];
            orderSheet = wb.Worksheets["订单数量"];
            Range selectionRange = wb.Application.Selection as Range;
            List<Product> products = ToolsHelper.GetProducts(selectionRange);

                switch (ctrl.Id)
                {
                    //转到订单
                    case "button1":
                        ExcelHelper.SkipOrder(wb, orderSheet, selectionRange);
                        break;
                    //打包计算
                    case "button2":
                        ExcelHelper.CalcPackage(Globals.ThisAddIn.Application);
                        break;
                    //3启动套料
                    case "button3":
                        string contractNumber = string.Empty;
                        Range usedRange;
                        //获取合同编号
                        contractNumber = ToolsHelper.GetCellValueByColumnHeader(selectionRange, "合同编号");
                        //获取订单数量表数据
                        usedRange = orderSheet.UsedRange;
                        if (usedRange == null || usedRange.Rows.Count < 1 || contractNumber == string.Empty) return;
                        // 将UsedRange的值转换为二维数组，索引从1开始（行列）
                        object[,] orderSheetArray = usedRange.Value2 as object[,];
                        //NestingFrom nestingFrom = new NestingFrom(contractNumber, orderSheetArray, product, selectionRange);
                        NestingFrom nestingFrom = new NestingFrom(products);
                        nestingFrom.ShowDialog();
                        break;
                    //标签打印
                    case "button4":
                        ExcelHelper.PrintLabel(selectionRange);
                        break;
                    //生产日报数量填充
                    case "button5":
                        // 如果未选中任何单元格，或选中的不是Range对象，则退出
                        if (selectionRange == null) return;
                        // 判断选中的是否为单独一行数据
                        // 注意：Rows.Count 返回的是选中区域的行数（包含所有区域的行数，适用于多区域选中）
                        if (selectionRange.Rows.Count != 1)
                        {
                            System.Windows.Forms.MessageBox.Show("请选择单独一行数据。", "提示");
                            return;
                        }
                        // 调用之前定义的处理函数，传入选中的区域
                        ExcelHelper.ProcessProductionDaily(selectionRange);
                        break;
                    //打包
                    case "button6":
                        selectionRange = Globals.ThisAddIn.Application.Selection as Range;
                        if (selectionRange.Worksheet.Name == "生产日报")
                        {
                            ExcelHelper.CopyPack(selectionRange, wb.Worksheets["打包记录"]);
                        }
                        break;
                    //材料计算
                    case "button7":
                        System.Data.DataTable dataTable = new System.Data.DataTable();

                        dataTable.Columns.Add("合同编号", typeof(string));
                        dataTable.Columns.Add("物料编码", typeof(string));
                        dataTable.Columns.Add("规格", typeof(string));
                        dataTable.Columns.Add("名称", typeof(string));
                        dataTable.Columns.Add("数量", typeof(int));
                        dataTable.Columns.Add("长度", typeof(double));
                        dataTable.Columns.Add("宽度", typeof(double));
                        dataTable.Columns.Add("厚度", typeof(double));
                        dataTable.Columns.Add("重量", typeof(double));
                        foreach (var item in products)
                        {
                            //MessageBox.Show($"合同编号: {item.ContractNumber}, 名称: {item.Name}, 数量: {item.Quantity}");
                            dataTable.Rows.Add(item.ContractNumber, item.Code, item.Specification, item.Name, item.Quantity, item.Length, item.Width, item.Thickness, item.Weight);
                        }
                        // 生成时间戳，格式 yyyyMMddHHmmss
                        string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                        string sheetName = $"材料计算_{timestamp}";

                        // 1. 在最后添加新工作表
                        Worksheet lastSheet = wb.Worksheets[wb.Worksheets.Count];
                        Worksheet newSheet = wb.Worksheets.Add(After: lastSheet);
                        newSheet.Name = sheetName;

                        // 2. 将 DataTable 写入工作表（高效数组方式）
                        if (dataTable != null && dataTable.Columns.Count > 0)
                        {
                            int rowCount = dataTable.Rows.Count;
                            int colCount = dataTable.Columns.Count;
                            object[,] data = new object[rowCount + 1, colCount];
                            for (int c = 0; c < colCount; c++)
                                data[0, c] = dataTable.Columns[c].ColumnName;
                            for (int r = 0; r < rowCount; r++)
                                for (int c = 0; c < colCount; c++)
                                    data[r + 1, c] = dataTable.Rows[r][c];
                            Range startCell = newSheet.Cells[1, 1];
                            Range endCell = newSheet.Cells[rowCount + 1, colCount];
                            Range range = newSheet.Range[startCell, endCell];
                            range.Value = data;
                            Marshal.ReleaseComObject(range);
                            Marshal.ReleaseComObject(endCell);
                            Marshal.ReleaseComObject(startCell);
                        }
                        else
                        {
                            // 如果表没有列，只写一个提示（可选）
                            newSheet.Cells[1, 1] = "无数据";
                        }

                        usedRange = newSheet.UsedRange;
                        int lastCol = usedRange.Columns.Count + 1;
                        newSheet.Cells[1, lastCol] = "总重量T";
                        newSheet.Cells[1, lastCol + 1] = "总面积";
                        newSheet.Cells[1, lastCol + 2] = "1250x2750张数";
                        for (int i = 2; i <= usedRange.Rows.Count; i++)
                        {
                            newSheet.Cells[i, lastCol].Formula = $"=I{i}*E{i}/1000";
                            newSheet.Cells[i, lastCol + 1].Formula = $"=(F{i}+4)*(G{i}+4)*E{i}";
                            newSheet.Cells[i, lastCol + 2].Formula = $"=K{i}/(1250*2750)";
                        }
                        int lastRow = usedRange.Rows.Count + 1;
                        newSheet.Cells[lastRow, lastCol].Formula = $"=SUM({newSheet.Cells[2, lastCol].Address}:{newSheet.Cells[lastRow - 1, lastCol].Address})";
                        newSheet.Cells[lastRow, lastCol + 1].Formula = $"=SUM({newSheet.Cells[2, lastCol + 1].Address}:{newSheet.Cells[lastRow - 1, lastCol + 1].Address})";
                        newSheet.Cells[lastRow, lastCol + 2].Formula = $"=SUM({newSheet.Cells[2, lastCol + 2].Address}:{newSheet.Cells[lastRow - 1, lastCol + 2].Address})";
                        // 3. 自动调整列宽
                        newSheet.Columns.AutoFit();
                        Range rowRange = newSheet.Rows[lastRow];
                        rowRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(244, 179, 130));
                        // 4. 释放工作表引用
                        Marshal.ReleaseComObject(rowRange);
                        Marshal.ReleaseComObject(newSheet);
                        Marshal.ReleaseComObject(lastSheet);
                        break;
                    default:
                        Marshal.ReleaseComObject(productSheet);
                        Marshal.ReleaseComObject(orderSheet);
                        Marshal.ReleaseComObject(selectionRange);
                        break;
                }
        }
    }
}
