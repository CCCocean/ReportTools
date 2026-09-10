using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace 激光报表
{
    internal class FileHelper
    {
        /// <summary>
        /// 将 DataTable 导出为 CSV 文件，返回生成文件的完整路径
        /// </summary>
        /// <param name="dt">数据源 DataTable</param>
        /// <param name="fileName">输出文件名（不含后缀）</param>
        /// <param name="folderPath">文件保存文件夹路径</param>
        /// <returns>成功返回完整文件路径；失败返回 null</returns>
        public static string ExportDataTableToCsv(DataTable dt, string fileName, string folderPath)
        {
            try
            {
                // 参数校验
                if (dt == null)
                    throw new ArgumentNullException(nameof(dt), "DataTable 不能为 null");
                if (string.IsNullOrWhiteSpace(fileName))
                    throw new ArgumentException("文件名不能为空或空白", nameof(fileName));
                if (string.IsNullOrWhiteSpace(folderPath))
                    throw new ArgumentException("文件夹路径不能为空或空白", nameof(folderPath));

                // 剔除文件名中的非法字符
                char[] invalidChars = Path.GetInvalidFileNameChars();
                string safeFileName = string.Concat(fileName.Split(invalidChars));
                if (string.IsNullOrWhiteSpace(safeFileName))
                    throw new ArgumentException("处理后文件名为空，请提供有效文件名", nameof(fileName));

                // 拼接完整文件路径
                string fullPath = Path.Combine(folderPath, safeFileName + ".csv");

                // 若文件夹不存在则创建
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                // 写入文件（UTF-8 带 BOM，便于 Excel 正确识别中文）
                using (StreamWriter sw = new StreamWriter(fullPath, false, Encoding.UTF8))
                {
                    // 写入列标题
                    if (dt.Columns.Count > 0)
                    {
                        var headers = dt.Columns.Cast<DataColumn>()
                            .Select(col => EscapeCsvField(col.ColumnName));
                        sw.WriteLine(string.Join(",", headers));
                    }

                    // 写入数据行
                    foreach (DataRow row in dt.Rows)
                    {
                        var fields = row.ItemArray
                            .Select(item => EscapeCsvField(item?.ToString() ?? string.Empty));
                        sw.WriteLine(string.Join(",", fields));
                    }
                }

                // 成功：返回完整路径
                return fullPath;
            }
            catch (Exception ex)
            {
                // VSTO 环境中可根据需要替换为日志记录或用户提示
                System.Diagnostics.Debug.WriteLine($"CSV 导出失败：{ex.Message}");
                return null;   // 失败返回 null
            }
        }

        /// <summary>
        /// 对 CSV 字段进行转义（处理逗号、引号、换行）
        /// </summary>
        private static string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return string.Empty;

            // 需要引号包裹的条件：包含逗号、双引号、换行符（\n、\r）
            bool mustQuote = field.Contains(",")
                             || field.Contains("\"")
                             || field.Contains("\n")
                             || field.Contains("\r");

            if (mustQuote)
            {
                string escaped = field.Replace("\"", "\"\"");
                return $"\"{escaped}\"";
            }

            return field;
        }
    }

}
