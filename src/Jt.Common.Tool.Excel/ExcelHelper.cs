using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace Jt.Common.Tool.Excel
{
    public class ExcelHelper : IDisposable
    {
        public ExcelPackage ExcelPackage { get; private set; }

        public ExcelHelper()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage = new ExcelPackage();
        }

        public ExcelHelper(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage = new ExcelPackage(new FileInfo(filePath));
        }

        /// <summary>
        /// 打开Excel文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        public void OpenExcel(string filePath)
        {
            if (ExcelPackage != null)
            {
                ExcelPackage.Dispose();
            }

            ExcelPackage = new ExcelPackage(new FileInfo(filePath));
        }

        /// <summary>
        /// 获取sheet，没有则创建
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public ExcelWorksheet GetOrAddSheet(string sheetName)
        {
            CheckOpenFile();
            ExcelWorksheet sheet = ExcelPackage.Workbook.Worksheets.FirstOrDefault(i => i.Name == sheetName);
            if (sheet == null)
            {
                sheet = ExcelPackage.Workbook.Worksheets.Add(sheetName);
            }
            return sheet;
        }

        /// <summary>
        /// 删除指定的sheet
        /// </summary>
        /// <param name="ExcelPackage"></param>
        /// <param name="sheetName"></param>
        public void DeleteSheet(string sheetName)
        {
            CheckOpenFile();
            ExcelWorksheet sheet = ExcelPackage.Workbook.Worksheets.FirstOrDefault(i => i.Name == sheetName);
            if (sheet != null)
            {
                ExcelPackage.Workbook.Worksheets.Delete(sheet);
            }
        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName">sheet名称</param>
        /// <param name="list">数据</param>
        /// <param name="isDeleteSameNameSheet">是否删除已存在的同名sheet，false时将重命名导出的sheet</param>
        public Stream Export<T>(string sheetName, IEnumerable<T> list, bool isDeleteSameNameSheet = true)
        {
            ExcelWorksheet sheet = AddSheet(sheetName, isDeleteSameNameSheet);
            sheet.Cells["A1"].LoadFromCollection(list, true);
            return ExcelPackage.Stream;
        }

        /// <summary>
        /// 插入行
        /// </summary>
        /// <param name="sheetName">sheet名称</param>
        /// <param name="values">数据</param>
        /// <param name="rowIndex">插入位置，起始位置为1</param>
        public void Insert<T>(string sheetName, IEnumerable<T> values, int rowIndex)
        {
            if (values == null)
            {
                throw new ArgumentNullException(nameof(values));
            }

            if (values.Count() == 0)
            {
                return;
            }

            var sheet = GetOrAddSheet(sheetName);
            sheet.InsertRow(rowIndex, values.Count());

            // 指定某个单元格
            sheet.Cells[rowIndex, 1].LoadFromCollection(values, true);
        }

        /// <summary>
        /// 读取excel数据到List
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="startIndex">起始位置</param>
        /// <returns></returns>
        public List<T> Read<T>(int startIndex = 1)
        {
            CheckOpenFile();
            var sheet = ExcelPackage.Workbook.Worksheets.FirstOrDefault();
            int colCount = sheet.Dimension.End.Column;
            int rowCount = sheet.Dimension.End.Row;
            if (startIndex > rowCount)
            {
                throw new Exception("起始位置超出范围");
            }

            Type type = typeof(T);
            var properties = type.GetProperties();
            List<T> result = new List<T>();
            for (int i = startIndex; i <= rowCount; i++)
            {
                var row = (T)Activator.CreateInstance(type);
                foreach (var item in properties)
                {
                    var attr = item.GetCustomAttributes(typeof(EpplusTableColumnAttribute), false).FirstOrDefault() as EpplusTableColumnAttribute;
                    if (attr != null)
                    {
                        var order = attr.Order; // 以order作为列的索引
                        if (order <= colCount)
                        {
                            string value = sheet.Cells[i, order].Value.ToString();
                            item.SetValue(row, Convert.ChangeType(value, item.PropertyType));
                        }
                    }
                }

                result.Add(row);
            }

            return result;
        }

        /// <summary>
        /// 保存修改
        /// </summary>
        public void Save()
        {
            CheckOpenFile();
            ExcelPackage.Save();
        }

        /// <summary>
        /// 保存修改
        /// </summary>
        public void SaveAs(string filePath)
        {
            CheckOpenFile();
            ExcelPackage.SaveAs(filePath);
        }

        /// <summary>
        /// 添加Sheet到ExcelPackage
        /// </summary>
        /// <param name="ExcelPackage">ExcelPackage</param>
        /// <param name="sheetName">sheet名称</param>
        /// <param name="isDeleteSameNameSheet">如果存在同名的sheet是否删除</param>
        /// <returns></returns>
        private ExcelWorksheet AddSheet(string sheetName, bool isDeleteSameNameSheet)
        {
            if (isDeleteSameNameSheet)
            {
                DeleteSheet(sheetName);
            }
            else
            {
                CheckOpenFile();
                if (ExcelPackage.Workbook.Worksheets.Any(i => i.Name == sheetName))
                {
                    sheetName += "(1)";
                }
            }

            ExcelWorksheet sheet = ExcelPackage.Workbook.Worksheets.Add(sheetName);
            return sheet;
        }

        /// <summary>
        /// 检查Excel是否打开
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private bool CheckOpenFile()
        {
            if (ExcelPackage == null)
            {
                throw new Exception("Excel未打开");
            }

            return true;
        }

        public void Dispose()
        {
            ExcelPackage.Dispose();
        }
    }
}
