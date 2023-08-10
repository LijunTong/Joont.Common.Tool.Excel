using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace Jt.Common.Tool.Excel
{
    public class ExcelHelper : IDisposable
    {
        private ExcelPackage _excelPackage;

        private ExcelWorksheet _excelWorksheet;

        public int RowCount
        {
            get
            {
                return _excelWorksheet.Dimension.End.Row;
            }
        }

        public int ColCount
        {
            get
            {
                return _excelWorksheet.Dimension.End.Column;
            }
        }

        public ExcelHelper()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// 打开一个空白Excel文件，创建一个sheet并选中
        /// </summary>
        public void OpenEmptyExcel()
        {
            OpenEmptyExcel("sheet");
        }

        /// <summary>
        /// 打开一个空白Excel文件，创建一个sheet并选中
        /// <param name="sheetName">sheet名称</param>
        /// </summary>
        public void OpenEmptyExcel(string sheetName)
        {
            if (_excelPackage != null)
            {
                _excelPackage.Dispose();
            }

            _excelPackage = new ExcelPackage();
            SelectOrCreateSheet(sheetName);
        }

        /// <summary>
        /// 打开Excel文件，默认选中第一个sheet
        /// </summary>
        /// <param name="filePath">文件路径</param>
        public void OpenFileExcel(string filePath)
        {
            if (_excelPackage != null)
            {
                _excelPackage.Dispose();
            }

            _excelPackage = new ExcelPackage(new FileInfo(filePath));
            _excelWorksheet = _excelPackage.Workbook.Worksheets.FirstOrDefault();
        }

        /// <summary>
        /// 打开Excel文件，默认选中第一个sheet
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="sheetName">sheet名称</param>
        public void OpenFileExcel(string filePath, string sheetName)
        {
            if (_excelPackage != null)
            {
                _excelPackage.Dispose();
            }

            _excelPackage = new ExcelPackage(new FileInfo(filePath));
            SelectOrCreateSheet(sheetName);
        }

        /// <summary>
        /// 选择sheet，如果不存在则创建
        /// </summary>
        /// <param name="sheetName">sheet名称</param>
        public void SelectOrCreateSheet(string sheetName)
        {
            CheckOpenFile();
            _excelWorksheet = _excelPackage.Workbook.Worksheets.FirstOrDefault(i => i.Name == sheetName);
            if (_excelWorksheet == null)
            {
                _excelWorksheet = _excelPackage.Workbook.Worksheets.Add(sheetName);
            }
        }

        /// <summary>
        /// 删除指定的sheet
        /// </summary>
        /// <param name="ExcelPackage"></param>
        /// <param name="sheetName"></param>
        public void DeleteSheet(string sheetName)
        {
            CheckOpenFile();
            ExcelWorksheet sheet = _excelPackage.Workbook.Worksheets.FirstOrDefault(i => i.Name == sheetName);
            if (sheet != null)
            {
                _excelPackage.Workbook.Worksheets.Delete(sheet);
            }

            if (_excelWorksheet.Name == sheetName)
            {
                _excelWorksheet = _excelPackage.Workbook.Worksheets.FirstOrDefault();
            }
        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">数据</param>
        public Stream Export<T>(IEnumerable<T> list)
        {
            _excelWorksheet.Cells["A1"].LoadFromCollection(list, true);
            return _excelPackage.Stream;
        }

        /// <summary>
        /// 插入行
        /// </summary>
        /// <param name="values">数据</param>
        /// <param name="rowIndex">插入位置，起始位置为1</param>
        public Stream Insert<T>(IEnumerable<T> values, int rowIndex)
        {
            if (values == null)
            {
                throw new ArgumentNullException(nameof(values));
            }

            if (values.Count() == 0)
            {
                return null;
            }

            _excelWorksheet.InsertRow(rowIndex, values.Count());
            _excelWorksheet.Cells[rowIndex, 1].LoadFromCollection(values, true);
            return _excelPackage.Stream;
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
            if (startIndex > RowCount)
            {
                throw new Exception("起始位置超出范围");
            }

            Type type = typeof(T);
            var properties = type.GetProperties();
            List<T> result = new List<T>();
            for (int i = startIndex; i <= RowCount; i++)
            {
                var row = (T)Activator.CreateInstance(type);
                foreach (var item in properties)
                {
                    var attr = item.GetCustomAttributes(typeof(EpplusTableColumnAttribute), false).FirstOrDefault() as EpplusTableColumnAttribute;
                    if (attr != null)
                    {
                        var order = attr.Order; // 以order作为列的索引
                        if (order <= ColCount)
                        {
                            string value = _excelWorksheet.Cells[i, order].Value.ToString();
                            item.SetValue(row, Convert.ChangeType(value, item.PropertyType));
                        }
                    }
                }

                result.Add(row);
            }

            return result;
        }

        /// <summary>
        /// 设置单元格格式
        /// </summary>
        /// <param name="startRow">起始行</param>
        /// <param name="startCol">起始列</param>
        /// <param name="endRow">结束行</param>
        /// <param name="endCol">结束列</param>
        /// <param name="fontName">字体名称</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="horizontalAlignment">水平对齐方式</param>
        /// <param name="verticalAlignment">垂直对齐方式</param>
        /// <param name="wrapText">自动换行</param>
        public void SetStyle(int startRow,
            int startCol,
            int endRow,
            int endCol, 
            string fontName = "宋体", 
            float fontSize = 13f,
            ExcelHorizontalAlignment horizontalAlignment = ExcelHorizontalAlignment.Center, 
            ExcelVerticalAlignment verticalAlignment = ExcelVerticalAlignment.Center,
            bool wrapText = false
            )
        {
            startRow = startRow < 1 ? 1 : startRow;
            startCol = startCol < 1 ? 1 : startCol;
            endRow = endRow > RowCount ? RowCount : endRow;
            endCol = endCol < ColCount ? ColCount : endCol;
            var style = _excelWorksheet.Cells[startRow, startCol, endRow, endCol].Style;
            style.Font.Name = fontName;
            style.Font.Size = fontSize;
            style.HorizontalAlignment = horizontalAlignment;
            style.VerticalAlignment = verticalAlignment;
            style.WrapText = wrapText;
        }

        /// <summary>
        /// 保存修改
        /// </summary>
        public void Save()
        {
            CheckOpenFile();
            _excelPackage.Save();
        }

        /// <summary>
        /// 保存修改
        /// </summary>
        public void SaveAs(string filePath)
        {
            CheckOpenFile();
            _excelPackage.SaveAs(filePath);
        }

        /// <summary>
        /// 检查Excel是否打开
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private bool CheckOpenFile()
        {
            if (_excelPackage == null)
            {
                throw new Exception("Excel未打开");
            }

            if (_excelWorksheet == null)
            {
                throw new Exception("sheet未选中");
            }

            return true;
        }

        public void Dispose()
        {
            _excelWorksheet.Dispose();
            _excelPackage.Dispose();
        }
    }
}
