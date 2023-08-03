using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing.Drawing2D;
using System.Drawing;
using System.IO;
using System.Linq;

namespace Jt.Common.Tool.Excel
{
    public class ExcelHelper
    {
        public ExcelPackage ExcelPackage { get; private set; }
        private Stream fs;

        public ExcelHelper()
        {
        }

        public void OpenOrCreate(string filePath)
        {
            if (File.Exists(filePath))
            {
                var file = new FileInfo(filePath);
                ExcelPackage = new ExcelPackage(file);
            }
            else
            {
                fs = File.Create(filePath);
                ExcelPackage = new ExcelPackage(fs);
            }
        }

        /// <summary>
        /// 获取sheet，没有则创建
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public ExcelWorksheet GetOrAddSheet(string sheetName)
        {
            ExcelWorksheet ws = ExcelPackage.Workbook.Worksheets.FirstOrDefault(i => i.Name == sheetName);
            if (ws == null)
            {
                ws = ExcelPackage.Workbook.Worksheets.Add(sheetName);
            }
            return ws;
        }

        /// <summary>
        /// 删除指定的sheet
        /// </summary>
        /// <param name="ExcelPackage"></param>
        /// <param name="sheetName"></param>
        public void DeleteSheet(string sheetName)
        {
            var sheet = ExcelPackage.Workbook.Worksheets.FirstOrDefault(i => i.Name == sheetName);
            if (sheet != null)
            {
                ExcelPackage.Workbook.Worksheets.Delete(sheet);
            }
        }

        /// <summary>
        /// 导出列表到excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">数据源</param>
        /// <param name="sheetName">sheet名称</param>
        /// <param name="isDeleteSameNameSheet">是否删除已存在的同名sheet，false时将重命名导出的sheet</param>
        public void AppendSheetToWorkBook<T>(IEnumerable<T> list, string sheetName, bool isDeleteSameNameSheet = true)
        {
            ExcelWorksheet ws = AddSheet(sheetName, isDeleteSameNameSheet);
            ws.Cells["A1"].LoadFromCollection(list, true);
        }

        /// <summary>
        /// 插入行
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="values">行类容，一个单元格一个对象</param>
        /// <param name="rowIndex">插入位置，起始位置为1</param>
        public void InsertValues(string sheetName, List<object> values, int rowIndex)
        {
            var sheet = GetOrAddSheet(sheetName);
            sheet.InsertRow(rowIndex, 1);
            int i = 1;
            foreach (var item in values)
            {
                sheet.SetValue(rowIndex, i, item);
                i++;
            }
        }

        /// <summary>
        /// 保存修改
        /// </summary>
        public void Save()
        {
            ExcelPackage.Save();
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
                while (ExcelPackage.Workbook.Worksheets.Any(i => i.Name == sheetName))
                {
                    sheetName = sheetName + "(1)";
                }
            }

            ExcelWorksheet ws = ExcelPackage.Workbook.Worksheets.Add(sheetName);
            return ws;
        }

        public void Dispose()
        {
            ExcelPackage.Dispose();
            if (fs != null)
            {
                fs.Dispose();
                fs.Close();
            }
        }
    }
}
