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
    public class ExcelHelper : IDisposable
    {
        public ExcelPackage ExcelPackage { get; private set; }
        private Stream fs;

        public ExcelHelper()
        {
            ExcelPackage = new ExcelPackage();
        }

        private bool CheckOpenFile()
        {
            if (ExcelPackage == null)
            {
                throw new Exception("Excel未打开");
            }

            return true;
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
        /// 填充数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">数据源</param>
        /// <param name="sheetName">sheet名称</param>
        /// <param name="isDeleteSameNameSheet">是否删除已存在的同名sheet，false时将重命名导出的sheet</param>
        public Stream FillData<T>(IEnumerable<T> list, string sheetName, bool isDeleteSameNameSheet = true)
        {
            ExcelWorksheet sheet = AddSheet(sheetName, isDeleteSameNameSheet);
            sheet.Cells["A1"].LoadFromCollection(list, true);
            return ExcelPackage.Stream;
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
