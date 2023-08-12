using NUnit.Framework;
using Jt.Common.Tool.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace Jt.Common.Tool.Excel.Tests
{
    [TestFixture()]
    public class ExcelHelperTests
    {
        ExcelHelper _helper = new ExcelHelper();

        [Test()]
        public void OpenEmptyExcelTest()
        {
            _helper.OpenEmptyExcel();
            _helper.SaveAs(@"E:\test\1.xlsx");
            Assert.Pass("ok");
        }

        [Test()]
        public void OpenEmptyExcelTest1()
        {
            _helper.OpenEmptyExcel("test");
            _helper.SaveAs(@"E:\test\2.xlsx");
            Assert.Pass("ok");
        }

        [Test()]
        public void OpenFileExcelTest()
        {
            _helper.OpenFileExcel(@"E:\test\1.xlsx");
            Assert.Pass("ok");
        }

        [Test()]
        public void SelectOrCreateSheetTest()
        {
            _helper.OpenEmptyExcel();
            _helper.SelectOrCreateSheet("test2");
            Assert.Pass("ok");
        }

        [Test()]
        public void OpenFileExcelTest1()
        {
            _helper.OpenFileExcel(@"E:\test\1.xlsx", "test2");
            Assert.Pass("ok");
        }

        [Test()]
        public void DeleteSheetTest()
        {
            _helper.OpenEmptyExcel();
            _helper.DeleteSheet("test2");
            Assert.Pass("ok");
        }

        [Test()]
        public void ExportTest()
        {
            _helper.OpenEmptyExcel();
            var data = Enumerable.Range(1, 10).Select(x => new User { Name = x.ToString(), Password = x.ToString() });
            _helper.Export(data);
            _helper.SaveAs(@"E:\test\3.xlsx");
            Assert.Pass("ok");
        }

        [Test()]
        public void InsertTest()
        {
            _helper.OpenEmptyExcel();
            var data = Enumerable.Range(1, 10).Select(x => new User { Name = x.ToString(), Password = x.ToString() });
            _helper.Insert(data, 3);
            _helper.SaveAs(@"E:\test\3.xlsx");
            Assert.Pass("ok");
        }

        [Test()]
        public void ReadTest()
        {
            _helper.OpenEmptyExcel();
            Assert.Pass("ok");
        }

        [Test()]
        public void SetStyleTest()
        {
            _helper.OpenEmptyExcel();
            var data = Enumerable.Range(1, 10).Select(x => new User { Name = x.ToString(), Password = x.ToString() });
            _helper.Export(data);
            _helper.SetStyle(1, 1, _helper.RowCount, _helper.ColCount);
            _helper.SaveAs(@"E:\test\3.xlsx");
            Assert.Pass("ok");
        }

        [Test()]
        public void SaveTest()
        {
            Assert.Pass("ok");
        }

        [Test()]
        public void SaveAsTest()
        {
            Assert.Pass("ok");
        }

        [Test()]
        public void DisposeTest()
        {
            Assert.Pass("ok");
        }
    }
}