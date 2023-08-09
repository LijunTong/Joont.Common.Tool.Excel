namespace Jt.Common.Tool.Excel.Tests
{
    [TestFixture()]
    public class ExcelHelperTests
    {
        ExcelHelper _excelHelper;

        public ExcelHelperTests()
        {
            _excelHelper = new ExcelHelper();
        }

        [Test()]
        public void GetOrAddSheetTest()
        {
            _excelHelper.GetOrAddSheet("test");
            _excelHelper.SaveAs(@"E:\test\a.xlsx");
            Assert.Pass("ok");
        }

        [Test()]
        public void DeleteSheetTest()
        {
            _excelHelper.DeleteSheet("test");
            _excelHelper.SaveAs(@"E:\test\a.xlsx");
            Assert.Pass("ok");
        }

        [Test()]
        public void FillDataTest()
        {
            var data = Enumerable.Range(1, 10).Select(x => new User { Name = x.ToString(), Password = x.ToString() });
            _excelHelper.Export("test", data);
            _excelHelper.SaveAs(@"E:\test\a.xlsx");
            Assert.Pass("ok");
        }

        [Test()]
        public void InsertValuesTest()
        {
            var data = Enumerable.Range(1, 10).Select(x => new User { Name = x.ToString(), Password = x.ToString() });
            _excelHelper.OpenExcel(@"E:\test\a.xlsx");
            _excelHelper.Insert("test", data, 5);
            _excelHelper.Save();
            Assert.Pass("ok");
        }

        [Test()]
        public void ReadTest()
        {
            _excelHelper.OpenExcel(@"E:\test\a.xlsx");
            var data = _excelHelper.Read<User>(6);
            Assert.Pass("ok");
        }
    }
}