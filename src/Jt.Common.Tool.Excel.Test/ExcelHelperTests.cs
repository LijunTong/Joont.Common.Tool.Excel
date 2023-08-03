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
            Assert.Pass();
        }

        [Test()]
        public void DeleteSheetTest()
        {
            _excelHelper.DeleteSheet("test");
            _excelHelper.SaveAs(@"E:\test\a.xlsx");
            Assert.Pass();
        }

        [Test()]
        public void FillDataTest()
        {
            var data = Enumerable.Range(1, 10).Select(x => new User { Name = x.ToString(), Password = x.ToString() });
            _excelHelper.FillData(data, "test");
            _excelHelper.SaveAs(@"E:\test\a.xlsx");
            Assert.Pass();
        }

        [Test()]
        public void InsertValuesTest()
        {
            Assert.Pass();
        }
    }
}