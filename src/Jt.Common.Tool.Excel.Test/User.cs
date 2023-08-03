using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jt.Common.Tool.Excel.Tests
{
    public class User
    {
        [EpplusTableColumn(Header = "名称")]
        public string Name { get; set; }

        [EpplusTableColumn(Header = "密码")]
        public string Password { get; set; }

        [EpplusIgnore]
        public int Age { get; set; }
    }
}
