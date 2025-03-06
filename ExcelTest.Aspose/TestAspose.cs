using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using NUnit.Framework;
using Aspose;
using Aspose.Cells;

namespace ExcelTest.Aspose
{
    [TestFixture]
    public class TestAspose
    {
        [Test]
        public void TestDeleteRow()
        {
            var wb = GetTestWorkbook();

            var sheet1 = wb.Worksheets[0];
            sheet1.Cells.DeleteRow(0);

            wb.Save(@"c:\temp\testSheet_deleteOneRow.xlsx");
        }

        [Test]
        public void TestWorkbookCreation()
        {
            var wb = GetTestWorkbook();
            wb.Save(@"c:\temp\testSheet.xlsx");
        }

        public Workbook GetTestWorkbook()
        {
            // For complete examples and data files, please go to https://github.com/aspose-cells/Aspose.Cells-for-.NET
            // Create workbook
            Workbook wb = new Workbook();

            // add some data to the first sheet
            Worksheet sht1 = wb.Worksheets[0];
            for (int i = 1; i<=100; i++)
            {
                sht1.Cells[$"A{i}"].PutValue($"Value {i}");

                sht1.Cells[$"C{i}"].PutValue(i);

                sht1.Cells[$"E{i+10}"].PutValue(i+10);
            }

            // create formula
            Cell b15 = sht1.Cells["B15"];
            b15.Formula = "=C15*E15";
            CellsFactory f = new CellsFactory();
            Style s = f.CreateStyle();     
            s.BackgroundColor = Color.Red;
            s.Font.IsBold = true;
            s.ForegroundColor = Color.Yellow;
            b15.SetStyle(s);

            // create some named range:
            Range range = sht1.Cells.CreateRange("B20", "E20");
            range.Name = "Test_Range"; //Setting the name of the named range           

            // set calculatioin mode
            wb.Settings.CalcMode = CalcModeType.Automatic;

            return wb;
        }
    }
}
