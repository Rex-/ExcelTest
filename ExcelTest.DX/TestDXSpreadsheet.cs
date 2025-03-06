using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using NUnit.Framework;
using DevExpress.Spreadsheet;


namespace ExcelTest.DX
{
    [TestFixture]
    public class TestDXSpreadsheet
    {
        [Test]
        public void TestDeleteRow()
        {
            var wb = GetTestWorkbook();

            var sheet1 = wb.Worksheets[0];
            sheet1.Rows[0].Delete();

            wb.SaveDocument(@"c:\temp\testSheet_DX_deleteOneRow.xlsx");
        }

        [Test]
        public void TestWorkbookCreation()
        {
            var wb = GetTestWorkbook();
            wb.SaveDocument(@"c:\temp\testSheet_DX.xlsx");
        }
        public Workbook GetTestWorkbook()
        {
            // For complete examples and data files, please go to https://github.com/aspose-cells/Aspose.Cells-for-.NET
            // Create workbook
            Workbook wb = new Workbook();
            wb.Unit = DevExpress.Office.DocumentUnit.Point;

            // add some data to the first sheet
            Worksheet sht1 = wb.Worksheets[0];

            wb.BeginUpdate();
            try
            {
                for (int i = 1; i <= 100; i++)
                {
                    sht1.Cells[$"A{i}"].SetValue($"Value {i}");

                    sht1.Cells[$"C{i}"].SetValue(i);

                    sht1.Cells[$"E{i + 10}"].SetValue(i + 10);
                }

                // create formula
                Cell b15 = sht1.Cells["B15"];
                b15.Formula = "=C15*E15";
                b15.Style.BeginUpdate();
                b15.FillColor = Color.Red;
                b15.Font.Color = Color.White;

                // create some named range:
                CellRange rangeB20E20 = sht1.Range["B20:E20"];
                rangeB20E20.Name = "Test_Range"; //Setting the name of the named range     

                // set calculation mode!!
                wb.DocumentSettings.Calculation.Mode = CalculationMode.Automatic;
            }
            finally
            {
                wb.EndUpdate();
            }

            wb.CalculateFull();
                  
            return wb;
        }

    }
}
