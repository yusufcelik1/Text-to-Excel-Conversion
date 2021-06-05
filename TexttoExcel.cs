using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using System.IO;
namespace DosyaKopyalama
{
    class Class1
    {
        public void TexttoExcel()
        {
            #region values
            string[] InputNamesLine = File.ReadAllLines(@"");//Your text file location
            excel.Application oXL;
            excel._Workbook oWB;
            excel._Worksheet oSheet;
            excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;
            #endregion
            try
            {
                //start excel and get application object
                oXL = new excel.Application() { Visible = true };
                //Create new workbook
                oWB = (excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (excel._Worksheet)oWB.ActiveSheet;
                //Add table headers going cell
                oSheet.Cells[1, 1] = "Parameter_Names";
                oSheet.Cells[1, 2] = "Values";
                //Format A1:B1 as bold and vertical alignment=center
                oSheet.get_Range("A1", "B1").Font.Bold = true;
                oSheet.get_Range("A1", "B1").VerticalAlignment = excel.XlVAlign.xlVAlignCenter;
                //Write value in cells,
                for (int i = 0; i < 1; i ++)
                {
                    foreach (string line in InputNamesLine)
                    {
                        if (line != "")
                        {
                            string[] columns = line.Split(' ','\r');
                            string column1 = columns[0];
                            string column2 = columns[1];
                            //oSheet.Cells[1][i + 1] = i;
                            oSheet.Cells[1][i + 2] = column1;
                            oSheet.Cells[2][i + 2] = column2;
                            i = i+1;
                        }
                    }
                }
                Thread.Sleep(5000);
                oRng = oSheet.get_Range("A1", "B1");
				//Your file already exists
                oXL.DisplayAlerts = false;
                oWB.SaveAs(@"LocationWriteHere.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,Type.Missing, Type.Missing);//excel save location
                oWB.Close();
                oXL.Quit();
            }
            catch (Exception e)
            {
				//if you have a error, show error with message box
                System.Windows.Forms.MessageBox.Show("exception" + e);
            }
        }
    }
}
