using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

using System.Data.OleDb;

namespace ReadExcelData
{
    class Program
    {
        static void Main(string[] args)
        {
            Program p = new Program();
            var a = p.userInput();
            if (a == "1")
            {
                //p.readExcelIronXL();
            }
            else
            {
                p.readExcelMicrosoft();
            }


        }

        public string userInput()
        {
            Console.WriteLine("1.ironXL:");
            Console.WriteLine("2.microsoft Office Interop:");

            string menu = Console.ReadLine();
            return menu;
        }
        //public void readExcelIronXL()
        //{
        //    try
        //    {
        //        WorkBook wb = WorkBook.Load(@"C:\Users\yjchia\Desktop\excelGet.xlsx");
        //        WorkSheet st = wb.WorkSheets.First();

        //        foreach (var cll in st["A2:A4"])
        //        {
        //            Console.WriteLine("Cell {0} has value '{1}'", cll.AddressString, cll.Text);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(ex);
        //    }

        //}

        public void readExcelMicrosoft()
        {
            try
            {
                //Create COM Objects. Create a COM object for everything that is referenced
                string filePath = "C:\\Users\\yjchia\\Desktop\\New folder (2)\\Copy of excelGet.xlsx";
                Application excel = new Application();
                Workbook wb = excel.Workbooks.Open(filePath);
                Worksheet ws = wb.Worksheets[1];

                Microsoft.Office.Interop.Excel.Range xlRange = ws.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                Console.WriteLine("row have:" + rowCount);
                Console.WriteLine("column have: " + colCount);
                user dsData = new user();

                //for(int row = 1, col = 1; row <= 5; row++)
                //{


                for (int i = 2; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        //new line
                        if (j == 1)
                        {
                            Console.Write("\r\n");
                        }


                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && Convert.ToString((ws.Cells[i, j] as Microsoft.Office.Interop.Excel.Range).Value) != null)
                        {
                            if (j == 1)
                            {
                                dsData.userAccList[i].ID = Convert.ToString((ws.Cells[i, j] as Microsoft.Office.Interop.Excel.Range).Value);
                            }
                            else
                            {
                                dsData.userAccList[i].Username = Convert.ToString((ws.Cells[i, j] as Microsoft.Office.Interop.Excel.Range).Value);
                            }
                            //Console.Write(Convert.ToString((ws.Cells[i, j] as Microsoft.Office.Interop.Excel.Range).Value) + "\t");
                        }


                    }
                }

                //var cellValue = (string)(ws.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range).Value;
                //Console.WriteLine(cellValue);

                //}
               
              
                calldataSet(dsData);
                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!
                //for (int i = 1; i <= rowCount; i++)
                //{
                //    for (int j = 1; j <= colCount; j++)
                //    {
                //        var sno = (Excel.Range)xlRange.Cells[i, j];
                //        //new line
                //        if (j == 1)
                //            Console.Write("\r\n");

                //        //write the value to the console
                //        if (xlRange.Cells[i, j] != null && sno.Value2 != null)
                //            Console.Write(sno.Value2.ToString() + "\t");
                //    }
                //}

                //cleanup

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        //public void readDbMethod()
        //{
        //    string con =
        //                @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\yjchia\Desktop\New folder (2)\excelGet.xls;" +
        //                @"Extended Properties='Excel 8.0;HDR=Yes;'";
        //    using (OleDbConnection connection = new OleDbConnection(con))
        //    {
        //        connection.Open();
        //        OleDbCommand command = new OleDbCommand("select * from [Sheet1$]", connection);
        //        using (OleDbDataReader dr = command.ExecuteReader())
        //        {
        //            while (dr.Read())
        //            {
        //                var row1Col0 = dr[0];
        //                Console.WriteLine(row1Col0);
        //            }
        //        }
        //    }
        //}

        public void calldataSet(user ds)
        {
            int i = 0;
            Console.WriteLine("this is new");
            foreach(var a in ds.userAccList)
            {
                Console.WriteLine(ds.userAccList[i].ID.ToString() + "," + ds.userAccList[i].Username.ToString());
            }
            Console.ReadKey();
        }


    }
}
