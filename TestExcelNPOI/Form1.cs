using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using NPOI.SS.UserModel;

namespace TestExcelNPOI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FileStream inputStream = new FileStream(@"input.xls", FileMode.Open);
            HSSFWorkbook inputBook = new HSSFWorkbook(inputStream);
            HSSFSheet inputSheet = (HSSFSheet)inputBook.GetSheetAt(0);

            //input內的標題列
            HSSFRow titleRow = (HSSFRow)inputSheet.GetRow(0);
            int titleCellCount = titleRow.LastCellNum;

            //取得資料(由1開始)
            int inputRowCount = inputSheet.LastRowNum;
            for (int i=1; i< inputRowCount; ++i)
            {
                HSSFWorkbook outputBook = new HSSFWorkbook();

                // 在 Excel 工作簿中建立工作表，名稱為 Sheet1
                HSSFSheet outputSheet = (HSSFSheet)outputBook.CreateSheet("sheet1");

                //取得第N列資料
                HSSFRow inputRow = (HSSFRow)inputSheet.GetRow(i);
                string name = "";

                //先寫入標題
                for (int j = 0; j < titleCellCount; ++j)
                {
                    if (j == 0)
                    {
                        outputSheet.CreateRow(0);
                        outputSheet.GetRow(0).Height = titleRow.Height;
                    }

                    //input.J copy to output.J
                    ICell curCell = titleRow.GetCell(j);
                    if (curCell != null)
                    {
                        Console.WriteLine("title cell = " + curCell.StringCellValue);
                        outputSheet.GetRow(0).CreateCell(j).SetCellValue(titleRow.GetCell(j).StringCellValue);

                        HSSFCellStyle newTitleCellStyle = (HSSFCellStyle)outputBook.CreateCellStyle();
                        newTitleCellStyle.CloneStyleFrom(titleRow.GetCell(j).CellStyle);
                        outputSheet.GetRow(0).GetCell(j).CellStyle = newTitleCellStyle;
                    }
                }

                //寫入資料
                for (int k=0; k< titleCellCount; ++k)
                {
                    if (k == 2)
                    {
                        //取得名字
                        name = inputRow.GetCell(k).StringCellValue;            
                    }

                    if (k == 0)
                    {
                        outputSheet.CreateRow(1);
                        outputSheet.GetRow(1).Height = inputRow.Height;
                    }

                    //數字
                    if (inputRow.GetCell(k).CellType == NPOI.SS.UserModel.CellType.Numeric || inputRow.GetCell(k).CellType == NPOI.SS.UserModel.CellType.Formula)
                    {
                        var value = inputRow.GetCell(k).NumericCellValue;
                        outputSheet.GetRow(1).CreateCell(k).SetCellValue(value);
                    }
                    else
                    {
                        outputSheet.GetRow(1).CreateCell(k).SetCellValue(inputRow.GetCell(k).StringCellValue);
                    }

                    HSSFCellStyle newCellStyle = (HSSFCellStyle)outputBook.CreateCellStyle();
                    newCellStyle.CloneStyleFrom(inputRow.GetCell(k).CellStyle);
                    outputSheet.GetRow(1).GetCell(k).CellStyle = newCellStyle;
                    outputSheet.GetRow(1).GetCell(k).CellStyle.WrapText = true;              
                    outputSheet.AutoSizeColumn(k);
                    outputSheet.ProtectSheet("123");
                    
                    /*
                    //標題列
                    if (k == 0)
                    {
                        
                    }
                    else
                    {
                        double value = inputRow.GetCell(j).NumericCellValue;
                        outputSheet.GetRow(0).CreateCell(j).SetCellType(NPOI.SS.UserModel.CellType.Numeric);
                        outputSheet.GetRow(0).CreateCell(j).SetCellValue(value);
                    }
                    */
                    
                }



                string fileName = @"綺綺\" + name + ".xls";
                Directory.CreateDirectory(Path.GetDirectoryName(fileName));
                FileStream outputStream = new FileStream(fileName, FileMode.Create);
                outputBook.Write(outputStream);
                outputBook.WriteProtectWorkbook("123", "456");
                outputStream.Close();

            }

            MessageBox.Show("END");

            /*
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook w = excelApp.Workbooks.Add("AAA");
            w.Password = "123";
            w.SaveAs();
            w.Close();
            */
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
