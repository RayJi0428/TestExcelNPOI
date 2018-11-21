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
        private int keepRow = 2;//保留列數

        private int nameCell = 2;//姓名所在欄

        private Boolean DEBUG = false;
        private string INPUT_URL = @"input.xls";
        private string OUTPUT_URL = @"綺綺\#1_#2.xls";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (DEBUG)
            {
                this.INPUT_URL = "../../../../" + this.INPUT_URL;
                this.OUTPUT_URL = "../../../../" + this.OUTPUT_URL;
            }
            try
            {
                HSSFWorkbook inputBook = new HSSFWorkbook(new FileStream(INPUT_URL, FileMode.Open));//input不可外流，放在GitHub外層
                HSSFSheet inputSheet = (HSSFSheet)inputBook.GetSheetAt(0);//目前只取第一分頁sheet0

                //取得資料
                int iNumRow = inputSheet.LastRowNum;
                for (int i = 0; i < iNumRow; ++i)
                {
                    //保留標題列略過
                    if (i < this.keepRow)
                        continue;

                    HSSFWorkbook outputBook = new HSSFWorkbook();
                    HSSFSheet outputSheet = (HSSFSheet)outputBook.CreateSheet("sheet1");// 在 Excel 工作簿中建立工作表，名稱為 Sheet1

                    HSSFRow iCurRow = (HSSFRow)inputSheet.GetRow(i);//取得第N列資料
                    

                    //標題列數(注意:最後一條保留列為資料標題)
                    int titleCellCount = 0;

                    int ci = 0;
                    //先寫入保留標題列---------------------------------------------------------------------------------------
                    for (int ki = 0; ki < this.keepRow; ++ki)
                    {
                        //查詢保留列
                        HSSFRow iTitleRow = (HSSFRow)inputSheet.GetRow(ki);
                        //保留列有多少欄?
                        titleCellCount = iTitleRow.LastCellNum;
                        //建立保留列
                        HSSFRow oTitleRow = (HSSFRow)outputSheet.CreateRow(ki);
                        //設定同等列高
                        oTitleRow.Height = iTitleRow.Height;

                        for (ci = 0; ci < titleCellCount; ++ci)
                        {
                            //input.J copy to output.J
                            ICell inTItleCell = iTitleRow.GetCell(ci);

                            if (inTItleCell != null)
                            {
                                ICell oTitleCell = oTitleRow.CreateCell(ci);
                                oTitleCell.SetCellValue(inTItleCell.StringCellValue);

                                //複製欄格式
                                HSSFCellStyle newTitleCellStyle = (HSSFCellStyle)outputBook.CreateCellStyle();
                                newTitleCellStyle.CloneStyleFrom(inTItleCell.CellStyle);
                                oTitleCell.CellStyle = newTitleCellStyle;
                            }
                        }
                    }

                    //寫入資料---------------------------------------------------------------------------------------
                    //建立員工資料列
                    string name = "temp";
                    string apartment = "temp";
                    HSSFRow oDataRow = (HSSFRow)outputSheet.CreateRow(this.keepRow);
                    oDataRow.Height = iCurRow.Height;
                    for (ci = 0; ci < titleCellCount; ++ci)
                    {
                        ICell inDataCell = iCurRow.GetCell(ci);

                        if (inDataCell == null)
                            continue;

                        //取得名字
                        if (ci == this.nameCell)
                        {
                            name = inDataCell.StringCellValue;
                        }
                        if (ci == this.nameCell + 1)
                        {
                            apartment = inDataCell.StringCellValue;
                        }

                        ICell oDataCell = oDataRow.CreateCell(ci);
                 
                        //到職日特別處理
                        if (ci == 0)
                        {
                            var value = inDataCell.DateCellValue;
                            oDataCell.SetCellValue(value.ToShortDateString());//轉換為2014/12/19
                        }
                        //數字
                        else if (inDataCell.CellType == NPOI.SS.UserModel.CellType.Numeric || inDataCell.CellType == NPOI.SS.UserModel.CellType.Formula)
                        {
                            var value = inDataCell.NumericCellValue;
                            oDataCell.SetCellValue(value);
                        }
                        else
                        {
                            oDataCell.SetCellValue(inDataCell.StringCellValue);
                        }

                        HSSFCellStyle newCellStyle = (HSSFCellStyle)outputBook.CreateCellStyle();
                        newCellStyle.CloneStyleFrom(inDataCell.CellStyle);
                        oDataCell.CellStyle = newCellStyle;
                        outputSheet.AutoSizeColumn(ci);
                        //oDataCell.CellStyle.WrapText = true;
                    }

                    //output不可外流，放在GitHub外層
                    string fileName = OUTPUT_URL;
                    fileName = fileName.Replace("#1", apartment);
                    fileName = fileName.Replace("#2", name);
                    Directory.CreateDirectory(Path.GetDirectoryName(fileName));
                    FileStream outputStream = new FileStream(fileName, FileMode.Create);
                    outputBook.Write(outputStream);
                    outputBook.WriteProtectWorkbook("123", "456");
                    outputStream.Close();

                }

                MessageBox.Show("薪資拆分成功!");
            }
            catch (Exception error)
            {
                MessageBox.Show("異常錯誤 " + error.Message);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            this.keepRow = Int32.Parse(textBox1.Text);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            this.nameCell = Int32.Parse(textBox2.Text);
        }

        private  void backup()
        {
            /*
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook w = excelApp.Workbooks.Add("AAA");
            w.Password = "123";
            w.SaveAs();
            w.Close();
            */
        }
    }
}
