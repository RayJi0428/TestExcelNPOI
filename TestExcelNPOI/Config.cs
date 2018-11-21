using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace TestExcelNPOI
{
    public class Config
    {
        public List<ConfigVO> voList = new List<ConfigVO>();

        public Config(string configURL)
        {
            try
            {
                HSSFWorkbook inputBook = new HSSFWorkbook(new FileStream(configURL, FileMode.Open));//input不可外流，放在GitHub外層
                HSSFSheet inputSheet = (HSSFSheet)inputBook.GetSheetAt(0);//目前只取第一分頁sheet0


                //取得資料
                int iNumRow = inputSheet.LastRowNum;
                for (int i = 0; i < iNumRow; ++i)
                {
                    if (i < 2)
                        continue;

                    HSSFRow curRow = (HSSFRow)inputSheet.GetRow(i);//取得第N列資料

                    ConfigVO vo = new ConfigVO(curRow);                 


                    this.voList.Add(vo);
                }

                Console.WriteLine("GG");
            }
            catch (Exception error)
            {
                MessageBox.Show("異常錯誤 " + error.Message);
            }
        }

        public Boolean needToSkipCol(string name, int col)
        {
            ConfigVO vo = this.voList.Find(x => x.name == name);
            if (vo != null)
            {
                return vo.needToSkip(col);
            }
            else
            {
                return false;
            }
            
        }

        private ConfigVO myFindFun(ConfigVO vo)
        {
            return vo;

        }
    }
}
