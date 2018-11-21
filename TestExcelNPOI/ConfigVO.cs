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
    public class ConfigVO
    {
        public string name;

        public int[] deleteColumns;

        public ConfigVO(HSSFRow row)
        {

            this.name = row.GetCell(2).ToString();

            string deleteColumnsStr = row.GetCell(4).ToString();
            if (deleteColumnsStr.Length == 0)
            {
                this.deleteColumns = new int[] { };
            }
            else
            {
                this.deleteColumns = Array.ConvertAll(row.GetCell(4).ToString().Split(','), int.Parse);
            }
        }

        public Boolean needToSkip(int col)
        {
            return (Array.IndexOf(this.deleteColumns, col) != -1);
        }
    }
}
