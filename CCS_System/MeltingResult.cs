using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace CCS_System
{
    public partial class MeltingResult : Form
    {
        private string path;
        public MeltingResult(string path)
        {
            InitializeComponent();
            this.path = path;
            ReadFromExcel();
        }

        // 从Excel读取成分，显示在面板中
        private void ReadFromExcel()
        {
            string tempPath = path;
            HSSFWorkbook wk = null;
            using (FileStream fs = File.Open(tempPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                wk = new HSSFWorkbook(fs);
                fs.Close();
            }
            //打开第四个工作表
            ISheet tb = wk.GetSheetAt(3);
            // 公式自动重新计算
            tb.ForceFormulaRecalculation = true;
            // 开始逐步读取单元格填充数据
            IRow row = null;
            ICell cell = null;
            row = tb.GetRow(4);
            cell = row.GetCell(2);
            this.textBox1.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(4);
            cell = row.GetCell(3);
            this.textBox2.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(4);
            cell = row.GetCell(4);
            this.textBox3.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(4);
            cell = row.GetCell(5);
            this.textBox4.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(4);
            cell = row.GetCell(6);
            this.textBox5.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(4);
            cell = row.GetCell(7);
            this.textBox6.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(4);
            cell = row.GetCell(8);
            this.textBox7.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(4);
            cell = row.GetCell(9);
            this.textBox8.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(4);
            cell = row.GetCell(10);
            this.textBox9.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(4);
            cell = row.GetCell(11);
            this.textBox10.Text = cell.NumericCellValue.ToString("0.00");

            row = tb.GetRow(5);
            cell = row.GetCell(2);
            this.textBox20.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(5);
            cell = row.GetCell(3);
            this.textBox19.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(5);
            cell = row.GetCell(4);
            this.textBox18.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(5);
            cell = row.GetCell(5);
            this.textBox17.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(5);
            cell = row.GetCell(6);
            this.textBox16.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(5);
            cell = row.GetCell(7);
            this.textBox15.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(5);
            cell = row.GetCell(8);
            this.textBox14.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(5);
            cell = row.GetCell(9);
            this.textBox13.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(5);
            cell = row.GetCell(10);
            this.textBox12.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(5);
            cell = row.GetCell(11);
            this.textBox11.Text = cell.NumericCellValue.ToString("0.00");

            row = tb.GetRow(6);
            cell = row.GetCell(2);
            this.textBox30.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(6);
            cell = row.GetCell(3);
            this.textBox29.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(6);
            cell = row.GetCell(4);
            this.textBox28.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(6);
            cell = row.GetCell(5);
            this.textBox27.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(6);
            cell = row.GetCell(6);
            this.textBox26.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(6);
            cell = row.GetCell(7);
            this.textBox25.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(6);
            cell = row.GetCell(8);
            this.textBox24.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(6);
            cell = row.GetCell(9);
            this.textBox23.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(6);
            cell = row.GetCell(10);
            this.textBox22.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(6);
            cell = row.GetCell(11);
            this.textBox21.Text = cell.NumericCellValue.ToString("0.00");

            // 冰铜
            row = tb.GetRow(9);
            cell = row.GetCell(1);
            this.textBox38.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(9);
            cell = row.GetCell(10);
            this.textBox37.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(9);
            cell = row.GetCell(11);
            this.textBox36.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(9);
            cell = row.GetCell(12);
            this.textBox35.Text = cell.NumericCellValue.ToString("0.00");

            row = tb.GetRow(10);
            cell = row.GetCell(1);
            this.textBox34.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(10);
            cell = row.GetCell(10);
            this.textBox33.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(10);
            cell = row.GetCell(11);
            this.textBox32.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(10);
            cell = row.GetCell(12);
            this.textBox31.Text = cell.NumericCellValue.ToString("0.00");

            // 熔渣
            row = tb.GetRow(23);
            cell = row.GetCell(1);
            this.textBox58.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(23);
            cell = row.GetCell(2);
            this.textBox57.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(23);
            cell = row.GetCell(3);
            this.textBox56.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(23);
            cell = row.GetCell(4);
            this.textBox55.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(23);
            cell = row.GetCell(5);
            this.textBox54.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(23);
            cell = row.GetCell(6);
            this.textBox53.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(23);
            cell = row.GetCell(7);
            this.textBox52.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(23);
            cell = row.GetCell(8);
            this.textBox51.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(23);
            cell = row.GetCell(9);
            this.textBox50.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(23);
            cell = row.GetCell(10);
            this.textBox49.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(23);
            cell = row.GetCell(11);
            this.textBox60.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(23);
            cell = row.GetCell(12);
            this.textBox62.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(23);
            cell = row.GetCell(13);
            if (cell != null)
            {
                this.textBox64.Text = cell.NumericCellValue.ToString("0.00");
            }

            row = tb.GetRow(24);
            cell = row.GetCell(1);
            this.textBox48.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(24);
            cell = row.GetCell(2);
            this.textBox47.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(24);
            cell = row.GetCell(3);
            this.textBox46.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(24);
            cell = row.GetCell(4);
            this.textBox45.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(24);
            cell = row.GetCell(5);
            this.textBox44.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(24);
            cell = row.GetCell(6);
            this.textBox43.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(24);
            cell = row.GetCell(7);
            this.textBox42.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(24);
            cell = row.GetCell(8);
            this.textBox41.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(24);
            cell = row.GetCell(9);
            this.textBox40.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(24);
            cell = row.GetCell(10);
            this.textBox39.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(24);
            cell = row.GetCell(11);
            this.textBox59.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(24);
            cell = row.GetCell(12);
            this.textBox61.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(24);
            cell = row.GetCell(13);
            if(cell != null)
            {
                this.textBox63.Text = cell.NumericCellValue.ToString("0.00");
            }
        }
    }
}
