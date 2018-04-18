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
            //if (cell != null)
            //{
            //    this.textBox64.Text = cell.NumericCellValue.ToString("0.00");
            //}

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
                this.num_Slag_viscosity.Text = cell.NumericCellValue.ToString("0.00");
            }

            //2018.04 新增4个需要公式计算的量
            // 温度写定为1205℃
            int T = 1205;
            // 从表格中获取公式需要的6个量
            double SiO2 = Convert.ToDouble(textBox54.Text);
            double CaO = Convert.ToDouble(textBox53.Text);
            double MgO = Convert.ToDouble(textBox52.Text);
            double Al2O3 = Convert.ToDouble(textBox51.Text);
            double Fe3O4 = Convert.ToDouble(textBox50.Text);
            double FeO = Convert.ToDouble(textBox56.Text);

            double amount = SiO2 + CaO + MgO + Al2O3 + Fe3O4 + FeO;
            //冰铜密度计算
            double Copper_density;
            Copper_density = 6.358 - 0.0763 * Convert.ToDouble(textBox31.Text) + 9.940E-3 * Convert.ToDouble(textBox33.Text) - 4.645E-4 * (T - 1000);
            num_Copper_density.Text = Copper_density.ToString("0.00");
            
            //熔化温度计算
            double Melting_T;
            Melting_T = 1309364.70084 - 1308204.22801 * SiO2 / amount - 1307386.12801 * CaO / amount
                - 1308265.96522 * MgO / amount - 1306842.13642 * Al2O3 / amount
                - 1307043.30635 * Fe3O4 / amount - 1308464.79346 * FeO / amount;
            num_Melting_T.Text = Melting_T.ToString("0.00");

            //粘度计算
            double lnA, B, Slag_viscosity;
            lnA = -6.14568 + 63.51280 * SiO2 / amount - 225.59277 * CaO / amount
                - 1464.28910 * MgO / amount + 556.12552 * Al2O3 / amount
                + 111.24994 * Fe3O4 / amount - 115.78691 * FeO / amount;
            B = 37543.65975 - 97469.98881 * SiO2 / amount + 237436.61511 * CaO / amount
                + 1739247.42661 * MgO / amount - 688152.29915 * Al2O3 / amount
                - 157198.36053 * Fe3O4 / amount + 96857.78733 * FeO / amount;
            Slag_viscosity = Math.Exp(lnA + B / T);
            num_Slag_viscosity.Text = Slag_viscosity.ToString("0.00");

            //密度计算
            double density;
            density = 5 - 0.03 * (Convert.ToDouble(textBox44.Text) + Convert.ToDouble(textBox40.Text)* 160 / 232)  
                - 0.02 * (Convert.ToDouble(textBox43.Text) + Convert.ToDouble(textBox42.Text) + Convert.ToDouble(textBox41.Text)) - 0.01 * (T - 1200);
            num_density.Text = density.ToString("0.00");
        }
    }
}
