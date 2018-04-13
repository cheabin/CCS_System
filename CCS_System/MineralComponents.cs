using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace CCS_System
{
    public partial class MineralComponents : Form
    {
        int counts;
        // 保存原始数据
        string[,] originaldata = null;
        ExpertSystem es;
        int[] storenum = new int[] { 1, 2, 3, 4, 8, 9, 10 };
        public MineralComponents()
        {
            InitializeComponent();
            int index = 0;
            while (index < 10)
            {
                // 填充表格
                index = this.dataGridView1.Rows.Add();
                if (index < 7)
                    this.dataGridView1.Rows[index].Cells[0].Value = storenum[index];
                else
                    this.dataGridView1.Rows[index].Cells[0].Value = "";
                this.dataGridView1.Rows[index].Cells[1].Value = "";
                this.dataGridView1.Rows[index].Cells[2].Value = "";
                this.dataGridView1.Rows[index].Cells[3].Value = "";
                this.dataGridView1.Rows[index].Cells[4].Value = "";
                this.dataGridView1.Rows[index].Cells[5].Value = "";
                this.dataGridView1.Rows[index].Cells[6].Value = "";
                this.dataGridView1.Rows[index].Cells[7].Value = "";
                this.dataGridView1.Rows[index].Cells[8].Value = "";
                this.dataGridView1.Rows[index].Cells[9].Value = "";
            }
        }

        public MineralComponents(List<string> minerals, ExpertSystem parent)
        {
            InitializeComponent();
            // 引用父类窗体实例
            this.es = parent;
            // 获取总行数
            counts = minerals.Count;
            // 第一列单元格只读
            int index = 0;
            while (index < 10)
            {
                // 填充表格
                index = this.dataGridView1.Rows.Add();
                if (index < 7)
                    this.dataGridView1.Rows[index].Cells[0].Value = storenum[index];
                else
                    this.dataGridView1.Rows[index].Cells[0].Value = "";
                if (index < minerals.Count)
                    this.dataGridView1.Rows[index].Cells[1].Value = minerals[index];
                else
                    this.dataGridView1.Rows[index].Cells[1].Value = "";
                if ("(无矿)".Equals(this.dataGridView1.Rows[index].Cells[1].Value))
                {
                    this.dataGridView1.Rows[index].Cells[2].Value = "0";
                    this.dataGridView1.Rows[index].Cells[3].Value = "0";
                    this.dataGridView1.Rows[index].Cells[4].Value = "0";
                    this.dataGridView1.Rows[index].Cells[5].Value = "0";
                    this.dataGridView1.Rows[index].Cells[6].Value = "0";
                    this.dataGridView1.Rows[index].Cells[7].Value = "0";
                    this.dataGridView1.Rows[index].Cells[8].Value = "0";
                    this.dataGridView1.Rows[index].Cells[9].Value = "0";
                }
                else
                {
                    if (es.Mdata == null) { 
                        this.dataGridView1.Rows[index].Cells[2].Value = "";
                        this.dataGridView1.Rows[index].Cells[3].Value = "";
                        this.dataGridView1.Rows[index].Cells[4].Value = "";
                        this.dataGridView1.Rows[index].Cells[5].Value = "";
                        this.dataGridView1.Rows[index].Cells[6].Value = "";
                        this.dataGridView1.Rows[index].Cells[7].Value = "";
                        this.dataGridView1.Rows[index].Cells[8].Value = "";
                        this.dataGridView1.Rows[index].Cells[9].Value = "";
                    }
                    else
                    {
                        if(index < 7)
                        {
                            this.dataGridView1.Rows[index].Cells[2].Value = es.Mdata[index, 2];
                            this.dataGridView1.Rows[index].Cells[3].Value = es.Mdata[index, 3];
                            this.dataGridView1.Rows[index].Cells[4].Value = es.Mdata[index, 4];
                            this.dataGridView1.Rows[index].Cells[5].Value = es.Mdata[index, 5];
                            this.dataGridView1.Rows[index].Cells[6].Value = es.Mdata[index, 6];
                            this.dataGridView1.Rows[index].Cells[7].Value = es.Mdata[index, 7];
                            this.dataGridView1.Rows[index].Cells[8].Value = es.Mdata[index, 8];
                            this.dataGridView1.Rows[index].Cells[9].Value = es.Mdata[index, 9];
                        }
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            originaldata = new string[7, 10];
            // 正则表达式
            string regex = @"^-?\d+\.?\d*$";
            for (int i = 0; i < counts; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    // 2016.11.24添加，防止精矿成分填写不完整而导致的程序崩溃
                    if (this.dataGridView1.Rows[i].Cells[j].Value == null 
                        || "".Equals(this.dataGridView1.Rows[i].Cells[j].Value.ToString()))
                    {
                        MessageBox.Show("精矿成分填写不完整，无法保存数据！（" + (i + 1) + "," + (j + 1) + "）");
                        return;
                    }
                    if (j > 1) // 第二列是精矿名，从第三列再开始判断
                    {
                        // 判断填写的精矿成分是否符合双精度实数规则，不符合则为非法字符（2016.11.24）
                        bool result = Regex.IsMatch(this.dataGridView1.Rows[i].Cells[j].Value.ToString(), regex);
                        if (result == false)
                        {
                            MessageBox.Show("精矿成分存在非法字符，无法保存数据！（" + (i + 1) + "," + (j + 1) + "）");
                            return;
                        }
                    }
                    originaldata[i, j] = this.dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
            es.Mdata = originaldata;
            MessageBox.Show("数据保存成功！");
            es.button3Info = "查看精矿成分";
            es.calcNewComponents();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int idx = this.es.selectedIndex;
            // 下标控制参数
            for (int i = 0; i<7; i++)
            {
                for(int j = 0; j< 8; j++)
                {
                    this.dataGridView1.Rows[i].Cells[j + 2].Value = es.CompData[idx, (i + 1) * 8 + j];
                }
            }
        }
    }
}
