using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Data.OleDb;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Diagnostics;
using System.Threading;

namespace CCS_System
{

    public partial class ExpertSystem : Form
    {
        #region 全局变量
        double[,] inPut;//数据
        int sub;//特征值数
        int length = 0;// 总长度
        bool isbutton3 = false;//导出配料单时的判断标志
        bool isbutton4 = false;//导出配料单时的判断标志
        #endregion

        #region 构造函数
        public ExpertSystem()
        {
            InitializeComponent();
            GetData();
            init();
            sub = 4;
        }
        #endregion

        #region 初始化程序所需文件
        public void init()
        {
            try
            {
                // 读取FactSage计算所需文件（ExpertSystem文件夹）释放日志
                string[] str = File.ReadAllLines("FactSageFilesRelease.log");
            }
            catch (FileNotFoundException)
            {
                // 第一次运行
                CopyDir("FactSageCalcData\\", "D:\\ExpertSystem");
                MessageBox.Show("FactSage计算文件创建成功！\n路径：D:\\ExpertSystem", "提示信息");
                string[] str = { "D:\\ExpertSystem", DateTime.Now.ToString() };
                // 创建FactSage计算所需文件（ExpertSystem文件夹）释放日志
                File.WriteAllLines("FactSageFilesRelease.log", str);
            }
            try
            {
                // 读取FactSage路径配置文件
                string str = File.ReadAllText("FactSagePath.dat");
                // 将路径显示在程序界面中
                this.textBox57.Text = str;
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show("FactSage安装路径尚未设置，程序将无法完成运算！\n请进入系统后点击左上方【系统维护】中的【更改】按钮完成设置！",
                    "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // 文件目录拷贝函数，用以将srcPath中的所有文件和文件夹拷贝到aimPath中
        private void CopyDir(string srcPath, string aimPath)
        {
            try
            {
                // 检查目标目录是否以目录分割字符结束如果不是则添加
                if (aimPath[aimPath.Length - 1] != Path.DirectorySeparatorChar)
                {
                    aimPath += Path.DirectorySeparatorChar;
                }
                // 判断目标目录是否存在如果不存在则新建
                if (!Directory.Exists(aimPath))
                {
                    Directory.CreateDirectory(aimPath);
                }
                // 得到源目录的文件列表，该里面是包含文件以及目录路径的一个数组
                string[] fileList = Directory.GetFileSystemEntries(srcPath);
                // 遍历所有的文件和目录
                foreach (string file in fileList)
                {
                    if (Directory.Exists(file))
                    {
                        CopyDir(file, aimPath + Path.GetFileName(file));
                    }
                    else
                    {
                        File.Copy(file, aimPath + Path.GetFileName(file), true);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "异常信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region 设置FactSage目录和生成CallFactSageEquilib.bat
        private void button1_Click(object sender, EventArgs e)
        {
            SetFactSagePath();
        }

        private void SetFactSagePath()
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择FactSage安装路径（本操作只需要执行一次）\n注意：请确保路径准确有效，否则程序将无法执行！";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string foldPath = dialog.SelectedPath;
                MessageBox.Show("FactSage安装路径已设为:" + foldPath, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.textBox57.Text = foldPath;
                // 将FactSage路径写入文件，下次直接进行读取
                File.WriteAllText("FactSagePath.dat", foldPath);
                string filePath = foldPath + "\\Run_Equilib_ISA.bat";
                // 拷贝Run_Equilib_ISA.bat文件到FactSage根目录下
                File.Copy("Run_Equilib_ISA.bat", filePath, true);
                MessageBox.Show("Run_Equilib_ISA.bat已经拷贝到：" + filePath);
                // 根据指定的路径，编写CallFactSageEquilib.bat文件
                // 分割路径
                string[] pathArray = filePath.Split('\\');
                // 在第2行到n-1行，都是要cd进去的
                for (int i = 1; i < pathArray.Length - 1; i++)
                {
                    pathArray[i] = "cd " + pathArray[i];
                }
                // 最后一行要用call
                pathArray[pathArray.Length - 1] = "call " + pathArray[pathArray.Length - 1];
                // 最大支持长度为63的路径，用于重新整合分割的路径
                string[] extraPathArray = new string[63];
                // 重新整合路径
                extraPathArray[0] = pathArray[0];
                // 解决相同目录下直接输入盘符无法回到根目录的问题
                extraPathArray[1] = "cd\\";
                // 将pathArray中其余的元素写入extraPathArray
                for (int i = 1; i < pathArray.Length; i++)
                {
                    extraPathArray[i + 1] = pathArray[i];
                }
                // 创建CallFactSageEquilib.bat
                File.WriteAllLines("CallFactSageEquilib.bat", extraPathArray);
                MessageBox.Show("CallFactSageEquilib.bat文件已创建！");
            }
        }
        #endregion

        #region 编辑FactSage所用的ISA.mac
        private void editMacroFile(string filename)
        {
            string[] str = File.ReadAllLines("FactSageCalcData\\ISA.mac", Encoding.GetEncoding("gb2312"));
            str[3] = "OLE1 " + filename + " Sheet2					//读取的Excel文件名";
            str[4] = "OLE2 " + filename + " Sheet3					//输出的Excel文件名";
            File.WriteAllLines("D:\\ExpertSystem\\ISA.mac", str, Encoding.GetEncoding("gb2312"));
            progressWindowChange("ISA.mac文件重写成功！");
        }
        #endregion

        #region 从数据库获取有关4个需求的基础数据
        // MySQL数据库操作 

        //public void GetData()
        //{
        //    string constr = "server=localhost;User Id=ISA;password=123456;Database=ccs";
        //    MySqlConnection mycon = new MySqlConnection(constr);
        //    mycon.Open();
        //    MySqlCommand mycmd = new MySqlCommand("SELECT * FROM production WHERE factMagIron <> 0", mycon);
        //    // MySqlCommand mycmd = new MySqlCommand("SELECT * FROM production", mycon);
        //    MySqlDataReader reader = mycmd.ExecuteReader();
        //    try
        //    {
        //        // 用5个纬度保存数据，4个参与聚类，1个用于标识配料单
        //        inPut = new double[10000, 5];
        //        int index = 0;
        //        while (reader.Read())
        //        {
        //            if (reader.HasRows)
        //            {
        //                inPut[index, 0] = reader.GetDouble("factMatte");
        //                inPut[index, 1] = reader.GetDouble("factMagIron");
        //                inPut[index, 2] = reader.GetDouble("factSiO2Fe");
        //                inPut[index, 3] = reader.GetDouble("factSiO2CaO");
        //                inPut[index, 4] = reader.GetDouble("id");
        //                index++;
        //            }
        //        }
        //        length = index;
        //        this.groupBox2.Text = "系统维护（MySQL在线数据库，" + length + "条数据）";
        //    }
        //    catch (Exception)
        //    {
        //        MessageBox.Show("【系统维护】数据库初始化失败！");
        //    }
        //    finally
        //    {
        //        reader.Close();
        //    }
        //    mycon.Close();
        //}

        // 8.23添加Access数据库操作，方便离线测试

        public void GetData()
        {
            string mdbPath = @"ccs.mdb";
            string strConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdbPath;
            OleDbConnection odcConnection = new OleDbConnection(strConn);
            odcConnection.Open();
            OleDbCommand odCommand = odcConnection.CreateCommand();
            odCommand.CommandText = "SELECT * FROM production WHERE factMagIron <> '0.00'";
            // odCommand.CommandText = "SELECT * FROM production";
            OleDbDataReader odrReader = odCommand.ExecuteReader();
            // 用5个纬度保存数据，4个参与聚类，1个用于标识配料单【由于数据量很小无需聚类，因此聚类已经撤销，直接检索】
            inPut = new double[10000, 5];
            int index = 0;
            while (odrReader.Read())
            {
                if (odrReader.HasRows)
                {
                    inPut[index, 0] = Convert.ToDouble(odrReader["factMatte"]);
                    inPut[index, 1] = Convert.ToDouble(odrReader["factMagIron"]);
                    inPut[index, 2] = Convert.ToDouble(odrReader["factSiO2Fe"]);
                    inPut[index, 3] = Convert.ToDouble(odrReader["factSiO2CaO"]);
                    inPut[index, 4] = Convert.ToDouble(odrReader["id"]);
                    index++;
                }
            }
            length = index;
            this.groupBox2.Text = "系统维护（Access离线数据库，" + length + "条数据）";
            odrReader.Close();
            odcConnection.Close();
        }

        #endregion

        #region 配料单初步查询（仅根据4项需求）
        // 存放结果点
        int[] resultdot = new int[5000];
        double totaldosage = 0;
        // 存储checkBox选中的精矿
        List<string> minerals = new List<string>();
        private void mine_recommend_Click(object sender, EventArgs e)
        {
            if ("".Equals(this.textBox1.Text.ToString()) || "".Equals(this.textBox2.Text.ToString())
                || "".Equals(this.textBox3.Text.ToString()) || "".Equals(this.textBox4.Text.ToString())
                || "".Equals(this.textBox37.Text.ToString()) || "".Equals(this.textBox54.Text.ToString())
                || "".Equals(this.textBox55.Text.ToString()))
            {
                MessageBox.Show("请完整地输入数据！");
                return;
            }
            int checknum = 0;
            minerals.Clear();
            foreach (Control c in groupBox7.Controls)
            {
                if (c is CheckBox && ((CheckBox)c).Checked == true)
                {
                    checknum++;
                    minerals.Add(c.Text.ToString());
                }
            }
            // 加入“无矿”
            minerals.Add("(无矿)");
            if (checknum < 1)
            {
                MessageBox.Show("请至少选择1种精矿！");
                return;
            }

            // 存储总下料量
            totaldosage = Convert.ToDouble(this.textBox37.Text.ToString());
            double[] target = new double[4];
            target[0] = Convert.ToDouble(this.textBox1.Text.ToString());
            target[1] = Convert.ToDouble(this.textBox2.Text.ToString());
            target[2] = Convert.ToDouble(this.textBox3.Text.ToString());
            target[3] = Convert.ToDouble(this.textBox4.Text.ToString());

            // 用于存放距离
            double[] distance = new double[5000];
            // 用于对应距离数组顺序存放点编号
            int[] rdot = new int[5000];
            int dindex = 0;
            // 遍历所有点，看离哪个点最近
            for (int j = 0; j < length; j++)
            {
                double tmpDis = 0.0;
                // 计算到每个点的距离
                for (int m = 0; m < sub; m++)
                {
                    tmpDis += Math.Pow((target[m] - inPut[j, m]), 2);
                }
                if (dindex == 0)
                {
                    distance[dindex] = tmpDis;
                    rdot[dindex] = j;
                    dindex++;
                }
                else
                {
                    // 直接插入排序
                    bool flag = false;
                    for (int i = 0; i < dindex; i++)
                    {
                        if (tmpDis < distance[i])
                        {
                            flag = true;
                            for (int m = dindex; m >= i; m--)
                            {
                                distance[m + 1] = distance[m];
                                rdot[m + 1] = rdot[m];
                            }
                            distance[i] = tmpDis;
                            // 记录点编号
                            rdot[i] = j;
                            break;
                        }
                    }
                    // 无法插入，放到队列尾部
                    if (flag == false)
                    {
                        distance[dindex] = tmpDis;
                        rdot[dindex] = j;
                    }
                    dindex++;
                }
                for (int i = 0; i < dindex; i++)
                {
                    // 结果点集（将组中点的编号转化为所有点对应的编号）
                    resultdot[i] = rdot[i];
                }
            }
            SearchData();
        }
        #endregion

        #region 配矿推荐数据的进一步获取（基于4个需求）
        // 全局二维数组，用于存放5条查出的记录
        string[,] realdata = new string[5, 4];
        string[,] productdata = new string[5, 22];
        // 现在需要存储详细的精矿成分信息
        string[,] productComponentData = new string[5, 64];
        string[,] ingredientdata = new string[5, 7];
        string[] numstr = new string[5];
        // 熔炼状况评分5个初始值
        string[] level = { "Unknown", "Unknown", "Unknown", "Unknown", "Unknown" };

        public string[,] CompData
        {
            get { return this.productComponentData; }
            set { this.productComponentData = value; }
        }
        // 搜索数据库函数
        private void SearchData()
        {
            int n = resultdot.GetLength(0);
            int[] id = new int[5];
            // 用于标明有几个符合条件的id
            int idindex = 0;
            ///*******************************↓↓↓ MySQL查询部分 ↓↓↓*******************************/

            //string constr = "server=localhost;User Id=ISA;password=123456;Database=ccs";
            //MySqlConnection mycon = new MySqlConnection(constr);
            //mycon.Open();
            //this.comboBox1.Items.Clear();
            //for (int i = 0; i < n; i++)
            //{
            //    // 最多存储5个
            //    if (idindex >= 5) break;
            //    id[idindex] = (int)inPut[resultdot[i], 4];
            //    realdata[idindex, 0] = inPut[resultdot[i], 0].ToString();
            //    realdata[idindex, 1] = inPut[resultdot[i], 1].ToString();
            //    realdata[idindex, 2] = inPut[resultdot[i], 2].ToString();
            //    realdata[idindex, 3] = inPut[resultdot[i], 3].ToString();
            //    // 读取下料情况表
            //    MySqlCommand mycmd = new MySqlCommand("SELECT * FROM production WHERE id = " + id[idindex], mycon);
            //    MySqlDataReader reader = mycmd.ExecuteReader();
            //    try
            //    {
            //        while (reader.Read())
            //        {
            //            if (reader.HasRows)
            //            {
            //                // 读取必要信息（下料量）
            //                productdata[idindex, 0] = reader.GetString("h1used");
            //                productdata[idindex, 1] = reader.GetString("h2used");
            //                productdata[idindex, 2] = reader.GetString("h3used");
            //                productdata[idindex, 3] = reader.GetString("h4used");
            //                productdata[idindex, 4] = reader.GetString("h8used");
            //                productdata[idindex, 5] = reader.GetString("h9used");
            //                productdata[idindex, 6] = reader.GetString("h10used");
            //                // 读取“喷枪端压、取样时氧料比、取样时小时料量、测样时的风量”
            //                productdata[idindex, 7] = reader.GetString("nozzle_pressure");
            //                productdata[idindex, 8] = reader.GetString("fuel_ratio");
            //                productdata[idindex, 9] = reader.GetString("amountPerHour");
            //                productdata[idindex, 10] = reader.GetString("air_volume");
            //                // 补钙
            //                productdata[idindex, 11] = reader.GetString("h5used");
            //                // 补硅
            //                productdata[idindex, 12] = reader.GetString("h6used");
            //                // 补煤
            //                productdata[idindex, 13] = reader.GetString("h7used");
            //                productdata[idindex, 14] = reader.GetString("number");
            //                // 熔炼状况等级
            //                productdata[idindex, 15] = reader.GetString("grade");
            //                // 2017.10 分配系数是否计算
            //                productdata[idindex, 16] = reader.GetString("iscalculation");
            //                productdata[idindex, 17] = reader.GetString("id");

            //                productdata[idindex, 18] = reader.GetString("p1");
            //                productdata[idindex, 19] = reader.GetString("p2");
            //                productdata[idindex, 20] = reader.GetString("p3");
            //                productdata[idindex, 21] = reader.GetString("p4");

            //            }
            //        }
            //        double tmptotal = Convert.ToDouble(productdata[idindex, 0]) + Convert.ToDouble(productdata[idindex, 1])
            //                            + Convert.ToDouble(productdata[idindex, 2]) + Convert.ToDouble(productdata[idindex, 3])
            //                            + Convert.ToDouble(productdata[idindex, 4]) + Convert.ToDouble(productdata[idindex, 5])
            //                            + Convert.ToDouble(productdata[idindex, 6]);
            //        // 超出规定的精矿总量，放弃此结果，继续遍历下一个点
            //        // 下料量调节功能已经产生，因此不必要再注意总下料量
            //        // if (tmptotal > totaldosage) continue;
            //    }
            //    catch (Exception)
            //    {
            //        MessageBox.Show("【配矿推荐1】数据库搜索失败！");
            //    }
            //    finally
            //    {
            //        reader.Close();
            //    }
            //    // 读取配料表
            //    mycmd = new MySqlCommand("SELECT * FROM ingredient WHERE number = '" + productdata[idindex, 14] + "'", mycon);
            //    reader = mycmd.ExecuteReader();
            //    bool isContained = true;
            //    try
            //    {
            //        while (reader.Read())
            //        {
            //            if (reader.HasRows)
            //            {
            //                // 读取必要信息（精矿名）
            //                ingredientdata[idindex, 0] = reader.GetString("NO1_name");
            //                ingredientdata[idindex, 1] = reader.GetString("NO2_name");
            //                ingredientdata[idindex, 2] = reader.GetString("NO3_name");
            //                ingredientdata[idindex, 3] = reader.GetString("NO4_name");
            //                ingredientdata[idindex, 4] = reader.GetString("NO8_name");
            //                ingredientdata[idindex, 5] = reader.GetString("NO9_name");
            //                ingredientdata[idindex, 6] = reader.GetString("NO10_name");
            //                // 读取精矿成分信息
            //                productComponentData[idindex, 0] = reader.GetString("Con_Cu");
            //                productComponentData[idindex, 1] = reader.GetString("Con_Fe");
            //                productComponentData[idindex, 2] = reader.GetString("Con_S");
            //                productComponentData[idindex, 3] = reader.GetString("Con_SiO2");
            //                productComponentData[idindex, 4] = reader.GetString("Con_CaO");
            //                productComponentData[idindex, 5] = reader.GetString("Con_MgO");
            //                productComponentData[idindex, 6] = reader.GetString("Con_Al2O3");
            //                productComponentData[idindex, 7] = reader.GetString("Con_Co");
            //                productComponentData[idindex, 8] = reader.GetString("NO1_Cu");
            //                productComponentData[idindex, 9] = reader.GetString("NO1_Fe");
            //                productComponentData[idindex, 10] = reader.GetString("NO1_S");
            //                productComponentData[idindex, 11] = reader.GetString("NO1_SiO2");
            //                productComponentData[idindex, 12] = reader.GetString("NO1_CaO");
            //                productComponentData[idindex, 13] = reader.GetString("NO1_MgO");
            //                productComponentData[idindex, 14] = reader.GetString("NO1_Al2O3");
            //                productComponentData[idindex, 15] = reader.GetString("NO1_Co");
            //                productComponentData[idindex, 16] = reader.GetString("NO2_Cu");
            //                productComponentData[idindex, 17] = reader.GetString("NO2_Fe");
            //                productComponentData[idindex, 18] = reader.GetString("NO2_S");
            //                productComponentData[idindex, 19] = reader.GetString("NO2_SiO2");
            //                productComponentData[idindex, 20] = reader.GetString("NO2_CaO");
            //                productComponentData[idindex, 21] = reader.GetString("NO2_MgO");
            //                productComponentData[idindex, 22] = reader.GetString("NO2_Al2O3");
            //                productComponentData[idindex, 23] = reader.GetString("NO2_Co");
            //                productComponentData[idindex, 24] = reader.GetString("NO3_Cu");
            //                productComponentData[idindex, 25] = reader.GetString("NO3_Fe");
            //                productComponentData[idindex, 26] = reader.GetString("NO3_S");
            //                productComponentData[idindex, 27] = reader.GetString("NO3_SiO2");
            //                productComponentData[idindex, 28] = reader.GetString("NO3_CaO");
            //                productComponentData[idindex, 29] = reader.GetString("NO3_MgO");
            //                productComponentData[idindex, 30] = reader.GetString("NO3_Al2O3");
            //                productComponentData[idindex, 31] = reader.GetString("NO3_Co");
            //                productComponentData[idindex, 32] = reader.GetString("NO4_Cu");
            //                productComponentData[idindex, 33] = reader.GetString("NO4_Fe");
            //                productComponentData[idindex, 34] = reader.GetString("NO4_S");
            //                productComponentData[idindex, 35] = reader.GetString("NO4_SiO2");
            //                productComponentData[idindex, 36] = reader.GetString("NO4_CaO");
            //                productComponentData[idindex, 37] = reader.GetString("NO4_MgO");
            //                productComponentData[idindex, 38] = reader.GetString("NO4_Al2O3");
            //                productComponentData[idindex, 39] = reader.GetString("NO4_Co");
            //                productComponentData[idindex, 40] = reader.GetString("NO8_Cu");
            //                productComponentData[idindex, 41] = reader.GetString("NO8_Fe");
            //                productComponentData[idindex, 42] = reader.GetString("NO8_S");
            //                productComponentData[idindex, 43] = reader.GetString("NO8_SiO2");
            //                productComponentData[idindex, 44] = reader.GetString("NO8_CaO");
            //                productComponentData[idindex, 45] = reader.GetString("NO8_MgO");
            //                productComponentData[idindex, 46] = reader.GetString("NO8_Al2O3");
            //                productComponentData[idindex, 47] = reader.GetString("NO8_Co");
            //                productComponentData[idindex, 48] = reader.GetString("NO9_Cu");
            //                productComponentData[idindex, 49] = reader.GetString("NO9_Fe");
            //                productComponentData[idindex, 50] = reader.GetString("NO9_S");
            //                productComponentData[idindex, 51] = reader.GetString("NO9_SiO2");
            //                productComponentData[idindex, 52] = reader.GetString("NO9_CaO");
            //                productComponentData[idindex, 53] = reader.GetString("NO9_MgO");
            //                productComponentData[idindex, 54] = reader.GetString("NO9_Al2O3");
            //                productComponentData[idindex, 55] = reader.GetString("NO9_Co");
            //                productComponentData[idindex, 56] = reader.GetString("NO10_Cu");
            //                productComponentData[idindex, 57] = reader.GetString("NO10_Fe");
            //                productComponentData[idindex, 58] = reader.GetString("NO10_S");
            //                productComponentData[idindex, 59] = reader.GetString("NO10_SiO2");
            //                productComponentData[idindex, 60] = reader.GetString("NO10_CaO");
            //                productComponentData[idindex, 61] = reader.GetString("NO10_MgO");
            //                productComponentData[idindex, 62] = reader.GetString("NO10_Al2O3");
            //                productComponentData[idindex, 63] = reader.GetString("NO10_Co");
            //            }
            //        }
            //        // 检查配料单中使用的矿是否包含在复选框中
            //        for (int m = 0; m < 7; m++)
            //        {
            //            if (!minerals.Contains(ingredientdata[idindex, m]))
            //            {
            //                isContained = false;
            //                break;
            //            }
            //        }
            //        if (isContained == false) continue;
            //    }
            //    catch (Exception)
            //    {
            //        MessageBox.Show("【配矿推荐2】数据库搜索失败！");
            //    }
            //    finally
            //    {
            //        reader.Close();
            //    }
            //    // 用于后续重新选择配料单时的对照
            //    numstr[idindex] = productdata[idindex, 14] + "(" + id[idindex] + ")等级--" + level[idindex];
            //    this.comboBox1.Items.Add(productdata[idindex, 14] + "(" + id[idindex] + ")等级--" + level[idindex]);
            //    idindex++;
            //}
            //mycon.Close();

            ///*******************************↑↑↑ MySQL查询部分结束 ↑↑↑*******************************/


            /*******************************↓↓↓ Access查询部分 ↓↓↓*******************************/

            string mdbPath = @"ccs.mdb";
            string strConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdbPath;
            OleDbConnection odcConnection = new OleDbConnection(strConn);
            odcConnection.Open();
            this.comboBox1.Items.Clear();
            for (int i = 0; i < n; i++)
            {
                // 最多存储5个
                if (idindex >= 5) break;
                id[idindex] = (int)inPut[resultdot[i], 4];
                realdata[idindex, 0] = inPut[resultdot[i], 0].ToString();
                realdata[idindex, 1] = inPut[resultdot[i], 1].ToString();
                realdata[idindex, 2] = inPut[resultdot[i], 2].ToString();
                realdata[idindex, 3] = inPut[resultdot[i], 3].ToString();
                // 读取下料情况表
                OleDbCommand odCommand = odcConnection.CreateCommand();
                odCommand.CommandText = "SELECT * FROM production WHERE id = " + id[idindex];
                OleDbDataReader odrReader = odCommand.ExecuteReader();
                try
                {
                    while (odrReader.Read())
                    {
                        if (odrReader.HasRows)
                        {
                            // 读取必要信息（下料量）
                            productdata[idindex, 0] = odrReader["h1used"].ToString();
                            productdata[idindex, 1] = odrReader["h2used"].ToString();
                            productdata[idindex, 2] = odrReader["h3used"].ToString();
                            productdata[idindex, 3] = odrReader["h4used"].ToString();
                            productdata[idindex, 4] = odrReader["h8used"].ToString();
                            productdata[idindex, 5] = odrReader["h9used"].ToString();
                            productdata[idindex, 6] = odrReader["h10used"].ToString();
                            // 读取“喷枪端压、取样时氧料比、取样时小时料量、测样时的风量”
                            productdata[idindex, 7] = odrReader["nozzle_pressure"].ToString();
                            productdata[idindex, 8] = odrReader["fuel_ratio"].ToString();
                            productdata[idindex, 9] = odrReader["amountPerHour"].ToString();
                            productdata[idindex, 10] = odrReader["air_volume"].ToString();
                            // 补钙
                            productdata[idindex, 11] = odrReader["h5used"].ToString();
                            // 补硅
                            productdata[idindex, 12] = odrReader["h6used"].ToString();
                            // 补煤
                            productdata[idindex, 13] = odrReader["h7used"].ToString();
                            productdata[idindex, 14] = odrReader["number"].ToString();
                            // 熔炼状况等级
                            productdata[idindex, 15] = odrReader["grade"].ToString();

                            //2017.10 分配系数是否计算
                            productdata[idindex, 16] = odrReader["iscalculation"].ToString();
                            productdata[idindex, 17] = odrReader["id"].ToString();

                            productdata[idindex, 18] = odrReader["p1"].ToString();
                            productdata[idindex, 19] = odrReader["p2"].ToString();
                            productdata[idindex, 20] = odrReader["p3"].ToString();
                            productdata[idindex, 21] = odrReader["p4"].ToString();
                        }
                    }
                    double tmptotal = Convert.ToDouble(productdata[idindex, 0]) + Convert.ToDouble(productdata[idindex, 1])
                                        + Convert.ToDouble(productdata[idindex, 2]) + Convert.ToDouble(productdata[idindex, 3])
                                        + Convert.ToDouble(productdata[idindex, 4]) + Convert.ToDouble(productdata[idindex, 5])
                                        + Convert.ToDouble(productdata[idindex, 6]);
                }
                catch (Exception)
                {
                    MessageBox.Show("【配矿推荐1】数据库搜索失败！");
                }
                finally
                {
                    odrReader.Close();
                }
                // 读取配料表
                odCommand.CommandText = "SELECT * FROM ingredient WHERE number = '" + productdata[idindex, 14] + "'";
                odrReader = odCommand.ExecuteReader();
                bool isContained = true;
                try
                {
                    while (odrReader.Read())
                    {
                        if (odrReader.HasRows)
                        {
                            // 读取必要信息（精矿名）
                            ingredientdata[idindex, 0] = odrReader["NO1_name"].ToString();
                            ingredientdata[idindex, 1] = odrReader["NO2_name"].ToString();
                            ingredientdata[idindex, 2] = odrReader["NO3_name"].ToString();
                            ingredientdata[idindex, 3] = odrReader["NO4_name"].ToString();
                            ingredientdata[idindex, 4] = odrReader["NO8_name"].ToString();
                            ingredientdata[idindex, 5] = odrReader["NO9_name"].ToString();
                            ingredientdata[idindex, 6] = odrReader["NO10_name"].ToString();
                            // 读取精矿成分信息
                            productComponentData[idindex, 0] = odrReader["Con_Cu"].ToString();
                            productComponentData[idindex, 1] = odrReader["Con_Fe"].ToString();
                            productComponentData[idindex, 2] = odrReader["Con_S"].ToString();
                            productComponentData[idindex, 3] = odrReader["Con_SiO2"].ToString();
                            productComponentData[idindex, 4] = odrReader["Con_CaO"].ToString();
                            productComponentData[idindex, 5] = odrReader["Con_MgO"].ToString();
                            productComponentData[idindex, 6] = odrReader["Con_Al2O3"].ToString();
                            productComponentData[idindex, 7] = odrReader["Con_Co"].ToString();
                            productComponentData[idindex, 8] = odrReader["NO1_Cu"].ToString();
                            productComponentData[idindex, 9] = odrReader["NO1_Fe"].ToString();
                            productComponentData[idindex, 10] = odrReader["NO1_S"].ToString();
                            productComponentData[idindex, 11] = odrReader["NO1_SiO2"].ToString();
                            productComponentData[idindex, 12] = odrReader["NO1_CaO"].ToString();
                            productComponentData[idindex, 13] = odrReader["NO1_MgO"].ToString();
                            productComponentData[idindex, 14] = odrReader["NO1_Al2O3"].ToString();
                            productComponentData[idindex, 15] = odrReader["NO1_Co"].ToString();
                            productComponentData[idindex, 16] = odrReader["NO2_Cu"].ToString();
                            productComponentData[idindex, 17] = odrReader["NO2_Fe"].ToString();
                            productComponentData[idindex, 18] = odrReader["NO2_S"].ToString();
                            productComponentData[idindex, 19] = odrReader["NO2_SiO2"].ToString();
                            productComponentData[idindex, 20] = odrReader["NO2_CaO"].ToString();
                            productComponentData[idindex, 21] = odrReader["NO2_MgO"].ToString();
                            productComponentData[idindex, 22] = odrReader["NO2_Al2O3"].ToString();
                            productComponentData[idindex, 23] = odrReader["NO2_Co"].ToString();
                            productComponentData[idindex, 24] = odrReader["NO3_Cu"].ToString();
                            productComponentData[idindex, 25] = odrReader["NO3_Fe"].ToString();
                            productComponentData[idindex, 26] = odrReader["NO3_S"].ToString();
                            productComponentData[idindex, 27] = odrReader["NO3_SiO2"].ToString();
                            productComponentData[idindex, 28] = odrReader["NO3_CaO"].ToString();
                            productComponentData[idindex, 29] = odrReader["NO3_MgO"].ToString();
                            productComponentData[idindex, 30] = odrReader["NO3_Al2O3"].ToString();
                            productComponentData[idindex, 31] = odrReader["NO3_Co"].ToString();
                            productComponentData[idindex, 32] = odrReader["NO4_Cu"].ToString();
                            productComponentData[idindex, 33] = odrReader["NO4_Fe"].ToString();
                            productComponentData[idindex, 34] = odrReader["NO4_S"].ToString();
                            productComponentData[idindex, 35] = odrReader["NO4_SiO2"].ToString();
                            productComponentData[idindex, 36] = odrReader["NO4_CaO"].ToString();
                            productComponentData[idindex, 37] = odrReader["NO4_MgO"].ToString();
                            productComponentData[idindex, 38] = odrReader["NO4_Al2O3"].ToString();
                            productComponentData[idindex, 39] = odrReader["NO4_Co"].ToString();
                            productComponentData[idindex, 40] = odrReader["NO8_Cu"].ToString();
                            productComponentData[idindex, 41] = odrReader["NO8_Fe"].ToString();
                            productComponentData[idindex, 42] = odrReader["NO8_S"].ToString();
                            productComponentData[idindex, 43] = odrReader["NO8_SiO2"].ToString();
                            productComponentData[idindex, 44] = odrReader["NO8_CaO"].ToString();
                            productComponentData[idindex, 45] = odrReader["NO8_MgO"].ToString();
                            productComponentData[idindex, 46] = odrReader["NO8_Al2O3"].ToString();
                            productComponentData[idindex, 47] = odrReader["NO8_Co"].ToString();
                            productComponentData[idindex, 48] = odrReader["NO9_Cu"].ToString();
                            productComponentData[idindex, 49] = odrReader["NO9_Fe"].ToString();
                            productComponentData[idindex, 50] = odrReader["NO9_S"].ToString();
                            productComponentData[idindex, 51] = odrReader["NO9_SiO2"].ToString();
                            productComponentData[idindex, 52] = odrReader["NO9_CaO"].ToString();
                            productComponentData[idindex, 53] = odrReader["NO9_MgO"].ToString();
                            productComponentData[idindex, 54] = odrReader["NO9_Al2O3"].ToString();
                            productComponentData[idindex, 55] = odrReader["NO9_Co"].ToString();
                            productComponentData[idindex, 56] = odrReader["NO10_Cu"].ToString();
                            productComponentData[idindex, 57] = odrReader["NO10_Fe"].ToString();
                            productComponentData[idindex, 58] = odrReader["NO10_S"].ToString();
                            productComponentData[idindex, 59] = odrReader["NO10_SiO2"].ToString();
                            productComponentData[idindex, 60] = odrReader["NO10_CaO"].ToString();
                            productComponentData[idindex, 61] = odrReader["NO10_MgO"].ToString();
                            productComponentData[idindex, 62] = odrReader["NO10_Al2O3"].ToString();
                            productComponentData[idindex, 63] = odrReader["NO10_Co"].ToString();
                        }
                    }
                    // 检查配料单中使用的矿是否包含在复选框中
                    for (int m = 0; m < 7; m++)
                    {
                        if (!minerals.Contains(ingredientdata[idindex, m]))
                        {
                            isContained = false;
                            break;
                        }
                    }
                    if (isContained == false) continue;
                }
                catch (Exception)
                {
                    MessageBox.Show("【配矿推荐2】数据库搜索失败！");
                }
                finally
                {
                    odrReader.Close();
                }

                // 转换“熔炼状况等级”
                if (!"".Equals(productdata[idindex, 15]))
                    level[idindex] = productdata[idindex, 15];
                else
                    level[idindex] = "Unknown";
                // 用于后续重新选择配料单时的对照
                numstr[idindex] = productdata[idindex, 14] + "(" + id[idindex] + ")等级--" + level[idindex];
                this.comboBox1.Items.Add(productdata[idindex, 14] + "(" + id[idindex] + ")等级--" + level[idindex]);
                idindex++;
            }
            odcConnection.Close();

            /********************************↑↑↑ Access查询部分结束 ↑↑↑*******************************/

            if (idindex == 0)
            {
                this.textBox22.Text = "";
                this.textBox21.Text = "";
                this.textBox20.Text = "";
                this.textBox19.Text = "";
                this.textBox5.Text = "";
                this.textBox7.Text = "";
                this.textBox9.Text = "";
                this.textBox11.Text = "";
                this.textBox12.Text = "";
                this.textBox13.Text = "";
                this.textBox17.Text = "";
                this.textBox6.Text = "";
                this.textBox8.Text = "";
                this.textBox10.Text = "";
                this.textBox14.Text = "";
                this.textBox15.Text = "";
                this.textBox16.Text = "";
                this.textBox18.Text = "";
                this.textBox23.Text = "";
                this.textBox24.Text = "";
                this.textBox25.Text = "";
                this.textBox26.Text = "";
                this.textBox27.Text = "";
                this.textBox28.Text = "";
                this.comboBox1.Text = "";
                this.textBox29.Text = "";
                this.textBox30.Text = "";
                this.textBox31.Text = "";
                this.textBox32.Text = "";
                this.textBox33.Text = "";
                this.textBox34.Text = "";
                this.textBox35.Text = "";
                this.textBox36.Text = "";
                this.label33.Text = "";
                this.textBox38.Text = "";
                this.textBox39.Text = "";
                this.textBox40.Text = "";
                this.textBox41.Text = "";
                this.textBox42.Text = "";
                this.textBox43.Text = "";
                this.textBox44.Text = "";
                this.textBox45.Text = "";
                this.textBox46.Text = "";
                this.textBox47.Text = "";
                this.textBox48.Text = "";
                this.textBox49.Text = "";
                this.textBox50.Text = "";
                this.textBox51.Text = "";
                this.textBox52.Text = "";
                this.textBox53.Text = "";
                MessageBox.Show("未找到符合指定条件的配料单");
                return;
            }
            // 解锁面板上的功能按钮
            this.button3.Enabled = true;
            this.button4.Enabled = true;
            this.button5.Enabled = true;
            this.button8.Enabled = true;
            // 信息填写
            this.textBox22.Text = realdata[0, 0];
            this.textBox21.Text = realdata[0, 1];
            this.textBox20.Text = realdata[0, 2];
            this.textBox19.Text = realdata[0, 3];
            this.textBox5.Text = ingredientdata[0, 0];
            this.textBox7.Text = ingredientdata[0, 1];
            this.textBox9.Text = ingredientdata[0, 2];
            this.textBox11.Text = ingredientdata[0, 3];
            this.textBox12.Text = ingredientdata[0, 4];
            this.textBox13.Text = ingredientdata[0, 5];
            this.textBox17.Text = ingredientdata[0, 6];
            this.textBox6.Text = productdata[0, 0];
            this.textBox8.Text = productdata[0, 1];
            this.textBox10.Text = productdata[0, 2];
            this.textBox14.Text = productdata[0, 3];
            this.textBox15.Text = productdata[0, 4];
            this.textBox16.Text = productdata[0, 5];
            this.textBox18.Text = productdata[0, 6];
            this.textBox23.Text = productdata[0, 7];
            this.textBox24.Text = productdata[0, 8];
            this.textBox25.Text = productdata[0, 9];
            this.textBox26.Text = productdata[0, 10];
            this.textBox27.Text = productdata[0, 11];
            this.textBox28.Text = productdata[0, 12];
            this.textBox45.Text = productdata[0, 13];
            this.comboBox1.Text = productdata[0, 14] + "(" + id[0] + ")等级--" + level[0];
            this.label47.Text = level[0];
            // 原配料单总下料量计算
            double total = Convert.ToDouble(productdata[0, 0]) + Convert.ToDouble(productdata[0, 1])
                                        + Convert.ToDouble(productdata[0, 2]) + Convert.ToDouble(productdata[0, 3])
                                        + Convert.ToDouble(productdata[0, 4]) + Convert.ToDouble(productdata[0, 5])
                                        + Convert.ToDouble(productdata[0, 6]);
            // 定义原始模块输入成分，用于对应该条下料情况记录
            double oriConCu = 0;
            double oriConFe = 0;
            double oriConS = 0;
            double oriConSiO2 = 0;
            double oriConCaO = 0;
            double oriConMgO = 0;
            double oriConAl2O3 = 0;
            double oriConCo = 0;
            for (int i = 8; i < 64; i += 8)
            {
                oriConCu += Convert.ToDouble(productComponentData[0, i + 0]) * Convert.ToDouble(productdata[0, i / 8 - 1]);
                oriConFe += Convert.ToDouble(productComponentData[0, i + 1]) * Convert.ToDouble(productdata[0, i / 8 - 1]);
                oriConS += Convert.ToDouble(productComponentData[0, i + 2]) * Convert.ToDouble(productdata[0, i / 8 - 1]);
                oriConSiO2 += Convert.ToDouble(productComponentData[0, i + 3]) * Convert.ToDouble(productdata[0, i / 8 - 1]);
                oriConCaO += Convert.ToDouble(productComponentData[0, i + 4]) * Convert.ToDouble(productdata[0, i / 8 - 1]);
                oriConMgO += Convert.ToDouble(productComponentData[0, i + 5]) * Convert.ToDouble(productdata[0, i / 8 - 1]);
                oriConAl2O3 += Convert.ToDouble(productComponentData[0, i + 6]) * Convert.ToDouble(productdata[0, i / 8 - 1]);
                oriConCo += Convert.ToDouble(productComponentData[0, i + 7]) * Convert.ToDouble(productdata[0, i / 8 - 1]);
            }
            // 计算原始模块输入成分
            oriConCu /= total;
            oriConFe /= total;
            oriConS /= total;
            oriConSiO2 /= total;
            oriConCaO /= total;
            oriConMgO /= total;
            oriConAl2O3 /= total;
            oriConCo /= total;
            // 填写原始模块输入成分
            this.textBox29.Text = oriConCu.ToString("0.00");
            this.textBox30.Text = oriConFe.ToString("0.00");
            this.textBox31.Text = oriConS.ToString("0.00");
            this.textBox32.Text = oriConSiO2.ToString("0.00");
            this.textBox33.Text = oriConCaO.ToString("0.00");
            this.textBox34.Text = oriConMgO.ToString("0.00");
            this.textBox35.Text = oriConAl2O3.ToString("0.00");
            this.textBox36.Text = oriConCo.ToString("0.00");
            // 显示原配料单总下料量
            this.label33.Text = "原配料单总下料量：" + total + " T/h";
            // 记录预期下料量/匹配配料单下料量的比率以做调整
            double rate = totaldosage / total;
            this.textBox38.Text = (Convert.ToDouble(productdata[0, 0]) * rate).ToString("0.00");
            this.textBox39.Text = (Convert.ToDouble(productdata[0, 1]) * rate).ToString("0.00");
            this.textBox40.Text = (Convert.ToDouble(productdata[0, 2]) * rate).ToString("0.00");
            this.textBox41.Text = (Convert.ToDouble(productdata[0, 3]) * rate).ToString("0.00");
            this.textBox42.Text = (Convert.ToDouble(productdata[0, 4]) * rate).ToString("0.00");
            this.textBox43.Text = (Convert.ToDouble(productdata[0, 5]) * rate).ToString("0.00");
            this.textBox44.Text = (Convert.ToDouble(productdata[0, 6]) * rate).ToString("0.00");
        }
        #endregion

        #region 选择配料单相应事件

        private int flag = 0;

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int n = 0;
            // 寻找选择的配料单
            for (int i = 0; i < 5; i++)
            {
                if (this.comboBox1.SelectedItem.ToString().Equals(numstr[i]))
                {
                    n = i;
                    break;
                }
            }
            this.label47.Text = level[n];
            this.label34.Text = @"修正下料量";
            this.label34.ForeColor = Color.Black;
            // 重填数据
            this.textBox22.Text = realdata[n, 0];
            this.textBox21.Text = realdata[n, 1];
            this.textBox20.Text = realdata[n, 2];
            this.textBox19.Text = realdata[n, 3];
            this.textBox5.Text = ingredientdata[n, 0];
            this.textBox7.Text = ingredientdata[n, 1];
            this.textBox9.Text = ingredientdata[n, 2];
            this.textBox11.Text = ingredientdata[n, 3];
            this.textBox12.Text = ingredientdata[n, 4];
            this.textBox13.Text = ingredientdata[n, 5];
            this.textBox17.Text = ingredientdata[n, 6];
            this.textBox6.Text = productdata[n, 0];
            this.textBox8.Text = productdata[n, 1];
            this.textBox10.Text = productdata[n, 2];
            this.textBox14.Text = productdata[n, 3];
            this.textBox15.Text = productdata[n, 4];
            this.textBox16.Text = productdata[n, 5];
            this.textBox18.Text = productdata[n, 6];
            this.textBox23.Text = productdata[n, 7];
            this.textBox24.Text = productdata[n, 8];
            this.textBox25.Text = productdata[n, 9];
            this.textBox26.Text = productdata[n, 10];
            this.textBox27.Text = productdata[n, 11];
            this.textBox28.Text = productdata[n, 12];
            this.textBox45.Text = productdata[n, 13];
            // 总下料量计算
            double total = Convert.ToDouble(productdata[n, 0]) + Convert.ToDouble(productdata[n, 1])
                            + Convert.ToDouble(productdata[n, 2]) + Convert.ToDouble(productdata[n, 3])
                            + Convert.ToDouble(productdata[n, 4]) + Convert.ToDouble(productdata[n, 5])
                            + Convert.ToDouble(productdata[n, 6]);
            // 定义原始模块输入成分，用于对应该条下料情况记录
            double oriConCu = 0;
            double oriConFe = 0;
            double oriConS = 0;
            double oriConSiO2 = 0;
            double oriConCaO = 0;
            double oriConMgO = 0;
            double oriConAl2O3 = 0;
            double oriConCo = 0;
            for (int i = 8; i < 64; i += 8)
            {
                oriConCu += Convert.ToDouble(productComponentData[n, i + 0]) * Convert.ToDouble(productdata[n, i / 8 - 1]);
                oriConFe += Convert.ToDouble(productComponentData[n, i + 1]) * Convert.ToDouble(productdata[n, i / 8 - 1]);
                oriConS += Convert.ToDouble(productComponentData[n, i + 2]) * Convert.ToDouble(productdata[n, i / 8 - 1]);
                oriConSiO2 += Convert.ToDouble(productComponentData[n, i + 3]) * Convert.ToDouble(productdata[n, i / 8 - 1]);
                oriConCaO += Convert.ToDouble(productComponentData[n, i + 4]) * Convert.ToDouble(productdata[n, i / 8 - 1]);
                oriConMgO += Convert.ToDouble(productComponentData[n, i + 5]) * Convert.ToDouble(productdata[n, i / 8 - 1]);
                oriConAl2O3 += Convert.ToDouble(productComponentData[n, i + 6]) * Convert.ToDouble(productdata[n, i / 8 - 1]);
                oriConCo += Convert.ToDouble(productComponentData[n, i + 7]) * Convert.ToDouble(productdata[n, i / 8 - 1]);
            }
            // 计算原始模块输入成分
            oriConCu /= total;
            oriConFe /= total;
            oriConS /= total;
            oriConSiO2 /= total;
            oriConCaO /= total;
            oriConMgO /= total;
            oriConAl2O3 /= total;
            oriConCo /= total;
            // 填写原始模块输入成分
            this.textBox29.Text = oriConCu.ToString("0.00");
            this.textBox30.Text = oriConFe.ToString("0.00");
            this.textBox31.Text = oriConS.ToString("0.00");
            this.textBox32.Text = oriConSiO2.ToString("0.00");
            this.textBox33.Text = oriConCaO.ToString("0.00");
            this.textBox34.Text = oriConMgO.ToString("0.00");
            this.textBox35.Text = oriConAl2O3.ToString("0.00");
            this.textBox36.Text = oriConCo.ToString("0.00");
            // 显示原始总下料量
            this.label33.Text = "原配料单总下料量：" + total + " T/h";
            // 记录预期下料量/匹配配料单下料量的比率以做调整
            double rate = totaldosage / total;
            this.textBox38.Text = (Convert.ToDouble(productdata[n, 0]) * rate).ToString("0.00");
            this.textBox39.Text = (Convert.ToDouble(productdata[n, 1]) * rate).ToString("0.00");
            this.textBox40.Text = (Convert.ToDouble(productdata[n, 2]) * rate).ToString("0.00");
            this.textBox41.Text = (Convert.ToDouble(productdata[n, 3]) * rate).ToString("0.00");
            this.textBox42.Text = (Convert.ToDouble(productdata[n, 4]) * rate).ToString("0.00");
            this.textBox43.Text = (Convert.ToDouble(productdata[n, 5]) * rate).ToString("0.00");
            this.textBox44.Text = (Convert.ToDouble(productdata[n, 6]) * rate).ToString("0.00");
            // 2016.9.7新增：每一次更改单号，都必须重填成分信息
            mineraldata = null;
            this.button3.Text = "填写精矿成分";
            // 2016.9.14新增：每一次更改单号，都会删除拟合成分信息
            this.label37.ForeColor = Color.Black;
            this.textBox46.Text = "";
            this.textBox47.Text = "";
            this.textBox48.Text = "";
            this.textBox49.Text = "";
            this.textBox50.Text = "";
            this.textBox51.Text = "";
            this.textBox52.Text = "";
            this.textBox53.Text = "";
            this.textBox58.Text = "";
            this.textBox60.Text = "";
            this.textBox61.Text = "";
            this.textBox62.Text = "";
            this.textBox63.Text = "";
            this.textBox64.Text = "";
            this.textBox65.Text = "";
            this.textBox66.Text = "";
            this.textBox67.Text = "";
            this.textBox68.Text = "";
            this.textBox69.Text = "";
            this.textBox70.Text = "";
            this.label43.ForeColor = Color.Black;
            this.label44.ForeColor = Color.Black;
            this.button6.Enabled = false;
            this.button7.Enabled = false;
            flag = n;
            // 2017.10 新增分配系数自动计算功能
            if ("0".Equals(productdata[n, 16]))
            {
                MessageBox.Show("本配料单未进行分配系数计算，可以进行分配系数计算");
                button8.PerformClick();
            }
            else
            {
                MessageBox.Show("本配料单分配系数已经计算");
            }
        }
        #endregion

        #region 精矿成分的更新和计算
        // 用于数据交换，存储矿物成分信息
        private string[,] mineraldata = null;
        // 用于对应string数组存放实际的double值
        public string[,] Mdata
        {
            get { return this.mineraldata; }
            set { this.mineraldata = value; }
        }
        public string button3Info
        {
            get { return this.button3.Text.ToString(); }
            set { this.button3.Text = value; }
        }
        // 矿名List
        private List<string> mineralList = null;
        private void button3_Click(object sender, EventArgs e)
        {
            isbutton3 = true;//该按钮已点击
            mineralList = new List<string>();
            mineralList.Add(this.textBox5.Text.ToString());
            mineralList.Add(this.textBox7.Text.ToString());
            mineralList.Add(this.textBox9.Text.ToString());
            mineralList.Add(this.textBox11.Text.ToString());
            mineralList.Add(this.textBox12.Text.ToString());
            mineralList.Add(this.textBox13.Text.ToString());
            mineralList.Add(this.textBox17.Text.ToString());
            MineralComponents mc = new MineralComponents(mineralList, this);
            mc.Show();
        }

        // 2016.11.6添加：计算新的混合精矿成分
        public void calcNewComponents()
        {
            // 期望得到的精矿总量
            double total = Convert.ToDouble(this.textBox37.Text.ToString());
            // 存储面板上的修正下料量
            double[] dosage = new double[]
            {
                Convert.ToDouble(Convert.ToDouble(this.textBox38.Text.ToString())),
                Convert.ToDouble(Convert.ToDouble(this.textBox39.Text.ToString())),
                Convert.ToDouble(Convert.ToDouble(this.textBox40.Text.ToString())),
                Convert.ToDouble(Convert.ToDouble(this.textBox41.Text.ToString())),
                Convert.ToDouble(Convert.ToDouble(this.textBox42.Text.ToString())),
                Convert.ToDouble(Convert.ToDouble(this.textBox43.Text.ToString())),
                Convert.ToDouble(Convert.ToDouble(this.textBox44.Text.ToString())),
            };
            // 定义修正后的精矿平均成分
            double fitConCu = 0;
            double fitConFe = 0;
            double fitConS = 0;
            double fitConSiO2 = 0;
            double fitConCaO = 0;
            double fitConMgO = 0;
            double fitConAl2O3 = 0;
            double fitConCo = 0;
            for (int i = 0; i < 7; i++)
            {
                fitConCu += Convert.ToDouble(mineraldata[i, 2]) * dosage[i];
                fitConFe += Convert.ToDouble(mineraldata[i, 3]) * dosage[i];
                fitConS += Convert.ToDouble(mineraldata[i, 4]) * dosage[i];
                fitConSiO2 += Convert.ToDouble(mineraldata[i, 5]) * dosage[i];
                fitConCaO += Convert.ToDouble(mineraldata[i, 6]) * dosage[i];
                fitConMgO += Convert.ToDouble(mineraldata[i, 7]) * dosage[i];
                fitConAl2O3 += Convert.ToDouble(mineraldata[i, 8]) * dosage[i];
                fitConCo += Convert.ToDouble(mineraldata[i, 9]) * dosage[i];
            }
            fitConCu /= total;
            fitConFe /= total;
            fitConS /= total;
            fitConSiO2 /= total;
            fitConCaO /= total;
            fitConMgO /= total;
            fitConAl2O3 /= total;
            fitConCo /= total;
            this.textBox46.Text = fitConCu.ToString("0.00");
            this.textBox47.Text = fitConFe.ToString("0.00");
            this.textBox48.Text = fitConS.ToString("0.00");
            this.textBox49.Text = fitConSiO2.ToString("0.00");
            this.textBox50.Text = fitConCaO.ToString("0.00");
            this.textBox51.Text = fitConMgO.ToString("0.00");
            this.textBox52.Text = fitConAl2O3.ToString("0.00");
            this.textBox53.Text = fitConCo.ToString("0.00");
            this.label37.ForeColor = Color.Red;
        }

        public int selectedIndex
        {
            get { return this.comboBox1.SelectedIndex; }
            set { this.comboBox1.SelectedIndex = value; }
        }
        #endregion

        #region 参数推荐——保存推荐结果
        // 定义全局文件变量，用以和SOMA计算所需表格对应
        string commonFilePath = null;
        private void button5_Click(object sender, EventArgs e)
        {
            if (mineraldata == null)
            {
                MessageBox.Show("请先填写精矿成分，再保存推荐结果！");
                return;
            }
            string[] dosage = { this.textBox38.Text.ToString(), this.textBox39.Text.ToString(),
                                            this.textBox40.Text.ToString(),this.textBox41.Text.ToString(),
                                            this.textBox42.Text.ToString(),this.textBox43.Text.ToString(),
                                             this.textBox44.Text.ToString()};
            // 定义真正的精矿名称和用量（同一种精矿会合并，不添加无矿）
            List<String> realMineralList = new List<string>();
            List<double> realDosage = new List<double>();
            for (int i = 0; i < 7; i++)
            {
                // 集合元素没有添加过，则添加到该集合中，同时添加用量
                if (!realMineralList.Contains(mineralList[i]))
                {
                    // 不添加无矿数据
                    if (!"(无矿)".Equals(mineralList[i]))
                    {
                        realMineralList.Add(mineralList[i]);
                        realDosage.Add(Convert.ToDouble(dosage[i]));
                    }
                }
                // 包含该元素，不添加到集合中，找到集合中已经存在的元素，修改相应的dosage数据
                else
                {
                    for (int j = 0; j < realMineralList.Count; j++)
                    {
                        if (realMineralList[j].Equals(mineralList[i]))
                        {
                            realDosage[j] += Convert.ToDouble(dosage[i]);
                        }
                    }
                }
            }
            // 准备打开工作簿模版文件（同目录）
            string tempPath = "input-output-template.xls";
            HSSFWorkbook wk = null;
            using (FileStream fs = File.Open(tempPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                wk = new HSSFWorkbook(fs);
                fs.Close();
            }
            // 打开第一个工作表
            ISheet tb = wk.GetSheetAt(0);
            // 公式自动重新计算
            tb.ForceFormulaRecalculation = true;
            IRow row = null;
            ICell cell = null;
            // 填充Excel表
            for (int i = 0; i < realMineralList.Count; i++)
            {
                for (int j = 0; j < 7; j++)
                {
                    if (realMineralList[i].Equals(mineralList[j]))
                    {
                        row = tb.GetRow(i + 1);
                        cell = row.GetCell(0);  //获取单元格
                        cell.SetCellValue(realMineralList[i]);//修改数据
                        cell = row.GetCell(1);
                        cell.SetCellValue(Convert.ToDouble(mineraldata[j, 2]));
                        cell = row.GetCell(2);
                        cell.SetCellValue(Convert.ToDouble(mineraldata[j, 3]));
                        cell = row.GetCell(3);
                        cell.SetCellValue(Convert.ToDouble(mineraldata[j, 4]));
                        cell = row.GetCell(4);
                        cell.SetCellValue(Convert.ToDouble(mineraldata[j, 5]));
                        cell = row.GetCell(5);
                        cell.SetCellValue(Convert.ToDouble(mineraldata[j, 6]));
                        cell = row.GetCell(6);
                        cell.SetCellValue(Convert.ToDouble(mineraldata[j, 7]));
                        cell = row.GetCell(7);
                        cell.SetCellValue(Convert.ToDouble(mineraldata[j, 8]));
                        break;
                    }
                }
                cell = row.GetCell(9);
                cell.SetCellValue(realDosage[i]);
            }
            for (int i = realMineralList.Count; i < 7; i++)
            {
                row = tb.GetRow(i + 1);
                cell = row.GetCell(0);  //获取单元格
                cell.SetCellType(CellType.Blank);//修改数据
                cell = row.GetCell(1);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(2);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(3);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(4);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(5);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(6);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(7);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(8);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(9);
                cell.SetCellType(CellType.Blank);
            }
            // 生成文件名
            string[] tmp = this.comboBox1.SelectedItem.ToString().Split('等');
            // 等级左边的数字串作为文件名
            string filename = tmp[0] + ".xls";
            string filepath = "D:\\ExpertSystem\\" + filename;
            commonFilePath = filepath;
            using (FileStream fs = File.Open(filepath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                wk.Write(fs);
                fs.Close();
            }
            editMacroFile(filename);
        }
        #endregion

        /**********************下面开始是物相计算和热力学计算必须的代码**********************/

        #region 全局变量
        // 定义12种精矿
        private Mineral LUANSHYA = null, KANSANSHI = null, LUMWANA = null, CHIBULUMA = null, ENRC = null,
            TF = null, COLD = null, REVERTS = null, LUBAMBE = null, NFCA = null, BOLO = null, CCS = null;
        // 定义石英砂，石灰石，碳
        private Mineral LimeStone = new Mineral();
        private Mineral QuartzSand = new Mineral();
        private Mineral Carbon = new Mineral();
        // 用于存储所有精矿
        private List<Mineral> AllMines = new List<Mineral>();
        // 用于存储当前表格的精矿
        private List<Mineral> Mines = new List<Mineral>();
        // 用于存储当前表格的精矿名称
        private List<String> MinesName = new List<String>();
        #endregion

        #region 参数推荐——物相计算
        // 用于显示计算进度和信息
        CalcProgress calcprogress = null;
        TextBox progressTXT = null;
        ProgressBar progressBAR = null;
        Label progressPERCENT = null;
        // 总共需要计算的精矿数
        private int totalitems = 0;
        // 当前已经计算完成的精矿数
        private int currentitems = 0;
        // 精矿信息合并函数，用于将精矿去重，相同精矿合并下料量
        private void mergeMineralInfo()
        {
            AllMines.Clear();
            MinesName.Clear();
            // 如果已经填写过精矿成分，就直接使用填写的精矿成分进行计算
            if (mineraldata != null)
            {
                // 先对集合元素进行去重工作
                string[] dosage = { this.textBox38.Text.ToString(), this.textBox39.Text.ToString(),
                                            this.textBox40.Text.ToString(),this.textBox41.Text.ToString(),
                                            this.textBox42.Text.ToString(),this.textBox43.Text.ToString(),
                                             this.textBox44.Text.ToString()};
                // 定义真正的精矿名称和用量（同一种精矿会合并，不添加无矿）
                List<String> realMineralList = new List<string>();
                List<double> realDosage = new List<double>();
                for (int i = 0; i < 7; i++)
                {
                    // 集合元素没有添加过，则添加到该集合中，同时添加用量
                    if (!realMineralList.Contains(mineralList[i]))
                    {
                        // 不添加无矿数据
                        if (!"(无矿)".Equals(mineralList[i]))
                        {
                            realMineralList.Add(mineralList[i]);
                            realDosage.Add(Convert.ToDouble(dosage[i]));
                        }
                    }
                    // 包含该元素，不添加到集合中，找到集合中已经存在的元素，修改相应的dosage数据
                    else
                    {
                        for (int j = 0; j < realMineralList.Count; j++)
                        {
                            if (realMineralList[j].Equals(mineralList[i]))
                            {
                                realDosage[j] += Convert.ToDouble(dosage[i]);
                            }
                        }
                    }
                }
                // 开始初始化精矿实体，为后面的物相计算作准备
                for (int j = 0; j < realMineralList.Count; j++)
                {
                    for (int i = 0; i < 7; i++)
                    {
                        // 只有当真实精矿列表和原始精矿列表的名字一样的时候，才能把i对应上
                        if (realMineralList[j].Equals(mineralList[i]))
                        {
                            if ("LUANSHYA".Equals(realMineralList[j]))
                            {
                                LUANSHYA = new Mineral(Double.Parse(mineraldata[i, 2]), Double.Parse(mineraldata[i, 3]),
                                   Double.Parse(mineraldata[i, 4]), Double.Parse(mineraldata[i, 5]), Double.Parse(mineraldata[i, 6]),
                                   Double.Parse(mineraldata[i, 7]), Double.Parse(mineraldata[i, 8]));
                                LUANSHYA.dosage = realDosage[j];
                                AllMines.Add(LUANSHYA);
                                MinesName.Add("LUANSHYA");
                            }
                            if ("KANSANSHI".Equals(realMineralList[j]))
                            {
                                KANSANSHI = new Mineral(Double.Parse(mineraldata[i, 2]), Double.Parse(mineraldata[i, 3]),
                                   Double.Parse(mineraldata[i, 4]), Double.Parse(mineraldata[i, 5]), Double.Parse(mineraldata[i, 6]),
                                   Double.Parse(mineraldata[i, 7]), Double.Parse(mineraldata[i, 8]));
                                KANSANSHI.dosage = realDosage[j];
                                AllMines.Add(KANSANSHI);
                                MinesName.Add("KANSANSHI");
                            }
                            if ("LUMWANA".Equals(realMineralList[j]))
                            {
                                LUMWANA = new Mineral(Double.Parse(mineraldata[i, 2]), Double.Parse(mineraldata[i, 3]),
                                   Double.Parse(mineraldata[i, 4]), Double.Parse(mineraldata[i, 5]), Double.Parse(mineraldata[i, 6]),
                                   Double.Parse(mineraldata[i, 7]), Double.Parse(mineraldata[i, 8]));
                                LUMWANA.dosage = realDosage[j];
                                AllMines.Add(LUMWANA);
                                MinesName.Add("LUMWANA");
                            }
                            if ("CHIBULUMA".Equals(realMineralList[j]))
                            {
                                CHIBULUMA = new Mineral(Double.Parse(mineraldata[i, 2]), Double.Parse(mineraldata[i, 3]),
                                   Double.Parse(mineraldata[i, 4]), Double.Parse(mineraldata[i, 5]), Double.Parse(mineraldata[i, 6]),
                                   Double.Parse(mineraldata[i, 7]), Double.Parse(mineraldata[i, 8]));
                                CHIBULUMA.dosage = realDosage[j];
                                AllMines.Add(CHIBULUMA);
                                MinesName.Add("CHIBULUMA");
                            }
                            if ("ENRC".Equals(realMineralList[j]) || "ENRC矿".Equals(realMineralList[j]))
                            {
                                ENRC = new Mineral(Double.Parse(mineraldata[i, 2]), Double.Parse(mineraldata[i, 3]),
                                   Double.Parse(mineraldata[i, 4]), Double.Parse(mineraldata[i, 5]), Double.Parse(mineraldata[i, 6]),
                                   Double.Parse(mineraldata[i, 7]), Double.Parse(mineraldata[i, 8]));
                                ENRC.dosage = realDosage[j];
                                AllMines.Add(ENRC);
                                MinesName.Add("ENRC");
                            }
                            if ("TF".Equals(realMineralList[j]) || "TF矿".Equals(realMineralList[j]))
                            {
                                TF = new Mineral(Double.Parse(mineraldata[i, 2]), Double.Parse(mineraldata[i, 3]),
                                   Double.Parse(mineraldata[i, 4]), Double.Parse(mineraldata[i, 5]), Double.Parse(mineraldata[i, 6]),
                                   Double.Parse(mineraldata[i, 7]), Double.Parse(mineraldata[i, 8]));
                                TF.dosage = realDosage[j];
                                AllMines.Add(TF);
                                MinesName.Add("TF");
                            }
                            if ("COLD".Equals(realMineralList[j]) || "COLD冷料".Equals(realMineralList[j]))
                            {
                                COLD = new Mineral(Double.Parse(mineraldata[i, 2]), Double.Parse(mineraldata[i, 3]),
                                   Double.Parse(mineraldata[i, 4]), Double.Parse(mineraldata[i, 5]), Double.Parse(mineraldata[i, 6]),
                                   Double.Parse(mineraldata[i, 7]), Double.Parse(mineraldata[i, 8]));
                                COLD.dosage = realDosage[j];
                                AllMines.Add(COLD);
                                MinesName.Add("COLD");
                            }
                            if ("REVERTS".Equals(realMineralList[j]))
                            {
                                REVERTS = new Mineral(Double.Parse(mineraldata[i, 2]), Double.Parse(mineraldata[i, 3]),
                                   Double.Parse(mineraldata[i, 4]), Double.Parse(mineraldata[i, 5]), Double.Parse(mineraldata[i, 6]),
                                   Double.Parse(mineraldata[i, 7]), Double.Parse(mineraldata[i, 8]));
                                REVERTS.dosage = realDosage[j];
                                AllMines.Add(REVERTS);
                                MinesName.Add("REVERTS");
                            }
                            if ("LUBAMBE".Equals(realMineralList[j]))
                            {
                                LUBAMBE = new Mineral(Double.Parse(mineraldata[i, 2]), Double.Parse(mineraldata[i, 3]),
                                   Double.Parse(mineraldata[i, 4]), Double.Parse(mineraldata[i, 5]), Double.Parse(mineraldata[i, 6]),
                                   Double.Parse(mineraldata[i, 7]), Double.Parse(mineraldata[i, 8]));
                                LUBAMBE.dosage = realDosage[j];
                                AllMines.Add(LUBAMBE);
                                MinesName.Add("LUBAMBE");
                            }
                            if ("NFCA".Equals(realMineralList[j]))
                            {
                                NFCA = new Mineral(Double.Parse(mineraldata[i, 2]), Double.Parse(mineraldata[i, 3]),
                                   Double.Parse(mineraldata[i, 4]), Double.Parse(mineraldata[i, 5]), Double.Parse(mineraldata[i, 6]),
                                   Double.Parse(mineraldata[i, 7]), Double.Parse(mineraldata[i, 8]));
                                NFCA.dosage = realDosage[j];
                                AllMines.Add(NFCA);
                                MinesName.Add("NFCA");
                            }
                            if ("BOLO".Equals(realMineralList[j]))
                            {
                                BOLO = new Mineral(Double.Parse(mineraldata[i, 2]), Double.Parse(mineraldata[i, 3]),
                                   Double.Parse(mineraldata[i, 4]), Double.Parse(mineraldata[i, 5]), Double.Parse(mineraldata[i, 6]),
                                   Double.Parse(mineraldata[i, 7]), Double.Parse(mineraldata[i, 8]));
                                BOLO.dosage = realDosage[j];
                                AllMines.Add(BOLO);
                                MinesName.Add("BOLO");
                            }
                            if ("CCS".Equals(realMineralList[j]) || "CCS矿".Equals(realMineralList[j]))
                            {
                                CCS = new Mineral(Double.Parse(mineraldata[i, 2]), Double.Parse(mineraldata[i, 3]),
                                   Double.Parse(mineraldata[i, 4]), Double.Parse(mineraldata[i, 5]), Double.Parse(mineraldata[i, 6]),
                                   Double.Parse(mineraldata[i, 7]), Double.Parse(mineraldata[i, 8]));
                                CCS.dosage = realDosage[j];
                                AllMines.Add(CCS);
                                MinesName.Add("CCS");
                            }
                            break;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("必须先填写精矿成分，才能使用参数推荐功能！");
                return;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            isbutton4 = true;//该按钮已点击
            // 合并精矿信息
            mergeMineralInfo();
            // 判断是否填写精矿成分，否则阻止用户使用参数推荐（2017年1月11日修改）
            if (mineraldata == null)
                return;
            calcprogress = new CalcProgress();
            progressTXT = calcprogress.textBox;
            progressBAR = calcprogress.progressBar;
            progressPERCENT = calcprogress.label;
            // 保存物相状况
            button5_Click(sender, e);
            // 对话框，选择是否立刻进行物相计算
            MessageBoxButtons messButton = MessageBoxButtons.YesNo;
            DialogResult dr = MessageBox.Show("精矿配比数据获取成功！是否立即进行计算？\n（本次计算将花费您较多的时间）", "提示信息", messButton);
            if (dr == DialogResult.Yes)
            {
                // 总计算量设为精矿名称数组的长度，精矿名称数组已经是根据去重的精矿列表筛选出来的
                totalitems = MinesName.Count;
                currentitems = 0;
                calcprogress.Show();
                progressWindowChange("开始进行物相计算...");
                foreach (String s in MinesName)
                {
                    BackgroundWorker bw = new BackgroundWorker();
                    bw.DoWork += new DoWorkEventHandler(bw_DoWork);
                    bw.WorkerSupportsCancellation = true;
                    bw.WorkerReportsProgress = true;
                    bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);
                    bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);
                    bw.RunWorkerAsync(s);
                }
            }
            else
            {
                progressTXT.Text = "";
                return;
            }
        }
        #endregion

        #region 参数推荐——BackgroundWorker相关工作
        // BackGroundWorker的工作执行代码
        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bw = sender as BackgroundWorker;
            string s = e.Argument.ToString();
            if ("LUANSHYA".Equals(s))
            {
                calc_LUANSHYA();
            }
            if ("KANSANSHI".Equals(s))
            {
                calc_KANSANSHI();
            }
            if ("LUMWANA".Equals(s))
            {
                calc_LUMWANA();
            }
            if ("CHIBULUMA".Equals(s))
            {
                calc_CHIBULUMA();
            }
            if ("ENRC".Equals(s))
            {
                calc_ENRC();
            }
            if ("TF".Equals(s))
            {
                calc_TF();
            }
            if ("COLD".Equals(s))
            {
                calc_COLD();
            }
            if ("REVERTS".Equals(s))
            {
                calc_REVERTS();
            }
            if ("LUBAMBE".Equals(s))
            {
                calc_LUBAMBE();
            }
            if ("NFCA".Equals(s))
            {
                calc_NFCA();
            }
            if ("BOLO".Equals(s))
            {
                calc_BOLO();
            }
            if ("CCS".Equals(s))
            {
                calc_CCS();
            }
            bw.ReportProgress(currentitems, s);
        }

        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBarChange(currentitems * 12);
            progressWindowChange("精矿【" + e.UserState.ToString() + "】物相计算完毕，已完成" + currentitems + "项，共" + totalitems + "项");
        }

        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (currentitems == totalitems)
            {
                progressWindowChange("所有精矿物相全部计算完成！");
                ExportProcessToExcel();
                BackgroundWorker bwFactSage = new BackgroundWorker();
                bwFactSage.DoWork += new DoWorkEventHandler(bwFactSage_DoWork);
                bwFactSage.WorkerSupportsCancellation = true;
                bwFactSage.WorkerReportsProgress = true;
                bwFactSage.ProgressChanged += new ProgressChangedEventHandler(bwFactSage_ProgressChanged);
                bwFactSage.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwFactSage_RunWorkerCompleted);
                bwFactSage.RunWorkerAsync();
            }
        }
        private void bwFactSage_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bwFactSage = sender as BackgroundWorker;
            CallFactSageEquilib(bwFactSage);
        }
        private void bwFactSage_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressWindowChange(e.UserState.ToString());
        }
        private void bwFactSage_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressWindowChange("FactSage计算完毕，参数推荐已全部完成！");
            progressBarChange(100);
            calcprogress.startTimer();
            fillResult();
            this.button7.Enabled = true;
        }
        // 创建同步锁
        object locker = new object();
        private void addCurrentItems()
        {
            lock (locker)
            {
                this.currentitems++;
            }
        }
        #endregion

        #region 进度条窗体操控
        // 进度说明文本
        private void progressWindowChange(String str)
        {
            progressTXT.AppendText(str + "\n");
            progressTXT.Refresh();
        }
        // 进度条
        private void progressBarChange(int value)
        {
            progressBAR.Value = value;
            progressBAR.Refresh();
            progressPERCENT.Text = value + "%";
        }
        #endregion

        #region 参数推荐——导出数据到Excel
        private void ExportProcessToExcel()
        {
            // 准备打开工作簿
            string tempPath = commonFilePath;
            HSSFWorkbook wk = null;
            using (FileStream fs = File.Open(tempPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                wk = new HSSFWorkbook(fs);
                fs.Close();
            }
            ISheet tb = wk.GetSheetAt(0);
            tb.ForceFormulaRecalculation = true;
            IRow row = null;
            ICell cell = null;
            // 从第10行开始修改数据
            int i = 10;
            foreach (Mineral m in AllMines)
            {
                row = tb.GetRow(i);
                cell = row.GetCell(0);
                cell.SetCellValue(MinesName[i - 10]);
                cell = row.GetCell(1);
                cell.SetCellValue(m.comp_CuFeS2);
                cell = row.GetCell(2);
                cell.SetCellValue(m.comp_CuS);
                cell = row.GetCell(3);
                cell.SetCellValue(m.comp_Cu2S);
                cell = row.GetCell(4);
                cell.SetCellValue(m.comp_Cu5FeS4);
                cell = row.GetCell(5);
                cell.SetCellValue(m.comp_Cu4SO4_OH_6);
                cell = row.GetCell(6);
                cell.SetCellValue(m.comp_Cu2_OH_2CO3);
                cell = row.GetCell(7);
                cell.SetCellValue(m.comp_FeS2);
                cell = row.GetCell(8);
                cell.SetCellValue(m.comp_Fe2O3);
                cell = row.GetCell(9);
                cell.SetCellValue(m.comp_SiO2);
                cell = row.GetCell(10);
                cell.SetCellValue(m.comp_Mg6Si8O20_OH_4);
                cell = row.GetCell(11);
                cell.SetCellValue(m.comp_KAlSi3O8);
                cell = row.GetCell(12);
                cell.SetCellValue(m.comp_KAl2_AlSi3O10__OH_2);
                cell = row.GetCell(13);
                cell.SetCellValue(m.comp_CaMg_CO3_2);
                cell = row.GetCell(14);
                cell.SetCellValue(m.comp_C);
                cell = row.GetCell(15);
                cell.SetCellValue(m.comp_Cu2O);
                cell = row.GetCell(16);
                cell.SetCellValue(m.comp_Cu);
                cell = row.GetCell(17);
                cell.SetCellValue(m.comp_Fe3O4);
                cell = row.GetCell(18);
                cell.SetCellValue(m.comp_Fe2SiO4);
                cell = row.GetCell(19);
                cell.SetCellValue(m.comp_Fe);
                cell = row.GetCell(20);
                cell.SetCellValue(m.comp_CaO);
                cell = row.GetCell(21);
                cell.SetCellValue(m.comp_Al2O3);
                cell = row.GetCell(22);
                cell.SetCellValue(m.comp_K2O);
                cell = row.GetCell(23);
                cell.SetCellValue(m.comp_MgO);
                cell = row.GetCell(24);
                cell.SetCellValue(m.comp_S2);
                i++;
            }
            for (int j = i; j < 17; j++)
            {
                row = tb.GetRow(j);
                row.GetCell(0).SetCellType(CellType.Blank);
                row.GetCell(1).SetCellType(CellType.Blank);
                row.GetCell(2).SetCellType(CellType.Blank);
                row.GetCell(3).SetCellType(CellType.Blank);
                row.GetCell(4).SetCellType(CellType.Blank);
                row.GetCell(5).SetCellType(CellType.Blank);
                row.GetCell(6).SetCellType(CellType.Blank);
                row.GetCell(7).SetCellType(CellType.Blank);
                row.GetCell(8).SetCellType(CellType.Blank);
                row.GetCell(9).SetCellType(CellType.Blank);
                row.GetCell(10).SetCellType(CellType.Blank);
                row.GetCell(11).SetCellType(CellType.Blank);
                row.GetCell(12).SetCellType(CellType.Blank);
                row.GetCell(13).SetCellType(CellType.Blank);
                row.GetCell(14).SetCellType(CellType.Blank);
                row.GetCell(15).SetCellType(CellType.Blank);
                row.GetCell(16).SetCellType(CellType.Blank);
                row.GetCell(17).SetCellType(CellType.Blank);
                row.GetCell(18).SetCellType(CellType.Blank);
                row.GetCell(19).SetCellType(CellType.Blank);
                row.GetCell(20).SetCellType(CellType.Blank);
                row.GetCell(21).SetCellType(CellType.Blank);
                row.GetCell(22).SetCellType(CellType.Blank);
                row.GetCell(23).SetCellType(CellType.Blank);
                row.GetCell(24).SetCellType(CellType.Blank);
                row.GetCell(25).SetCellType(CellType.Blank);
                row.GetCell(26).SetCellType(CellType.Blank);
            }
            // 填写富氧浓度，下料量，氧气纯度到Excel文档中（2016.11.20）
            cell = tb.GetRow(40).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox55.Text.ToString()));
            cell = tb.GetRow(40).GetCell(4);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox37.Text.ToString()));
            cell = tb.GetRow(40).GetCell(6);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox54.Text.ToString()));

            // 打开第二个工作表，即FactSage可直接使用的工作表
            tb = wk.GetSheetAt(1);
            // 公式自动计算
            tb.ForceFormulaRecalculation = true;
            string p1 = Convert.ToSingle(productdata[flag, 18]) * 100 + "%";
            string p2 = Convert.ToSingle(productdata[flag, 19]) * 100 + "%";
            string p3 = Convert.ToSingle(productdata[flag, 20]) * 100 + "%";
            string p4 = Convert.ToSingle(productdata[flag, 21]) * 100 + "%";
            // 初始化四个分区系数（2017.1.6更新）
            cell = tb.GetRow(13).GetCell(13);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(p1);
            cell = tb.GetRow(13).GetCell(14);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(p2);
            cell = tb.GetRow(13).GetCell(15);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(p3);
            cell = tb.GetRow(13).GetCell(16);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(p4);

            tb = wk.GetSheetAt(3);
            tb.ForceFormulaRecalculation = true;
            // 填写各种原值
            cell = tb.GetRow(32).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox23.Text.ToString()));
            cell = tb.GetRow(33).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox24.Text.ToString()));
            // 第一次计算时氧料比的原值和推荐值相同
            cell = tb.GetRow(33).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox24.Text.ToString()));
            cell = tb.GetRow(34).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox25.Text.ToString()));
            cell = tb.GetRow(35).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox26.Text.ToString()));
            cell = tb.GetRow(41).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox22.Text.ToString()));
            cell = tb.GetRow(42).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox21.Text.ToString()));
            cell = tb.GetRow(43).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox20.Text.ToString()));
            cell = tb.GetRow(44).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox19.Text.ToString()));
            // 填写石英砂，石灰石，煤到Excel文档中
            cell = tb.GetRow(37).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox27.Text.ToString()));
            // 第一次计算时石英砂原值和推荐值相同
            cell = tb.GetRow(37).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox27.Text.ToString()));
            cell = tb.GetRow(38).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox28.Text.ToString()));
            cell = tb.GetRow(39).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox45.Text.ToString()));
            // 填写精矿铜含量、硅含量的原值和推荐值（原混合成分和推荐成分）
            cell = tb.GetRow(47).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox29.Text.ToString()));
            cell = tb.GetRow(48).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox32.Text.ToString()));
            cell = tb.GetRow(47).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox46.Text.ToString()));
            cell = tb.GetRow(48).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox49.Text.ToString()));
            // 填写新Excel表中的四个目标值（2017.1.5更新）
            cell = tb.GetRow(41).GetCell(3);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox1.Text.ToString()));
            cell = tb.GetRow(42).GetCell(3);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox2.Text.ToString()));
            cell = tb.GetRow(43).GetCell(3);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox3.Text.ToString()));
            cell = tb.GetRow(44).GetCell(3);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox4.Text.ToString()));
            // 保存文件
            using (FileStream fs = File.Open(tempPath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                wk.Write(fs);
                fs.Close();
                progressWindowChange("创建精矿物相及辅料情况Excel成功！路径：" + tempPath);
            }
        }
        #endregion

        #region 调用FactSage批处理
        private void CallFactSageEquilib(BackgroundWorker bwFactSage)
        {
            bwFactSage.ReportProgress(1, "准备调用FactSage进行计算...");
            Process process = new Process();
            process.StartInfo.FileName = "CallFactSageEquilib.bat";
            process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            // 参数传递
            process.Start();
            bwFactSage.ReportProgress(1, "正在使用FactSage进行计算，请稍后...");
            // 同步执行 
            process.WaitForExit();
        }
        #endregion

        #region 填充推荐值与预测值
        private void fillResult()
        {
            HSSFWorkbook wk = null;
            using (FileStream fs = File.Open(commonFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                wk = new HSSFWorkbook(fs);
                fs.Close();
            }
            ISheet tb = wk.GetSheetAt(3);
            tb.ForceFormulaRecalculation = true;
            IRow row = null;
            ICell cell = null;
            row = tb.GetRow(32);
            cell = row.GetCell(2);
            this.textBox66.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(33);
            cell = row.GetCell(2);
            this.textBox65.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(34);
            cell = row.GetCell(2);
            this.textBox64.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(35);
            cell = row.GetCell(2);
            this.textBox63.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(36);
            cell = row.GetCell(2);
            this.textBox62.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(37);
            cell = row.GetCell(2);
            this.textBox61.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(38);
            cell = row.GetCell(2);
            this.textBox60.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(39);
            cell = row.GetCell(2);
            this.textBox58.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(41);
            cell = row.GetCell(2);
            this.textBox67.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(42);
            cell = row.GetCell(2);
            this.textBox68.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(43);
            cell = row.GetCell(2);
            this.textBox69.Text = cell.NumericCellValue.ToString("0.00");
            row = tb.GetRow(44);
            cell = row.GetCell(2);
            this.textBox70.Text = cell.NumericCellValue.ToString("0.00");
            this.label43.ForeColor = Color.Red;
            this.label44.ForeColor = Color.Red;
            this.button6.Enabled = true;
            this.predict.Enabled = true;
        }
        #endregion

        #region 参数推荐——熔炼结果预测（直接显示）
        private void button6_Click(object sender, EventArgs e)
        {
            MeltingResult mr = new MeltingResult(commonFilePath);
            mr.Show();
            if (isbutton3 && isbutton4)
                this.Export_file.Enabled = true;
        }
        #endregion

        #region 参数推荐——FactSage迭代计算冰铜和硅铁比（修正结果）
        // 目标冰铜品位
        double targetMatte = 0;
        // 目标硅铁比
        double targetSiO2_Fe = 0;
        double curMatte = 0;
        double curSiO2_Fe = 0;
        double suggestOxyRatio = 0;
        double suggestAddSi = 0;
        double suggestAddCa = 0;
        double suggestAddCoal = 0;
        int ReCalcTimes = 0;
        private void button7_Click(object sender, EventArgs e)
        {
            ReCalcTimes = 0;
            // 获取目标冰铜品位，目标硅铁比，预测冰铜品位，预测硅铁比
            targetMatte = Convert.ToDouble(this.textBox1.Text.ToString());
            targetSiO2_Fe = Convert.ToDouble(this.textBox3.Text.ToString());
            curMatte = Convert.ToDouble(this.textBox67.Text.ToString());
            curSiO2_Fe = Convert.ToDouble(this.textBox69.Text.ToString());
            // 获取当前氧料比和补硅的推荐值
            suggestOxyRatio = Convert.ToDouble(this.textBox65.Text.ToString());
            suggestAddSi = Convert.ToDouble(this.textBox61.Text.ToString());
            // 初始化progress窗体
            calcprogress = new CalcProgress();
            progressTXT = calcprogress.textBox;
            progressBAR = calcprogress.progressBar;
            progressPERCENT = calcprogress.label;
            calcprogress.Show();
            progressWindowChange("===========初始状态===========");
            progressWindowChange("当前冰铜品位：" + curMatte + "，目标值：" + targetMatte);
            progressWindowChange("当前硅铁比：" + curSiO2_Fe + "，目标值：" + targetSiO2_Fe);
            progressWindowChange("当前氧料比：" + suggestOxyRatio);
            progressWindowChange("当前补硅：" + suggestAddSi);
            if (Math.Abs(targetMatte - curMatte) < 0.2)
            {
                // 必须保证冰铜品位符合要求的情况下，才能调整Si
                if (Math.Abs(targetSiO2_Fe - curSiO2_Fe) <= 0.01)
                {
                    MessageBox.Show("预测值已经满足系统精度要求，无需修正！");
                    calcprogress.enabledCloseButton();
                    calcprogress.Close();
                    return;
                }
                else
                {
                    // 预测值偏高，降低补硅
                    if (curSiO2_Fe - targetSiO2_Fe > 0)
                    {
                        suggestAddSi -= Math.Abs(targetSiO2_Fe - curSiO2_Fe) * 10;
                    }
                    else
                    {
                        suggestAddSi += Math.Abs(targetSiO2_Fe - curSiO2_Fe) * 10;
                    }
                }
            }
            else
            {
                // 预测值偏高，降低氧料比
                if (curMatte - targetMatte > 0)
                {
                    suggestOxyRatio -= Math.Abs(targetMatte - curMatte) * 2.5;
                }
                else
                {
                    suggestOxyRatio += Math.Abs(targetMatte - curMatte) * 2.5;
                }
            }
            // 将变换后的氧料比和补硅写入Excel
            ModifyExcelFile(suggestOxyRatio, suggestAddSi);
            progressWindowChange("准备使用FactSage进行迭代计算...");
            // 初次投入FactSage进行计算
            BackgroundWorker bwFactSageReCalc = new BackgroundWorker();
            bwFactSageReCalc.DoWork += new DoWorkEventHandler(bwFactSageReCalc_DoWork);
            bwFactSageReCalc.WorkerSupportsCancellation = true;
            bwFactSageReCalc.WorkerReportsProgress = true;
            bwFactSageReCalc.ProgressChanged += new ProgressChangedEventHandler(bwFactSageReCalc_ProgressChanged);
            bwFactSageReCalc.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwFactSageReCalc_RunWorkerCompleted);
            bwFactSageReCalc.RunWorkerAsync();
            ReCalcTimes++;
        }
        private void bwFactSageReCalc_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bwFactSageReCalc = sender as BackgroundWorker;
            CallFactSageEquilib(bwFactSageReCalc);
        }
        private void bwFactSageReCalc_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }
        private void bwFactSageReCalc_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 重新从Excel中获取数据
            fillResult();
            progressWindowChange("========第" + ReCalcTimes + "次迭代完成========");
            // 重新获取相关的值
            curMatte = Convert.ToDouble(this.textBox67.Text.ToString());
            curSiO2_Fe = Convert.ToDouble(this.textBox69.Text.ToString());
            suggestOxyRatio = Convert.ToDouble(this.textBox65.Text.ToString());
            suggestAddSi = Convert.ToDouble(this.textBox61.Text.ToString());
            progressWindowChange("当前冰铜品位：" + curMatte + "，目标值：" + targetMatte);
            progressWindowChange("当前硅铁比：" + curSiO2_Fe + "，目标值：" + targetSiO2_Fe);
            progressWindowChange("当前氧料比：" + suggestOxyRatio);
            progressWindowChange("当前补硅：" + suggestAddSi);
            if (Math.Abs(targetMatte - curMatte) < 0.2)
            {
                if (Math.Abs(targetSiO2_Fe - curSiO2_Fe) <= 0.01)
                {
                    progressWindowChange("当前预测值已经满足系统精度要求，计算结束！");
                    progressWindowChange("===========计算完成===========");
                    progressBarChange(100);
                    // 解禁关闭窗体按钮，可供用户关闭窗体
                    calcprogress.enabledCloseButton();
                    return;
                }
                else
                {
                    // 预测值偏高，降低补硅
                    if (curSiO2_Fe - targetSiO2_Fe > 0)
                    {
                        suggestAddSi -= Math.Abs(targetSiO2_Fe - curSiO2_Fe) * 10;
                    }
                    else
                    {
                        suggestAddSi += Math.Abs(targetSiO2_Fe - curSiO2_Fe) * 10;
                    }
                }
            }
            else
            {
                // 预测值偏高，降低氧料比
                if (curMatte - targetMatte > 0)
                {
                    suggestOxyRatio -= Math.Abs(targetMatte - curMatte) * 2.5;
                }
                else
                {
                    suggestOxyRatio += Math.Abs(targetMatte - curMatte) * 2.5;
                }
            }
            // 将变换后的氧料比和补硅写入Excel
            ModifyExcelFile(suggestOxyRatio, suggestAddSi);
            ReCalcTimes++;
            progressWindowChange("准备进行第" + ReCalcTimes + "次迭代计算...");
            // 再次投入FactSage进行计算
            BackgroundWorker bwFactSageReCalc = new BackgroundWorker();
            bwFactSageReCalc.DoWork += new DoWorkEventHandler(bwFactSageReCalc_DoWork);
            bwFactSageReCalc.WorkerSupportsCancellation = true;
            bwFactSageReCalc.WorkerReportsProgress = true;
            bwFactSageReCalc.ProgressChanged += new ProgressChangedEventHandler(bwFactSageReCalc_ProgressChanged);
            bwFactSageReCalc.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwFactSageReCalc_RunWorkerCompleted);
            bwFactSageReCalc.RunWorkerAsync();
        }

        // 单独修改Excel文件的氧料比和硅铁比的推荐值，用于迭代计算
        private void ModifyExcelFile(double OxyRatio, double AddSi)
        {
            string tempPath = commonFilePath;
            HSSFWorkbook wk = null;
            using (FileStream fs = File.Open(tempPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                wk = new HSSFWorkbook(fs);
                fs.Close();
            }
            //打开第四个工作表
            ISheet tb = wk.GetSheetAt(3);
            tb.ForceFormulaRecalculation = true;
            ICell cell = null;
            cell = tb.GetRow(33).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(OxyRatio);
            cell = tb.GetRow(37).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(AddSi);
            // 设置前两个表的公式自动计算
            tb = wk.GetSheetAt(0);
            tb.ForceFormulaRecalculation = true;
            tb = wk.GetSheetAt(1);
            tb.ForceFormulaRecalculation = true;
            // 保存文件
            using (FileStream fs = File.Open(tempPath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                wk.Write(fs);
                fs.Close();
            }
        }
        #endregion

        #region 分配系数——一次性导出数据到Excel
        private void ExportAllToExcel()
        {
            if (mineraldata == null)
            {
                MessageBox.Show("对不起，精矿成分获取失败，不能进行分配系数计算:(");
                return;
            }
            string[] dosage = { this.textBox38.Text.ToString(), this.textBox39.Text.ToString(),
                                            this.textBox40.Text.ToString(),this.textBox41.Text.ToString(),
                                            this.textBox42.Text.ToString(),this.textBox43.Text.ToString(),
                                             this.textBox44.Text.ToString()};
            // 定义真正的精矿名称和用量（同一种精矿会合并，不添加无矿）
            List<String> realMineralList = new List<string>();
            List<double> realDosage = new List<double>();
            for (int i = 0; i < 7; i++)
            {
                // 集合元素没有添加过，则添加到该集合中，同时添加用量
                if (!realMineralList.Contains(mineralList[i]))
                {
                    // 不添加无矿数据
                    if (!"(无矿)".Equals(mineralList[i]))
                    {
                        realMineralList.Add(mineralList[i]);
                        realDosage.Add(Convert.ToDouble(dosage[i]));
                    }
                }
                // 包含该元素，不添加到集合中，找到集合中已经存在的元素，修改相应的dosage数据
                else
                {
                    for (int j = 0; j < realMineralList.Count; j++)
                    {
                        if (realMineralList[j].Equals(mineralList[i]))
                        {
                            realDosage[j] += Convert.ToDouble(dosage[i]);
                        }
                    }
                }
            }
            string tempPath = "input-output-template.xls";
            HSSFWorkbook wk = null;
            using (FileStream fs = File.Open(tempPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                wk = new HSSFWorkbook(fs);
                fs.Close();
            }
            ISheet tb = wk.GetSheetAt(0);
            tb.ForceFormulaRecalculation = true;
            IRow row = null;
            ICell cell = null;
            // 填充Excel表
            for (int i = 0; i < realMineralList.Count; i++)
            {
                for (int j = 0; j < 7; j++)
                {
                    if (realMineralList[i].Equals(mineralList[j]))
                    {
                        row = tb.GetRow(i + 1);
                        cell = row.GetCell(0);
                        cell.SetCellValue(realMineralList[i]);
                        cell = row.GetCell(1);
                        cell.SetCellValue(Convert.ToDouble(mineraldata[j, 2]));
                        cell = row.GetCell(2);
                        cell.SetCellValue(Convert.ToDouble(mineraldata[j, 3]));
                        cell = row.GetCell(3);
                        cell.SetCellValue(Convert.ToDouble(mineraldata[j, 4]));
                        cell = row.GetCell(4);
                        cell.SetCellValue(Convert.ToDouble(mineraldata[j, 5]));
                        cell = row.GetCell(5);
                        cell.SetCellValue(Convert.ToDouble(mineraldata[j, 6]));
                        cell = row.GetCell(6);
                        cell.SetCellValue(Convert.ToDouble(mineraldata[j, 7]));
                        cell = row.GetCell(7);
                        cell.SetCellValue(Convert.ToDouble(mineraldata[j, 8]));
                        break;
                    }
                }
                cell = row.GetCell(9);
                cell.SetCellValue(realDosage[i]);
            }
            // 精矿不足7种，则把剩余精矿栏位清空
            for (int i = realMineralList.Count; i < 7; i++)
            {
                row = tb.GetRow(i + 1);
                cell = row.GetCell(0);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(1);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(2);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(3);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(4);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(5);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(6);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(7);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(8);
                cell.SetCellType(CellType.Blank);
                cell = row.GetCell(9);
                cell.SetCellType(CellType.Blank);
            }
            // 从第10行开始修改数据
            int k = 10;
            foreach (Mineral m in AllMines)
            {
                row = tb.GetRow(k);
                cell = row.GetCell(0);
                cell.SetCellValue(MinesName[k - 10]);
                cell = row.GetCell(1);
                cell.SetCellValue(m.comp_CuFeS2);
                cell = row.GetCell(2);
                cell.SetCellValue(m.comp_CuS);
                cell = row.GetCell(3);
                cell.SetCellValue(m.comp_Cu2S);
                cell = row.GetCell(4);
                cell.SetCellValue(m.comp_Cu5FeS4);
                cell = row.GetCell(5);
                cell.SetCellValue(m.comp_Cu4SO4_OH_6);
                cell = row.GetCell(6);
                cell.SetCellValue(m.comp_Cu2_OH_2CO3);
                cell = row.GetCell(7);
                cell.SetCellValue(m.comp_FeS2);
                cell = row.GetCell(8);
                cell.SetCellValue(m.comp_Fe2O3);
                cell = row.GetCell(9);
                cell.SetCellValue(m.comp_SiO2);
                cell = row.GetCell(10);
                cell.SetCellValue(m.comp_Mg6Si8O20_OH_4);
                cell = row.GetCell(11);
                cell.SetCellValue(m.comp_KAlSi3O8);
                cell = row.GetCell(12);
                cell.SetCellValue(m.comp_KAl2_AlSi3O10__OH_2);
                cell = row.GetCell(13);
                cell.SetCellValue(m.comp_CaMg_CO3_2);
                cell = row.GetCell(14);
                cell.SetCellValue(m.comp_C);
                cell = row.GetCell(15);
                cell.SetCellValue(m.comp_Cu2O);
                cell = row.GetCell(16);
                cell.SetCellValue(m.comp_Cu);
                cell = row.GetCell(17);
                cell.SetCellValue(m.comp_Fe3O4);
                cell = row.GetCell(18);
                cell.SetCellValue(m.comp_Fe2SiO4);
                cell = row.GetCell(19);
                cell.SetCellValue(m.comp_Fe);
                cell = row.GetCell(20);
                cell.SetCellValue(m.comp_CaO);
                cell = row.GetCell(21);
                cell.SetCellValue(m.comp_Al2O3);
                cell = row.GetCell(22);
                cell.SetCellValue(m.comp_K2O);
                cell = row.GetCell(23);
                cell.SetCellValue(m.comp_MgO);
                cell = row.GetCell(24);
                cell.SetCellValue(m.comp_S2);
                k++;
            }
            for (int j = k; j < 17; j++)
            {
                row = tb.GetRow(j);
                row.GetCell(0).SetCellType(CellType.Blank);
                row.GetCell(1).SetCellType(CellType.Blank);
                row.GetCell(2).SetCellType(CellType.Blank);
                row.GetCell(3).SetCellType(CellType.Blank);
                row.GetCell(4).SetCellType(CellType.Blank);
                row.GetCell(5).SetCellType(CellType.Blank);
                row.GetCell(6).SetCellType(CellType.Blank);
                row.GetCell(7).SetCellType(CellType.Blank);
                row.GetCell(8).SetCellType(CellType.Blank);
                row.GetCell(9).SetCellType(CellType.Blank);
                row.GetCell(10).SetCellType(CellType.Blank);
                row.GetCell(11).SetCellType(CellType.Blank);
                row.GetCell(12).SetCellType(CellType.Blank);
                row.GetCell(13).SetCellType(CellType.Blank);
                row.GetCell(14).SetCellType(CellType.Blank);
                row.GetCell(15).SetCellType(CellType.Blank);
                row.GetCell(16).SetCellType(CellType.Blank);
                row.GetCell(17).SetCellType(CellType.Blank);
                row.GetCell(18).SetCellType(CellType.Blank);
                row.GetCell(19).SetCellType(CellType.Blank);
                row.GetCell(20).SetCellType(CellType.Blank);
                row.GetCell(21).SetCellType(CellType.Blank);
                row.GetCell(22).SetCellType(CellType.Blank);
                row.GetCell(23).SetCellType(CellType.Blank);
                row.GetCell(24).SetCellType(CellType.Blank);
                row.GetCell(25).SetCellType(CellType.Blank);
                row.GetCell(26).SetCellType(CellType.Blank);
            }
            // 填写富氧浓度，下料量，氧气纯度到Excel文档中（2016.11.20）
            cell = tb.GetRow(40).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox55.Text.ToString()));
            cell = tb.GetRow(40).GetCell(4);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox37.Text.ToString()));
            cell = tb.GetRow(40).GetCell(6);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox54.Text.ToString()));
            // 打开第二个工作表，即FactSage可直接使用的工作表
            tb = wk.GetSheetAt(1);
            // 公式自动计算
            tb.ForceFormulaRecalculation = true;
            // 初始化四个分区系数（2017.1.6更新）
            cell = tb.GetRow(13).GetCell(13);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue("32%");
            cell = tb.GetRow(13).GetCell(14);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue("48%");
            cell = tb.GetRow(13).GetCell(15);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue("20%");
            cell = tb.GetRow(13).GetCell(16);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue("0%");
            // 打开第四个工作表
            tb = wk.GetSheetAt(3);
            // 公式自动计算
            tb.ForceFormulaRecalculation = true;
            // 填写各种原值和推荐值，这两个值完全相同
            cell = tb.GetRow(32).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox23.Text.ToString()));
            cell = tb.GetRow(33).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox24.Text.ToString()));
            cell = tb.GetRow(33).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox24.Text.ToString()));
            cell = tb.GetRow(34).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox25.Text.ToString()));
            cell = tb.GetRow(34).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox25.Text.ToString()));
            cell = tb.GetRow(35).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox26.Text.ToString()));
            // 冰铜品位，Fe3O4，硅铁比，硅钙比是后期迭代需要对比的数据，不能和原值相同
            cell = tb.GetRow(41).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox22.Text.ToString()));
            cell = tb.GetRow(42).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox21.Text.ToString()));
            cell = tb.GetRow(43).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox20.Text.ToString()));
            cell = tb.GetRow(44).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox19.Text.ToString()));
            // 填写石英砂，石灰石，煤到Excel文档中（原值和推荐值相同）
            cell = tb.GetRow(37).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox27.Text.ToString()));
            cell = tb.GetRow(37).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox27.Text.ToString()));
            cell = tb.GetRow(38).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox28.Text.ToString()));
            cell = tb.GetRow(38).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox28.Text.ToString()));
            cell = tb.GetRow(39).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox45.Text.ToString()));
            cell = tb.GetRow(39).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox45.Text.ToString()));
            // 填写精矿铜含量、硅含量的原值和推荐值（原混合成分和推荐成分）
            cell = tb.GetRow(47).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox29.Text.ToString()));
            cell = tb.GetRow(48).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox32.Text.ToString()));
            cell = tb.GetRow(47).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox46.Text.ToString()));
            cell = tb.GetRow(48).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox49.Text.ToString()));
            // 填写新Excel表中的四个目标值（2017.1.5更新）
            cell = tb.GetRow(41).GetCell(3);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox1.Text.ToString()));
            cell = tb.GetRow(42).GetCell(3);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox2.Text.ToString()));
            cell = tb.GetRow(43).GetCell(3);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox3.Text.ToString()));
            cell = tb.GetRow(44).GetCell(3);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(Double.Parse(this.textBox4.Text.ToString()));
            // 生成文件名
            string[] tmp = this.comboBox1.SelectedItem.ToString().Split('等');
            // 等级左边的数字串加上_Partition标识符作为文件名，以和参数推荐的文件进行区分
            string filename = tmp[0] + "_Partition.xls";
            string filepath = "D:\\ExpertSystem\\" + filename;
            // 将filepath存入commonFilePath供后面计算（ExportProcessToExcel）使用
            commonFilePath = filepath;
            using (FileStream fs = File.Open(filepath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                wk.Write(fs);
                fs.Close();
                progressWindowChange("创建分配系数计算Excel成功！路径：" + filepath);
            }
            // 编辑ISA.mac文件中的文件名为这一次推荐所生成的文件名，为后续计算做准备
            editMacroFile(filename);
        }
        #endregion

        #region 分配系数——初次计算
        // 获取mineraldata
        private void getMineralData()
        {
            string[,] tmpMineralData = new string[7, 10];
            // 定义仓号列表
            string[] storenum = new string[] { "1", "2", "3", "4", "8", "9", "10" };
            // 获取原始精矿名称列表
            mineralList = new List<string>();
            mineralList.Add(this.textBox5.Text.ToString());
            mineralList.Add(this.textBox7.Text.ToString());
            mineralList.Add(this.textBox9.Text.ToString());
            mineralList.Add(this.textBox11.Text.ToString());
            mineralList.Add(this.textBox12.Text.ToString());
            mineralList.Add(this.textBox13.Text.ToString());
            mineralList.Add(this.textBox17.Text.ToString());
            // 初始化tmpMineralData
            for (int i = 0; i < 7; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    if (j == 0)
                    {
                        tmpMineralData[i, j] = storenum[i];
                    }
                    else if (j == 1)
                    {
                        tmpMineralData[i, j] = mineralList[i];
                    }
                    else
                    {
                        tmpMineralData[i, j] = productComponentData[selectedIndex, (i + 1) * 8 + j - 2];
                    }
                }
            }
            // mineraldata赋值，完成取精矿成分
            this.mineraldata = tmpMineralData;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            // 获取当前配料单数据
            getMineralData();
            // 计算混合精矿成分
            calcNewComponents();
            // 合并精矿信息（源码在“参数推荐”部分）
            mergeMineralInfo();
            calcprogress = new CalcProgress();
            progressTXT = calcprogress.textBox;
            progressBAR = calcprogress.progressBar;
            progressPERCENT = calcprogress.label;
            // 对话框，选择是否立刻进行物相计算
            MessageBoxButtons messButton = MessageBoxButtons.YesNo;
            DialogResult dr = MessageBox.Show("数据获取成功！是否立即进行分配系数计算？\n（本次计算为迭代计算，将花费您很长时间）", "提示信息", messButton);
            if (dr == DialogResult.Yes)//如果点击“确定”按钮
            {
                // 总计算量设为精矿名称数组的长度，精矿名称数组已经是根据去重的精矿列表筛选出来的
                totalitems = MinesName.Count;
                // 重置currentitems，根据线程计算来增加
                currentitems = 0;
                calcprogress.Show();
                progressWindowChange("开始进行物相计算...");
                foreach (String s in MinesName)
                {
                    BackgroundWorker bwPartition = new BackgroundWorker();
                    bwPartition.DoWork += new DoWorkEventHandler(bwPartition_DoWork);
                    bwPartition.WorkerSupportsCancellation = true;
                    bwPartition.WorkerReportsProgress = true;
                    bwPartition.ProgressChanged += new ProgressChangedEventHandler(bwPartition_ProgressChanged);
                    bwPartition.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwPartition_RunWorkerCompleted);
                    // 参数传递，将当前所指向的MinesName传给BackGroundWorker
                    bwPartition.RunWorkerAsync(s);
                }
            }
            else
            {
                progressTXT.Text = "";
                return;
            }
        }
        #endregion

        #region 分配系数——BackgroundWorker相关工作
        private void bwPartition_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bwPartition = sender as BackgroundWorker;
            string s = e.Argument.ToString();
            if ("LUANSHYA".Equals(s))
            {
                calc_LUANSHYA();
            }
            if ("KANSANSHI".Equals(s))
            {
                calc_KANSANSHI();
            }
            if ("LUMWANA".Equals(s))
            {
                calc_LUMWANA();
            }
            if ("CHIBULUMA".Equals(s))
            {
                calc_CHIBULUMA();
            }
            if ("ENRC".Equals(s))
            {
                calc_ENRC();
            }
            if ("TF".Equals(s))
            {
                calc_TF();
            }
            if ("COLD".Equals(s))
            {
                calc_COLD();
            }
            if ("REVERTS".Equals(s))
            {
                calc_REVERTS();
            }
            if ("LUBAMBE".Equals(s))
            {
                calc_LUBAMBE();
            }
            if ("NFCA".Equals(s))
            {
                calc_NFCA();
            }
            if ("BOLO".Equals(s))
            {
                calc_BOLO();
            }
            if ("CCS".Equals(s))
            {
                calc_CCS();
            }
            bwPartition.ReportProgress(currentitems, s);
        }

        private void bwPartition_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBarChange(currentitems * 12);
            progressWindowChange("精矿【" + e.UserState.ToString() + "】物相计算完毕，已完成" + currentitems + "项，共" + totalitems + "项");
        }

        private void bwPartition_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (currentitems == totalitems)
            {
                progressWindowChange("所有精矿物相全部计算完成！");
                // 计算数据并一次性写入Excel，覆盖sheet4的公式
                ExportAllToExcel();
                BackgroundWorker bwPartitionFactSage = new BackgroundWorker();
                bwPartitionFactSage.DoWork += new DoWorkEventHandler(bwPartitionFactSage_DoWork);
                bwPartitionFactSage.WorkerSupportsCancellation = true;
                bwPartitionFactSage.WorkerReportsProgress = true;
                bwPartitionFactSage.ProgressChanged += new ProgressChangedEventHandler(bwPartitionFactSage_ProgressChanged);
                bwPartitionFactSage.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwPartitionFactSage_RunWorkerCompleted);
                bwPartitionFactSage.RunWorkerAsync();
            }
        }
        private void bwPartitionFactSage_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bwPartitionFactSage = sender as BackgroundWorker;
            CallFactSageEquilib(bwPartitionFactSage);
        }
        private void bwPartitionFactSage_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressWindowChange(e.UserState.ToString());
        }
        private void bwPartitionFactSage_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // FactSage计算完成，刷新进度条到90%
            progressWindowChange("FactSage初次计算完毕，准备进行迭代计算...");
            progressBarChange(90);
            reCalcPartition();
        }
        #endregion

        #region 分配系数——迭代计算
        // 配料单冰铜品位原值
        double originalMatte = 0;
        // 当前迭代计算后的冰铜品位（与参数推荐区分）
        double partitionMatte = 0;
        // 配料单Fe3O4原值
        double originalFe3O4 = 0;
        // 当前迭代计算后的Fe3O4值
        double partitionFe3O4 = 0;
        // 初始分配系数
        double p1 = 0.28, p2 = 0.48, p3 = 0.24, p4 = 0.00;
        // 分配系数迭代计算次数
        int partitionCalcTimes;
        // 分配系数调整步长设定(2018.04从0.01改为0.03和0.03)
        double matteStep = 0.03;
        double Fe3O4Step = 0.03;
        // 设置冰铜调整精确度(2018.04从0.2改为0.5)
        double matteEPS = 0.5;
        // 设置Fe3O4调整精度(2018.04从0.3改为0.4)
        double Fe3O4EPS = 0.4;
        // 分配系数迭代计算函数
        private void reCalcPartition()
        {
            // 重新初始化确保数值准确
            p1 = 0.32;
            p2 = 0.48;
            p3 = 0.20;
            p4 = 0.00;
            // 重置分配系数计算器
            partitionCalcTimes = 0;
            //Debug中，原值为0.2和0.3
            matteEPS = 0.5;
            Fe3O4EPS = 0.4;
            getMatteFe3O4FromExcel();
            progressWindowChange("===========初始状态===========");
            progressWindowChange("当前冰铜品位：" + partitionMatte + "，原值：" + originalMatte);
            progressWindowChange("当前Fe3O4：" + partitionFe3O4 + "，原值：" + originalFe3O4);
            progressWindowChange("当前分配系数：" + p1.ToString("P") + "，" + p2.ToString("P") + "，"
                + p3.ToString("P") + "，" + p4.ToString("P"));
            // 先判断冰铜是否满足
            if (Math.Abs(originalMatte - partitionMatte) < matteEPS)
            {
                // 必须保证冰铜品位符合要求的情况下，才能迭代计算Fe3O4
                if (Math.Abs(originalFe3O4 - partitionFe3O4) <= Fe3O4EPS)
                {
                    progressWindowChange("当前数值与原值差异已经满足系统精度要求，无需迭代计算！");
                    progressWindowChange("===========计算结束===========");
                    progressBarChange(100);
                    //2017.10 新增将计算完毕的分配系数保存到数据库，并更新是否计算的flag

                    //// ---------------MySQL数据库-------------------

                    //string constr = "server=localhost;User Id=ISA;password=123456;Database=ccs";
                    //MySqlConnection mycon = new MySqlConnection(constr);
                    //try
                    //{

                    //    mycon.Open();
                    //    MySqlCommand mycmd = new MySqlCommand("UPDATE production SET p1='" + p1 + "',p2='" + p2 + "',p3='" + p3 + "',p4='" + p4 + "',iscalculation=1 WHERE id=" + productdata[flag, 17], mycon);
                    //    mycmd.ExecuteNonQuery();

                    //    //更新运行程序中的flag
                    //    productdata[flag, 16] = "1";
                    //    //更新程序中存储的分配系数
                    //    productdata[flag, 18] = p1.ToString();
                    //    productdata[flag, 19] = p2.ToString();
                    //    productdata[flag, 20] = p3.ToString();
                    //    productdata[flag, 21] = p4.ToString();
                    //}
                    //catch
                    //{
                    //    MessageBox.Show("数据库出错");
                    //}
                    //finally
                    //{
                    //    mycon.Close();
                    //}

                    //---------------ACCESS数据库------------------
                    string mdbPath = @"ccs.mdb";
                    string strConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdbPath;
                    OleDbConnection odcConnection = new OleDbConnection(strConn);
                    try
                    {

                        odcConnection.Open();
                        OleDbCommand odCommand = odcConnection.CreateCommand();
                        odCommand.CommandText = "UPDATE production SET p1='" + p1 + "',p2='" + p2 + "',p3='" + p3 + "',p4='" + p4 + "',iscalculation=1 WHERE id=" + productdata[flag, 17];
                        odCommand.ExecuteNonQuery();
                        //更新运行程序中的flag
                        productdata[flag, 16] = "1";

                        //更新程序中存储的分配系数
                        productdata[flag, 18] = p1.ToString();
                        productdata[flag, 19] = p2.ToString();
                        productdata[flag, 20] = p3.ToString();
                        productdata[flag, 21] = p4.ToString();
                    }
                    catch
                    {
                        MessageBox.Show("数据库出错");
                    }
                    finally
                    {
                        odcConnection.Close();
                    }

                    MessageBox.Show("分配系数保存成功:" + productdata[flag, 17]);
                    // 解禁关闭窗体按钮
                    calcprogress.enabledCloseButton();
                    return;
                }
                else
                {
                    // Fe3O4调整步长自适应变化
                    Fe3O4Step = Math.Abs(partitionFe3O4 - originalFe3O4) / 100;
                    // 当差值小于2的时候，步长要变得更短；小于1的时候降进行微调
                    if (Fe3O4Step < 0.03) Fe3O4Step /= 2;//（2018.04从Fe3O4Step /= 2 改为Fe3O4Step = 0.02）
                    // Fe3O4计算值偏高，降低N14(p1)，升高P14(p3)，变化量的百分点数值相同
                    if (partitionFe3O4 - originalFe3O4 > 0)
                    {
                        p1 -= Fe3O4Step;
                        p3 += Fe3O4Step;
                    }
                    else
                    {
                        p1 += Fe3O4Step;
                        p3 -= Fe3O4Step;
                    }
                }
            }
            else // 冰铜不满足要求，优先调整冰铜
            {
                // 冰铜调整步长自适应变化，最后稳定在1%
                matteStep = Math.Abs(partitionMatte - originalMatte) / 50;
                if (matteStep < 0.02) matteStep = 0.02;//（2018.04由0.01改为0.02）
                // 冰铜计算值偏高，同比例降低p1,p2,p3，降低量为1%
                if (partitionMatte - originalMatte > 0)
                {
                    p1 *= (1 - matteStep);
                    p2 *= (1 - matteStep);
                    p3 *= (1 - matteStep);
                    p4 = (1 - p1 - p2 - p3) / 2;
                }
                else // 冰铜计算值偏低
                {
                    p1 *= (1 + matteStep);
                    p2 *= (1 + matteStep);
                    p3 *= (1 + matteStep);
                    p4 = (1 - p1 - p2 - p3) / 2;
                    if (p4 < 0)
                    {
                        progressWindowChange("本次分区分配系数调整后，总和将超过100%，计算已强制终止！");
                        progressWindowChange("===========计算结束===========");
                        // 解禁关闭窗体按钮
                        calcprogress.enabledCloseButton();
                        return;
                    }
                }
            }
            writePartitionToExcel(p1, p2, p3, p4);
            progressWindowChange("准备进行第1次迭代计算...");
            // 初次投入FactSage进行计算
            BackgroundWorker bwPartitionReCalc = new BackgroundWorker();
            bwPartitionReCalc.DoWork += new DoWorkEventHandler(bwPartitionReCalc_DoWork);
            bwPartitionReCalc.WorkerSupportsCancellation = true;
            bwPartitionReCalc.WorkerReportsProgress = true;
            bwPartitionReCalc.ProgressChanged += new ProgressChangedEventHandler(bwPartitionReCalc_ProgressChanged);
            bwPartitionReCalc.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwPartitionReCalc_RunWorkerCompleted);
            bwPartitionReCalc.RunWorkerAsync();
            partitionCalcTimes++;
        }
        private void bwPartitionReCalc_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bwPartitionReCalc = sender as BackgroundWorker;
            // 调用FactSage批处理，进行物相计算，sender参数传递过去，可随时汇报进度
            CallFactSageEquilib(bwPartitionReCalc);
        }
        private void bwPartitionReCalc_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }
        bool isFe3O4Complete = false;
        private void bwPartitionReCalc_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressWindowChange("========第" + partitionCalcTimes + "次迭代完成========");
            // 从Excel中重新获取数据
            getMatteFe3O4FromExcel();
            progressWindowChange("当前冰铜品位：" + partitionMatte + "，原值：" + originalMatte);
            progressWindowChange("当前Fe3O4：" + partitionFe3O4 + "，原值：" + originalFe3O4);
            progressWindowChange("当前分配系数：" + p1.ToString("P") + "，" + p2.ToString("P") + "，"
                + p3.ToString("P") + "，" + p4.ToString("P"));
            // 调整冰铜和Fe3O4的精确度
            if (Math.Abs(originalFe3O4 - partitionFe3O4) < 0.4 || isFe3O4Complete)
            {
                // 一旦调整到0.4之内了，就可以永久进入此分支了
                isFe3O4Complete = true;
                matteEPS = 0.5;
            }
            else
            {
                // 冰铜品位调整好了，但是Fe3O4还没有调整好，这时就不管冰铜品位了，允许的误差极大扩增
                if (Math.Abs(originalMatte - partitionMatte) < 0.5) matteEPS = 10;
            }
            // 先判断冰铜是否满足
            if (Math.Abs(originalMatte - partitionMatte) < matteEPS)
            {
                // 必须保证冰铜品位符合要求的情况下，才能迭代计算Fe3O4
                if (Math.Abs(originalFe3O4 - partitionFe3O4) <= Fe3O4EPS)
                {
                    progressWindowChange("当前数值与原值差异已经满足系统精度要求，迭代计算结束！");
                    progressWindowChange("===========计算完成===========");
                    progressBarChange(100);


                    //2017.10 新增将计算完毕的分配系数保存到数据库，并更新是否计算的flag

                    //// ---------------MySQL数据库-------------------

                    //string constr = "server=localhost;User Id=ISA;password=123456;Database=ccs";
                    //MySqlConnection mycon = new MySqlConnection(constr);
                    //try
                    //{

                    //    mycon.Open();
                    //    MySqlCommand mycmd = new MySqlCommand("UPDATE production SET p1='" + p1 + "',p2='" + p2 + "',p3='" + p3 + "',p4='" + p4 + "',iscalculation=1 WHERE id=" + productdata[flag, 17], mycon);
                    //    mycmd.ExecuteNonQuery();

                    //    //更新运行程序中的flag
                    //    productdata[flag, 16] = "1";
                    //    //更新程序中存储的分配系数
                    //    productdata[flag, 18] = p1.ToString();
                    //    productdata[flag, 19] = p2.ToString();
                    //    productdata[flag, 20] = p3.ToString();
                    //    productdata[flag, 21] = p4.ToString();
                    //}
                    //catch
                    //{
                    //    MessageBox.Show("数据库出错");
                    //}
                    //finally
                    //{
                    //    mycon.Close();
                    //}

                    //---------------ACCESS数据库------------------
                    string mdbPath = @"ccs.mdb";
                    string strConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdbPath;
                    OleDbConnection odcConnection = new OleDbConnection(strConn);
                    try
                    {

                        odcConnection.Open();
                        OleDbCommand odCommand = odcConnection.CreateCommand();
                        odCommand.CommandText = "UPDATE production SET p1='" + p1 + "',p2='" + p2 + "',p3='" + p3 + "',p4='" + p4 + "',iscalculation=1 WHERE id=" + productdata[flag, 17];
                        odCommand.ExecuteNonQuery();
                        //更新运行程序中的flag
                        productdata[flag, 16] = "1";

                        //更新程序中存储的分配系数
                        productdata[flag, 18] = p1.ToString();
                        productdata[flag, 19] = p2.ToString();
                        productdata[flag, 20] = p3.ToString();
                        productdata[flag, 21] = p4.ToString();
                    }
                    catch
                    {
                        MessageBox.Show("数据库出错");
                    }
                    finally
                    {
                        odcConnection.Close();
                    }

                    MessageBox.Show("分配系数保存成功，配料单序号为" + productdata[flag, 17]);
                    // 解禁关闭窗体按钮
                    calcprogress.enabledCloseButton();
                    return;
                }
                else
                {
                    // Fe3O4调整步长自适应变化
                    Fe3O4Step = Math.Abs(partitionFe3O4 - originalFe3O4) / 100;
                    // 当差值小于2的时候，步长要变得更短；小于1的时候降进行微调
                    if (Fe3O4Step < 0.03) Fe3O4Step /= 2;
                    // Fe3O4计算值偏高，降低N14(p1)，升高P14(p3)，变化量的百分点数值相同
                    if (partitionFe3O4 - originalFe3O4 > 0)
                    {
                        p1 -= Fe3O4Step;
                        p3 += Fe3O4Step;
                    }
                    else
                    {
                        p1 += Fe3O4Step;
                        p3 -= Fe3O4Step;
                    }
                }
            }
            else // 冰铜不满足要求，优先调整冰铜
            {
                // 冰铜调整步长自适应变化，最后稳定在1%
                matteStep = Math.Abs(partitionMatte - originalMatte) / 50;
                if (matteStep < 0.02) matteStep = 0.02;//（2018.04由0.01改为0.02）
                // 冰铜计算值偏高，同比例降低p1,p2,p3，降低量为1%
                if (partitionMatte - originalMatte > 0)
                {
                    p1 *= (1 - matteStep);
                    p2 *= (1 - matteStep);
                    p3 *= (1 - matteStep);
                    p4 = (1 - p1 - p2 - p3) / 2;
                }
                else // 冰铜计算值偏低
                {
                    p1 *= (1 + matteStep);
                    p2 *= (1 + matteStep);
                    p3 *= (1 + matteStep);
                    p4 = (1 - p1 - p2 - p3) / 2;
                    if (p4 < 0)
                    {
                        progressWindowChange("分区分配系数总和已经超过100%，计算强制终止！");
                        progressWindowChange("===========计算结束===========");
                        return;
                    }
                }
            }
            writePartitionToExcel(p1, p2, p3, p4);
            partitionCalcTimes++;
            progressWindowChange("准备进行第" + partitionCalcTimes + "次迭代计算...");
            // 再次次投入FactSage进行计算
            BackgroundWorker bwPartitionReCalc = new BackgroundWorker();
            bwPartitionReCalc.DoWork += new DoWorkEventHandler(bwPartitionReCalc_DoWork);
            bwPartitionReCalc.WorkerSupportsCancellation = true;
            bwPartitionReCalc.WorkerReportsProgress = true;
            bwPartitionReCalc.ProgressChanged += new ProgressChangedEventHandler(bwPartitionReCalc_ProgressChanged);
            bwPartitionReCalc.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwPartitionReCalc_RunWorkerCompleted);
            bwPartitionReCalc.RunWorkerAsync();
        }

        // 将分配系数信息写入Excel表格
        private void writePartitionToExcel(double part1, double part3, double part6, double part5)
        {
            // 准备打开工作簿
            string tempPath = commonFilePath;
            HSSFWorkbook wk = null;
            using (FileStream fs = File.Open(tempPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //把xls文件读入workbook变量里，之后就可以关闭了  
                wk = new HSSFWorkbook(fs);
                fs.Close();
            }
            //打开第二个工作表
            ISheet tb = wk.GetSheetAt(1);
            // 公式自动重新计算
            tb.ForceFormulaRecalculation = true;
            ICell cell = null;
            cell = tb.GetRow(13).GetCell(13);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(part1.ToString("P"));
            cell = tb.GetRow(13).GetCell(14);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(part3.ToString("P"));
            cell = tb.GetRow(13).GetCell(15);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(part6.ToString("P"));
            cell = tb.GetRow(13).GetCell(16);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(part5.ToString("P"));
            // 设置第一个表和第四个表的公式也自动计算
            tb = wk.GetSheetAt(0);
            tb.ForceFormulaRecalculation = true;
            tb = wk.GetSheetAt(3);
            tb.ForceFormulaRecalculation = true;
            // 保存文件
            using (FileStream fs = File.Open(tempPath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                wk.Write(fs);
                fs.Close();
            }
        }
        // 从Excel表格中获取冰铜品位和Fe3O4的信息
        private void getMatteFe3O4FromExcel()
        {
            string tempPath = commonFilePath;
            HSSFWorkbook wk = null;
            using (FileStream fs = File.Open(tempPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                wk = new HSSFWorkbook(fs);
                fs.Close();
            }
            ISheet tb = wk.GetSheetAt(3);
            tb.ForceFormulaRecalculation = true;
            ICell cell = null;
            cell = tb.GetRow(41).GetCell(1);
            originalMatte = cell.NumericCellValue;
            cell = tb.GetRow(41).GetCell(2);
            partitionMatte = cell.NumericCellValue;
            cell = tb.GetRow(42).GetCell(1);
            originalFe3O4 = cell.NumericCellValue;
            cell = tb.GetRow(42).GetCell(2);
            partitionFe3O4 = cell.NumericCellValue;
        }
        #endregion

        #region 12种精矿的化合物组成计算方法
        private void calc_LUANSHYA()
        {
            // 局部变量定义
            double[] mink = new double[] { 47, 0, 1, 8, 0 };
            double[] maxk = new double[] { 64, 4, 7, 18, 4 };
            int D = 5;
            double[,] x = new double[,] { { 0.347826087,0.666666667,0.634920635,0,0 },
                                         { 0.304347826,0,0.111111111,0.466666667,0.7 },
                                         { 0.347826087,0.333333333,0.253968254,0.533333333,0 } };
            double[] y = new double[] { LUANSHYA.Cu - 1.106666667, LUANSHYA.Fe - 0, LUANSHYA.S - 0.16 };
            // SOMA执行
            SOMA soma = new SOMA(D, mink, maxk, x, y);
            Result_K rk = soma.startMigrate();
            double[] k = rk.K;
            // 初始化该种矿石的化合元素成分
            LUANSHYA.comp_CuFeS2 = k[0];
            LUANSHYA.comp_CuS = k[1];
            LUANSHYA.comp_Cu5FeS4 = k[2];
            LUANSHYA.comp_FeS2 = k[3];
            LUANSHYA.comp_Fe2O3 = k[4];
            LUANSHYA.comp_Cu2S = 0.800;
            LUANSHYA.comp_Cu2O = 0.300;
            LUANSHYA.comp_Cu = 0.200;
            LUANSHYA.comp_CaO = LUANSHYA.CaO;
            // 需要计算的部分
            LUANSHYA.comp_Mg6Si8O20_OH_4 = LUANSHYA.MgO * 3.15;
            LUANSHYA.comp_KAlSi3O8 = LUANSHYA.Al2O3 * 5.451;
            LUANSHYA.comp_SiO2 = LUANSHYA.SiO2 - LUANSHYA.comp_KAlSi3O8 * 0.6475
                - LUANSHYA.comp_Mg6Si8O20_OH_4 * 0.635;
            // 计算实际值和模拟值的差异，调整混合成分（2016.10.20）
            // 模拟值-实际值
            double DiffCu = soma.yfit[0] - soma.yreal[0];
            // 若Cu偏高，则降CuFeS2（降的量为差值*184/64），若偏低，则增Cu物相的量。
            if (DiffCu > 0)
            {
                LUANSHYA.comp_CuFeS2 -= DiffCu * 184 / 64;
                k[0] = LUANSHYA.comp_CuFeS2;
            }
            else
                LUANSHYA.comp_Cu -= DiffCu;// 偏低时，DiffCu为负数
            // 重新计算S和Fe
            soma.yfit[2] = soma.calcFit(k, x, 2);
            soma.yfit[1] = soma.calcFit(k, x, 1);
            double DiffS = soma.yfit[2] - soma.yreal[2];
            // 若S偏高，则降FeS2的量（降的量为差值*120/64），若S偏低，则增加S2（增量为差值）
            if (DiffS > 0)
            {
                // 自己理解：降低了FeS2，将会同步降低Fe，降的量为上述差值*120/56
                if (LUANSHYA.comp_FeS2 > DiffS * 120 / 64)
                {
                    // 大于的话才能减，小于等于置为0
                    LUANSHYA.comp_FeS2 -= DiffS * 120 / 64;
                    // 同步降低Fe
                    soma.yfit[1] -= DiffS * 56 / 64;
                }
                else
                {
                    LUANSHYA.comp_FeS2 = 0;
                    // FeS2比差值要小，那么就把FeS2变成0为止，此时的差值便为FeS2的量（可能产生误差）
                    soma.yfit[1] -= LUANSHYA.comp_FeS2 * 56 / 64;
                }
            }
            else
                LUANSHYA.comp_S2 -= DiffS;// 偏低时，DiffS为负数
            double DiffFe = soma.yfit[1] - soma.yreal[1];
            // 若Fe偏高，则降Fe2O3（降的量为差值*160/112），若Fe偏低，则增加Fe2O3（增的量为差值*160/112）
            LUANSHYA.comp_Fe2O3 -= DiffFe * 160 / 112;// 不论偏高还是偏低，DiffFe的符号都会自动变化
            if (LUANSHYA.comp_Fe2O3 < 0) LUANSHYA.comp_Fe2O3 = 0;
            Mines.Add(LUANSHYA);
            // 改变当前完成精矿计数（2016.11.24）
            addCurrentItems();
        }
        private void calc_KANSANSHI()
        {
            // 局部变量定义
            double[] mink = new double[] { 45, 0, 0, 23 };
            double[] maxk = new double[] { 60.4, 6, 9, 36 };
            int D = 4;
            double[,] x = new double[,] { { 0.347826087,0.666666667,0.8,0 },
                                          { 0.304347826,0,0,0.466666667 },
                                          { 0.347826087,0.333333333,0.2,0.533333333 } };
            double[] y = new double[] { KANSANSHI.Cu - 0, KANSANSHI.Fe - 0, KANSANSHI.S - 0 };
            // SOMA执行
            SOMA soma = new SOMA(D, mink, maxk, x, y);
            Result_K rk = soma.startMigrate();
            double[] k = rk.K;
            // 初始化该种矿石的化合元素成分
            KANSANSHI.comp_CuFeS2 = k[0];
            KANSANSHI.comp_CuS = k[1];
            KANSANSHI.comp_Cu2S = k[2];
            KANSANSHI.comp_FeS2 = k[3];
            KANSANSHI.comp_K2O = 0.500;
            KANSANSHI.comp_SiO2 = KANSANSHI.SiO2;
            KANSANSHI.comp_Al2O3 = KANSANSHI.Al2O3;
            KANSANSHI.comp_CaO = KANSANSHI.CaO;
            KANSANSHI.comp_MgO = KANSANSHI.MgO;
            // 计算实际值和模拟值的差异，调整混合成分（2016.10.20）
            // 模拟值-实际值
            double DiffCu = soma.yfit[0] - soma.yreal[0];
            // 若Cu偏高，则降CuFeS2（降的量为差值*184/64），若偏低，则增Cu物相的量。
            if (DiffCu > 0)
            {
                KANSANSHI.comp_CuFeS2 -= DiffCu * 184 / 64;
                k[0] = KANSANSHI.comp_CuFeS2;
            }
            else
                KANSANSHI.comp_Cu -= DiffCu;// 偏低时，DiffCu为负数
            // 重新计算S，Fe的拟合情况
            soma.yfit[2] = soma.calcFit(k, x, 2);
            soma.yfit[1] = soma.calcFit(k, x, 1);
            // 计算S的差值
            double DiffS = soma.yfit[2] - soma.yreal[2];
            // 若S偏高，则降FeS2的量（降的量为差值*120/64），若S偏低，则增加S2（增量为差值）
            if (DiffS > 0)
            {
                // 自己理解：降低了FeS2，将会同步降低Fe，降的量为上述差值*120/56
                if (KANSANSHI.comp_FeS2 > DiffS * 120 / 64)
                {
                    // 大于的话才能减，小于等于置为0
                    KANSANSHI.comp_FeS2 -= DiffS * 120 / 64;
                    soma.yfit[1] -= DiffS * 56 / 64;
                }
                else
                {
                    KANSANSHI.comp_FeS2 = 0;
                    // FeS2比差值要小，那么就把FeS2变成0为止，此时的差值便为FeS2的量
                    soma.yfit[1] -= KANSANSHI.comp_FeS2 * 56 / 64;
                }
            }
            else
            {
                KANSANSHI.comp_S2 -= DiffS;// 偏低时，DiffS为负数
                // MessageBox.Show("【S调整】KANSANSHI的S2已经增加：" + (-1) * DiffS);
            }
            // 最后修正Fe（2016.11.4）
            double DiffFe = soma.yfit[1] - soma.yreal[1];
            // 若Fe偏高，则降FeS2的量（降的量为差值*120/56），若Fe偏低，则增加Fe物相的量。
            // 此时S偏低，则增加S2（增量为差值）
            if (DiffFe < 0)
                KANSANSHI.comp_Fe -= DiffFe;
            else
            {
                KANSANSHI.comp_FeS2 -= DiffFe * 120 / 56;
                // 不论S2是否为0，前面都已经调好了S2，这次降低FeS2必然使S偏低，因此直接加上即可。
                KANSANSHI.comp_S2 += DiffFe * 64 / 56;
            }
            Mines.Add(KANSANSHI);
            // 改变当前完成精矿计数（2016.11.24）
            addCurrentItems();
        }

        private void calc_LUMWANA()// 2016.11.3已更新算法
        {
            // 局部变量定义
            double[] mink = new double[] { 49.7, 0, 0, 15.2 };
            double[] maxk = new double[] { 60, 6, 4, 23.2 };
            int D = 4;
            double[,] x = new double[,] { { 0.347826087,0.666666667,0.8,0.634920635 },
                                          { 0.304347826,0,0,0.111111111 },
                                          { 0.347826087,0.333333333,0.2,0.253968254 } };
            double[] y = new double[] { LUMWANA.Cu - 0, LUMWANA.Fe - 0.35, LUMWANA.S - 0 };
            // SOMA执行
            SOMA soma = new SOMA(D, mink, maxk, x, y);
            Result_K rk = soma.startMigrate();
            double[] k = rk.K;
            LUMWANA.comp_CuFeS2 = k[0];
            LUMWANA.comp_CuS = k[1];
            LUMWANA.comp_Cu2S = k[2];
            LUMWANA.comp_Cu5FeS4 = k[3];
            LUMWANA.comp_KAl2_AlSi3O10__OH_2 = LUMWANA.Al2O3 * 2.601;
            LUMWANA.comp_Fe2O3 = 0.5;
            LUMWANA.comp_Mg6Si8O20_OH_4 = LUMWANA.MgO * 3.15;
            LUMWANA.comp_SiO2 = LUMWANA.SiO2 - (LUMWANA.comp_KAl2_AlSi3O10__OH_2 * 0.4523
                + LUMWANA.comp_Mg6Si8O20_OH_4 * 0.635);
            // 防止负数判定
            if (LUMWANA.comp_SiO2 < 0) LUMWANA.comp_SiO2 = 0;
            LUMWANA.comp_C = 0.000;
            LUMWANA.comp_CaO = LUMWANA.CaO;
            // 计算实际值和模拟值的差异，调整混合成分（2016.10.20）
            // 模拟值-实际值
            double DiffCu = soma.yfit[0] - soma.yreal[0];
            // 若Cu偏高，则降Cu5FeS4（降的量为差值*504/320），若偏低，则增Cu物相的量。
            if (DiffCu > 0)
            {
                LUMWANA.comp_Cu5FeS4 -= DiffCu * 504 / 320;
                // 修正拟合出来的k[3]，以便重新计算S/Fe的总量是否正确（2016.11.4）
                k[3] = LUMWANA.comp_Cu5FeS4;
            }
            else
                LUMWANA.comp_Cu -= DiffCu;// 偏低时，DiffCu为负数
            // 第二个计算的是S，严格按照顺序来（2016.11.4）
            // 重新计算S的拟合情况
            soma.yfit[2] = soma.calcFit(k, x, 2);
            // 由于Fe也发生了改变，Fe的拟合结果也需要重算
            soma.yfit[1] = soma.calcFit(k, x, 1);
            // 计算差值
            double DiffS = soma.yfit[2] - soma.yreal[2];
            // 若S偏高，则降FeS2的量（降的量为差值*120/64），若S偏低，则增加S2（增量为差值）
            if (DiffS > 0)
            {
                // 自己理解：降低了FeS2，将会同步降低Fe，降的量为上述差值*120/56
                if (LUMWANA.comp_FeS2 > DiffS * 120 / 64)
                {
                    // 大于的话才能减，小于等于置为0
                    LUMWANA.comp_FeS2 -= DiffS * 120 / 64;
                    // 同步降低Fe的量
                    soma.yfit[1] -= DiffS * 56 / 64;
                }
                else
                {
                    LUMWANA.comp_FeS2 = 0;
                    // FeS2比差值要小，那么就把FeS2变成0为止，此时的差值便为FeS2的量
                    soma.yfit[1] -= LUMWANA.comp_FeS2 * 56 / 64;
                }
            }
            else
                LUMWANA.comp_S2 -= DiffS;// 偏低时，DiffS为负数
            double DiffFe = soma.yfit[1] - soma.yreal[1];
            // 若Fe偏高，则降Fe2O3（降的量为差值*160/112），若Fe偏低，则增加Fe2O3（增的量为差值*160/112）。
            LUMWANA.comp_Fe2O3 -= DiffFe * 160 / 112;// 不论偏高还是偏低，DiffFe的符号都会自动变化
            // 防止出现负数
            if (LUMWANA.comp_Fe2O3 < 0) LUMWANA.comp_Fe2O3 = 0;
            Mines.Add(LUMWANA);
            // 改变当前完成精矿计数（2016.11.24）
            addCurrentItems();
        }

        private void calc_CHIBULUMA()
        {
            // 局部变量定义
            double[] mink = new double[] { 0, 2, 49, 0, 2 };
            double[] maxk = new double[] { 5, 9, 59, 4, 7 };
            int D = 5;
            double[,] x = new double[,] { { 0.347826087,0.8,0.634920635,0,0 },
                                          { 0.304347826,0,0.111111111,0.466666667,0.7 },
                                          { 0.347826087,0.2,0.253968254,0.533333333,0 } };
            double[] y = new double[] { CHIBULUMA.Cu - 1.628698258, CHIBULUMA.Fe - 0, CHIBULUMA.S - 0.553891336 };
            // SOMA执行
            SOMA soma = new SOMA(D, mink, maxk, x, y);
            Result_K rk = soma.startMigrate();
            double[] k = rk.K;
            CHIBULUMA.comp_CuFeS2 = k[0];
            CHIBULUMA.comp_Cu2S = k[1];
            CHIBULUMA.comp_Cu5FeS4 = k[2];
            CHIBULUMA.comp_FeS2 = k[3];
            CHIBULUMA.comp_Fe2O3 = k[4];
            CHIBULUMA.comp_CuS = 1.52;
            CHIBULUMA.comp_Cu4SO4_OH_6 = 0.67;
            CHIBULUMA.comp_Cu2_OH_2CO3 = 0.36;
            CHIBULUMA.comp_SiO2 = 5.1;
            CHIBULUMA.comp_CaMg_CO3_2 = 1.85;
            CHIBULUMA.comp_Mg6Si8O20_OH_4 = (CHIBULUMA.SiO2 - 6.8) * 1.575;
            CHIBULUMA.comp_KAlSi3O8 = 1.57;
            CHIBULUMA.comp_KAl2_AlSi3O10__OH_2 = 1.51;
            CHIBULUMA.comp_Al2O3 = CHIBULUMA.Al2O3 - 0.8685;
            CHIBULUMA.comp_Cu = 0.03;
            CHIBULUMA.comp_CaO = CHIBULUMA.CaO - 0.563;
            CHIBULUMA.comp_K2O = 0.5;
            // 防负数判定
            if (CHIBULUMA.comp_Mg6Si8O20_OH_4 < 0) CHIBULUMA.comp_Mg6Si8O20_OH_4 = 0;
            if (CHIBULUMA.comp_Al2O3 < 0) CHIBULUMA.comp_Al2O3 = 0;
            if (CHIBULUMA.comp_CaO < 0) CHIBULUMA.comp_CaO = 0;
            // 计算实际值和模拟值的差异，调整混合成分（2016.10.20）
            // 模拟值-实际值
            double DiffCu = soma.yfit[0] - soma.yreal[0];
            // 若Cu偏高，则降Cu5FeS4（降的量为差值*504/320），若偏低，则增Cu物相的量。
            if (DiffCu > 0)
            {
                CHIBULUMA.comp_Cu5FeS4 -= DiffCu * 504 / 320;
                k[2] = CHIBULUMA.comp_Cu5FeS4;
            }
            else
                CHIBULUMA.comp_Cu -= DiffCu;// 偏低时，DiffCu为负数
            // 重新计算S和Fe
            soma.yfit[2] = soma.calcFit(k, x, 2);
            soma.yfit[1] = soma.calcFit(k, x, 1);
            double DiffS = soma.yfit[2] - soma.yreal[2];
            // 若S偏高，则降FeS2的量（降的量为差值*120/64），若S偏低，则增加S2（增量为差值）
            if (DiffS > 0)
            {
                // 自己理解：降低了FeS2，将会同步降低Fe，降的量为上述差值*120/56
                if (CHIBULUMA.comp_FeS2 > DiffS * 120 / 64)
                {
                    // 大于的话才能减，小于等于置为0
                    CHIBULUMA.comp_FeS2 -= DiffS * 120 / 64;
                    // 同步降低Fe
                    soma.yfit[1] -= DiffS * 56 / 64;
                }
                else
                {
                    CHIBULUMA.comp_FeS2 = 0;
                    // FeS2比差值要小，那么就把FeS2变成0为止，此时的差值便为FeS2的量
                    soma.yfit[1] -= CHIBULUMA.comp_FeS2 * 56 / 64;
                }
            }
            else
                CHIBULUMA.comp_S2 -= DiffS;// 偏低时，DiffS为负数
            double DiffFe = soma.yfit[1] - soma.yreal[1];
            // 若Fe偏高，则降Fe2O3（降的量为差值*160/112），若Fe偏低，则增加Fe2O3（增的量为差值*160/112）。
            CHIBULUMA.comp_Fe2O3 -= DiffFe * 160 / 112;// 不论偏高还是偏低，DiffFe的符号都会自动变化
            // 防止出现负数
            if (CHIBULUMA.comp_Fe2O3 < 0) CHIBULUMA.comp_Fe2O3 = 0;
            Mines.Add(CHIBULUMA);
            // 改变当前完成精矿计数（2016.11.24）
            addCurrentItems();
        }
        // 11月20日修复此精矿Cu计算不准确的bug
        private void calc_ENRC()
        {
            // 局部变量定义
            double[] mink = new double[] { 58, 0, 0, 10 };
            double[] maxk = new double[] { 70, 7, 4, 25 };
            int D = 4;
            double[,] x = new double[,] { { 0.347826087,0.666666667,0.8,0 },
                                          { 0.304347826,0,0,0.466666667 },
                                          { 0.347826087,0.333333333,0.2,0.533333333 } };
            double[] y = new double[] { ENRC.Cu - 0, ENRC.Fe - 0, ENRC.S - 0 };
            // SOMA执行
            SOMA soma = new SOMA(D, mink, maxk, x, y);
            Result_K rk = soma.startMigrate();
            double[] k = rk.K;
            ENRC.comp_CuFeS2 = k[0];
            ENRC.comp_CuS = k[1];
            ENRC.comp_Cu2S = k[2];
            ENRC.comp_FeS2 = k[3];
            ENRC.comp_KAlSi3O8 = ENRC.Al2O3 * 5.451;
            ENRC.comp_SiO2 = ENRC.SiO2 - ENRC.comp_KAlSi3O8 * 0.6475;
            if (ENRC.comp_SiO2 < 0) ENRC.comp_SiO2 = 0;
            ENRC.comp_CaO = ENRC.CaO;
            ENRC.comp_MgO = ENRC.MgO;
            // 计算实际值和模拟值的差异，调整混合成分（2016.10.20）
            // 模拟值-实际值
            double DiffCu = soma.yfit[0] - soma.yreal[0];
            // 若Cu偏高，则降CuFeS2（降的量为差值*184/64），若偏低，则增Cu物相的量。
            if (DiffCu > 0)
            {
                ENRC.comp_CuFeS2 -= DiffCu * 184 / 64;
                k[0] = ENRC.comp_CuFeS2;
            }
            else
                ENRC.comp_Cu -= DiffCu;// 偏低时，DiffCu为负数
            // 重新计算S和Fe
            soma.yfit[2] = soma.calcFit(k, x, 2);
            soma.yfit[1] = soma.calcFit(k, x, 1);
            double DiffS = soma.yfit[2] - soma.yreal[2];
            // 若S偏高，则降FeS2的量（降的量为差值*120/64），若S偏低，则增加S2（增量为差值）
            if (DiffS > 0)
            {
                // 自己理解：降低了FeS2，将会同步降低Fe，降的量为上述差值*120/56
                if (ENRC.comp_FeS2 > DiffS * 120 / 64)
                {
                    // 大于的话才能减，小于等于置为0
                    ENRC.comp_FeS2 -= DiffS * 120 / 64;
                    // 同步降低Fe
                    soma.yfit[1] -= DiffS * 56 / 64;
                }
                else
                {
                    ENRC.comp_FeS2 = 0;
                    // FeS2比差值要小，那么就把FeS2变成0为止，此时的差值便为FeS2的量
                    soma.yfit[1] -= ENRC.comp_FeS2 * 56 / 64;
                }
            }
            else
                ENRC.comp_S2 -= DiffS;// 偏低时，DiffS为负数
            double DiffFe = soma.yfit[1] - soma.yreal[1];
            // 若Fe偏高，则降FeS2的量（降的量为差值 * 120 / 64），若Fe偏低，则增加Fe物相的量。
            // 此时S偏低，则增加S2（增量为差值）
            if (DiffFe < 0)
                ENRC.comp_Fe -= DiffFe;// 偏低时，DiffFe为负数
            else
            {
                ENRC.comp_FeS2 -= DiffFe * 120 / 56;
                // 不论S2是否为0，前面都已经调好了S2，这次降低FeS2必然使S偏低，因此直接加上差值即可。
                ENRC.comp_S2 += DiffFe * 64 / 56;
            }
            Mines.Add(ENRC);
            // 改变当前完成精矿计数（2016.11.24）
            addCurrentItems();
        }
        // 2016.11.20更新TF矿的最新计算方法
        private void calc_TF()
        {
            // 局部变量定义
            double[] mink = new double[] { 40, 0, 0, 0, 0 };
            double[] maxk = new double[] { 51, 6, 5, 8, 6 };
            int D = 5;
            double[,] x = new double[,] { { 0.347826087,0.666666667,0.8,0,0 },
                                          { 0.304347826,0,0,0.466666667,0.7 },
                                          { 0.347826087,0.333333333,0.2,0.533333333,0 } };
            double[] y = new double[] { TF.Cu - 6.3, TF.Fe - 0, TF.S - 0 };
            // SOMA执行
            SOMA soma = new SOMA(D, mink, maxk, x, y);
            Result_K rk = soma.startMigrate();
            double[] k = rk.K;
            TF.comp_CuFeS2 = k[0];
            TF.comp_CuS = k[1];
            TF.comp_Cu2S = k[2];
            TF.comp_FeS2 = k[3];
            TF.comp_Fe2O3 = k[4];
            TF.comp_Al2O3 = TF.Al2O3 - 0.9;
            TF.comp_Mg6Si8O20_OH_4 = TF.MgO * 3.15;
            TF.comp_SiO2 = TF.SiO2 - 3.17 - TF.comp_Mg6Si8O20_OH_4 * 0.635;
            TF.comp_CaO = TF.CaO;
            TF.comp_KAlSi3O8 = 4.900;
            TF.comp_Cu = 6.300;
            // 防负数判定
            if (TF.comp_SiO2 < 0) TF.comp_SiO2 = 0;
            if (TF.comp_Al2O3 < 0) TF.comp_Al2O3 = 0;
            // 计算实际值和模拟值的差异，调整混合成分（2016.11.20）
            // 模拟值-实际值
            double DiffCu = soma.yfit[0] - soma.yreal[0];
            // 若Cu偏高，则降Cu（降的量为差值），若偏低，则增Cu物相的量。
            TF.comp_Cu -= DiffCu;// 不论偏高还是偏低，符号都将根据DiffCu的值自动改变
            double DiffS = soma.yfit[2] - soma.yreal[2];
            // 若S偏高，则降FeS2的量（降的量为差值 * 120 / 64），若S偏低，则增加S2（增量为差值）
            if (DiffS > 0)
            {
                // 降低了FeS2，将会同步降低Fe，降的量为上述差值*56/64
                if (TF.comp_FeS2 > DiffS * 120 / 64)
                {
                    // 大于的话才能减，小于等于置为0
                    TF.comp_FeS2 -= DiffS * 120 / 64;
                    // 同步降低Fe
                    soma.yfit[1] -= DiffS * 56 / 64;
                }
                else
                {
                    TF.comp_FeS2 = 0;
                    // FeS2比差值要小，那么就把FeS2变成0为止，此时的差值便为FeS2的量
                    soma.yfit[1] -= TF.comp_FeS2 * 56 / 64;
                }
            }
            else
                TF.comp_S2 -= DiffS;// 偏低时，DiffS为负数
            double DiffFe = soma.yfit[1] - soma.yreal[1];
            // 若Fe偏高，则降Fe2O3（降的量为差值 * 160 / 112），若Fe偏低，则增加Fe2O3（增的量为差值 * 160 / 112）
            TF.comp_Fe2O3 -= DiffFe * 160 / 112;// 不论偏高还是偏低，DiffFe的符号都会自动变化
            if (TF.comp_Fe2O3 < 0) TF.comp_Fe2O3 = 0;
            Mines.Add(TF);
            // 改变当前完成精矿计数（2016.11.24）
            addCurrentItems();
        }
        // 2016.11.20更新，和Reverts计算方式相同
        private void calc_COLD()
        {
            // 局部变量定义
            double[] mink = new double[] { 0, 5, 38, 7, 4 };
            double[] maxk = new double[] { 7, 15, 51, 18, 13 };
            int D = 5;
            double[,] x = new double[,] { { 0.347826087,0.8,0.634920635,0,0 },
                                          { 0.304347826,0,0.111111111,0.724137931,0.549019608 },
                                          { 0.347826087,0.2,0.253968254,0,0 } };
            double[] y = new double[] { COLD.Cu - 4.1, COLD.Fe - 0.3, COLD.S - 0 };
            // SOMA执行
            SOMA soma = new SOMA(D, mink, maxk, x, y);
            Result_K rk = soma.startMigrate();
            double[] k = rk.K;
            COLD.comp_CuFeS2 = k[0];
            COLD.comp_Cu2S = k[1];
            COLD.comp_Cu5FeS4 = k[2];
            COLD.comp_Fe3O4 = k[3];
            COLD.comp_Fe2SiO4 = k[4];
            COLD.comp_KAlSi3O8 = 2.500;
            COLD.comp_Cu = 4.1;
            COLD.comp_Fe = 0.3;
            COLD.comp_CaO = COLD.CaO;
            COLD.comp_Al2O3 = COLD.Al2O3 - 0.4586;
            COLD.comp_MgO = COLD.MgO;
            COLD.comp_SiO2 = COLD.SiO2 - 1.619 - k[4] * 0.2941;
            // 防负数判定
            if (COLD.comp_Al2O3 < 0) COLD.comp_Al2O3 = 0;
            if (COLD.comp_SiO2 < 0) COLD.comp_SiO2 = 0;
            // 计算实际值和模拟值的差异，调整混合成分（2016.10.20）
            // 模拟值-实际值
            double DiffS = soma.yfit[2] - soma.yreal[2];
            // 若S偏高，则降CuFeS2的量（降的量为差值*184/64），若S偏低，则增加S2（增量为差值）
            if (DiffS > 0)
            {
                COLD.comp_CuFeS2 -= DiffS * 184 / 64;
                k[0] = COLD.comp_CuFeS2;
            }
            else
                COLD.comp_S2 -= DiffS;// 偏低时，DiffS为负数
            // 重新计算Cu和Fe
            soma.yfit[0] = soma.calcFit(k, x, 0);
            soma.yfit[1] = soma.calcFit(k, x, 1);
            double DiffCu = soma.yfit[0] - soma.yreal[0];
            double DiffFe = soma.yfit[1] - soma.yreal[1];
            // 若Cu偏高，则降Cu（降的量为差值），若偏低，则增Cu物相的量。
            COLD.comp_Cu -= DiffCu;// 不论偏高还是偏低，DiffCu符号会自动改变
            // 若Fe偏高，则降Fe3O4（降的量为差值*232/168），若Fe偏低，则增加Fe3O4（增的量为差值*232/168）           
            COLD.comp_Fe3O4 -= DiffFe * 232 / 168;// 不论偏高还是偏低，DiffFe的符号都会自动变化
            if (COLD.comp_Fe3O4 < 0) COLD.comp_Fe3O4 = 0;
            Mines.Add(COLD);
            // 改变当前完成精矿计数（2016.11.24）
            addCurrentItems();
        }
        // 2016.11.5已更新
        private void calc_REVERTS()
        {
            // 局部变量定义
            double[] mink = new double[] { 0, 5, 38, 7, 4 };
            double[] maxk = new double[] { 7, 15, 51, 18, 13 };
            int D = 5;
            double[,] x = new double[,] { { 0.347826087,0.8,0.634920635,0,0 },
                                          { 0.304347826,0,0.111111111,0.724137931,0.549019608 },
                                          { 0.347826087,0.2,0.253968254,0,0 } };
            double[] y = new double[] { REVERTS.Cu - 4.1, REVERTS.Fe - 0.3, REVERTS.S - 0 };
            // SOMA执行
            SOMA soma = new SOMA(D, mink, maxk, x, y);
            Result_K rk = soma.startMigrate();
            double[] k = rk.K;
            REVERTS.comp_CuFeS2 = k[0];
            REVERTS.comp_Cu2S = k[1];
            REVERTS.comp_Cu5FeS4 = k[2];
            REVERTS.comp_Fe3O4 = k[3];
            REVERTS.comp_Fe2SiO4 = k[4];
            REVERTS.comp_KAlSi3O8 = 2.500;
            REVERTS.comp_Cu = 4.1;
            REVERTS.comp_Fe = 0.3;
            REVERTS.comp_CaO = REVERTS.CaO;
            REVERTS.comp_Al2O3 = REVERTS.Al2O3 - 0.4586;
            REVERTS.comp_MgO = REVERTS.MgO;
            REVERTS.comp_SiO2 = REVERTS.SiO2 - 1.619 - k[4] * 0.2941;
            // 防负数判定
            if (REVERTS.comp_Al2O3 < 0) REVERTS.comp_Al2O3 = 0;
            if (REVERTS.comp_SiO2 < 0) REVERTS.comp_SiO2 = 0;
            // 计算实际值和模拟值的差异，调整混合成分（2016.10.20）
            // 模拟值-实际值
            double DiffS = soma.yfit[2] - soma.yreal[2];
            // 若S偏高，则降CuFeS2的量（降的量为差值*184/64），若S偏低，则增加S2（增量为差值）
            if (DiffS > 0)
            {
                REVERTS.comp_CuFeS2 -= DiffS * 184 / 64;
                k[0] = REVERTS.comp_CuFeS2;
            }
            else
                REVERTS.comp_S2 -= DiffS;// 偏低时，DiffS为负数
            // 重新计算Cu和Fe
            soma.yfit[0] = soma.calcFit(k, x, 0);
            soma.yfit[1] = soma.calcFit(k, x, 1);
            double DiffCu = soma.yfit[0] - soma.yreal[0];
            double DiffFe = soma.yfit[1] - soma.yreal[1];
            // 若Cu偏高，则降Cu（降的量为差值），若偏低，则增Cu物相的量。
            REVERTS.comp_Cu -= DiffCu;// 不论偏高还是偏低，DiffCu符号会自动改变
            // 若Fe偏高，则降Fe3O4（降的量为差值*232/168），若Fe偏低，则增加Fe3O4（增的量为差值*232/168）           
            REVERTS.comp_Fe3O4 -= DiffFe * 232 / 168;// 不论偏高还是偏低，DiffFe的符号都会自动变化
            if (REVERTS.comp_Fe3O4 < 0) REVERTS.comp_Fe3O4 = 0;
            Mines.Add(REVERTS);
            // 改变当前完成精矿计数（2016.11.24）
            addCurrentItems();
        }
        private void calc_LUBAMBE()
        {
            // 局部变量定义
            double[] mink = new double[] { 8, 35, 0, 0, 0 };
            double[] maxk = new double[] { 12, 45, 5, 9, 5 };
            int D = 5;
            double[,] x = new double[,] { { 0.347826087,0.8,0.634920635,0.576576577,0 },
                                          { 0.304347826,0,0.111111111,0,0.7 },
                                          { 0.347826087,0.2,0.253968254,0,0 } };
            double[] y = new double[] { LUBAMBE.Cu - 0.095555556, LUBAMBE.Fe - 0.018666667, LUBAMBE.S - 0.051333333 };
            // SOMA执行
            SOMA soma = new SOMA(D, mink, maxk, x, y);
            Result_K rk = soma.startMigrate();
            double[] k = rk.K;
            LUBAMBE.comp_CuFeS2 = k[0];
            LUBAMBE.comp_Cu2S = k[1];
            LUBAMBE.comp_Cu5FeS4 = k[2];
            LUBAMBE.comp_Cu2_OH_2CO3 = k[3];
            LUBAMBE.comp_Fe2O3 = k[4];
            LUBAMBE.comp_CuS = 0.090;
            LUBAMBE.comp_FeS2 = 0.040;
            LUBAMBE.comp_KAlSi3O8 = (LUBAMBE.Al2O3 - 1.034) * 5.451;
            LUBAMBE.comp_SiO2 = LUBAMBE.SiO2 - LUBAMBE.comp_KAlSi3O8 * 0.6475 - 1.217;
            LUBAMBE.comp_KAl2_AlSi3O10__OH_2 = 2.690;
            LUBAMBE.comp_CaMg_CO3_2 = 0.000;
            LUBAMBE.comp_Cu2O = 0.040;
            LUBAMBE.comp_CaO = LUBAMBE.CaO - 0.0426;
            LUBAMBE.comp_MgO = LUBAMBE.MgO - 0.0304;
            // 防负数判定
            if (LUBAMBE.comp_KAlSi3O8 < 0) LUBAMBE.comp_KAlSi3O8 = 0;
            if (LUBAMBE.comp_SiO2 < 0) LUBAMBE.comp_SiO2 = 0;
            if (LUBAMBE.comp_CaO < 0) LUBAMBE.comp_CaO = 0;
            if (LUBAMBE.comp_MgO < 0) LUBAMBE.comp_MgO = 0;
            // 计算实际值和模拟值的差异，调整混合成分（2016.10.20）
            // 模拟值-实际值
            double DiffCu = soma.yfit[0] - soma.yreal[0];
            // 若Cu偏高，则降Cu2S（降的量为差值*160/128），若偏低，则增Cu物相的量。
            if (DiffCu > 0)
            {
                LUBAMBE.comp_Cu2S -= DiffCu * 160 / 128;
                k[1] = LUBAMBE.comp_Cu2S;
            }
            else
                LUBAMBE.comp_Cu -= DiffCu;// 偏低时，DiffCu为负数
            // 重新计算S
            soma.yfit[2] = soma.calcFit(k, x, 2);
            double DiffS = soma.yfit[2] - soma.yreal[2];
            // 若S偏高，则降FeS2的量（降的量为差值*120/64），若S偏低，则增加S2（增量为差值）
            if (DiffS > 0)
            {
                // 自己理解：降低了FeS2，将会同步降低Fe，降的量为上述差值*120/56
                if (LUBAMBE.comp_FeS2 > DiffS * 120 / 64)
                {
                    // 大于的话才能减，小于等于置为0
                    LUBAMBE.comp_FeS2 -= DiffS * 120 / 64;
                    // 同步降低Fe
                    soma.yfit[1] -= DiffS * 56 / 64;
                }
                else
                {
                    LUBAMBE.comp_FeS2 = 0;
                    // FeS2比差值要小，那么就把FeS2变成0为止，此时的差值便为FeS2的量（可能产生误差）
                    soma.yfit[1] -= LUBAMBE.comp_FeS2 * 56 / 64;
                }
            }
            else
                LUBAMBE.comp_S2 -= DiffS;// 偏低时，DiffS为负数
            double DiffFe = soma.yfit[1] - soma.yreal[1];
            // 若Fe偏高，则降Fe2O3（降的量为差值*160/112），若Fe偏低，则增加Fe2O3（增的量为差值*160/112）。
            LUBAMBE.comp_Fe2O3 -= DiffFe * 160 / 112;// 不论偏高还是偏低，DiffFe的符号都会自动变化
            if (LUBAMBE.comp_Fe2O3 < 0) LUBAMBE.comp_Fe2O3 = 0;
            Mines.Add(LUBAMBE);
            // 改变当前完成精矿计数（2016.11.24）
            addCurrentItems();
        }

        private void calc_NFCA()
        {
            // 局部变量定义
            double[] mink = new double[] { 45, 0, 8, 4 };
            double[] maxk = new double[] { 62, 6, 18, 14 };
            int D = 4;
            double[,] x = new double[,] { { 0.347826087,0.666666667,0.634920635,0 },
                                          { 0.304347826,0,0.111111111,0.466666667 },
                                          { 0.347826087,0.333333333,0.253968254,0.533333333 } };
            double[] y = new double[] { NFCA.Cu - 1.04, NFCA.Fe - 0.35, NFCA.S - 0.26 };
            // SOMA执行
            SOMA soma = new SOMA(D, mink, maxk, x, y);
            Result_K rk = soma.startMigrate();
            double[] k = rk.K;
            NFCA.comp_CuFeS2 = k[0];
            NFCA.comp_CuS = k[1];
            NFCA.comp_Cu5FeS4 = k[2];
            NFCA.comp_FeS2 = k[3];
            NFCA.comp_Cu2S = 1.300;
            NFCA.comp_Fe2O3 = 0.500;
            NFCA.comp_KAl2_AlSi3O10__OH_2 = NFCA.Al2O3 * 2.6;
            NFCA.comp_SiO2 = NFCA.SiO2 - NFCA.comp_KAl2_AlSi3O10__OH_2 * 0.4523;
            NFCA.comp_CaO = NFCA.CaO;
            NFCA.comp_MgO = NFCA.MgO;
            if (NFCA.comp_SiO2 < 0) NFCA.comp_SiO2 = 0;
            // 计算实际值和模拟值的差异，调整混合成分（2016.10.20）
            // 模拟值-实际值
            double DiffCu = soma.yfit[0] - soma.yreal[0];
            // 若Cu偏高，则降CuFeS2（降的量为差值*184/64），若偏低，则增Cu物相的量。
            if (DiffCu > 0)
            {
                NFCA.comp_CuFeS2 -= DiffCu * 184 / 64;
                k[0] = NFCA.comp_CuFeS2;
            }
            else
                NFCA.comp_Cu -= DiffCu;// 偏低时，DiffCu为负数
            // 重新计算S和Fe
            soma.yfit[2] = soma.calcFit(k, x, 2);
            soma.yfit[1] = soma.calcFit(k, x, 1);
            double DiffS = soma.yfit[2] - soma.yreal[2];
            // 若S偏高，则降FeS2的量（降的量为差值*120/64），若S偏低，则增加S2（增量为差值）
            if (DiffS > 0)
            {
                // 自己理解：降低了FeS2，将会同步降低Fe，降的量为上述差值*56/64
                if (NFCA.comp_FeS2 > DiffS * 120 / 64)
                {
                    // 大于的话才能减，小于等于置为0
                    NFCA.comp_FeS2 -= DiffS * 120 / 64;
                    // 同步降低Fe
                    soma.yfit[1] -= DiffS * 56 / 64;
                }
                else
                {
                    NFCA.comp_FeS2 = 0;
                    // FeS2比差值要小，那么就把FeS2变成0为止，此时的差值便为FeS2的量
                    soma.yfit[1] -= NFCA.comp_FeS2 * 56 / 64;
                }
            }
            else
                NFCA.comp_S2 -= DiffS;// 偏低时，DiffS为负数
            double DiffFe = soma.yfit[1] - soma.yreal[1];
            // 若Fe偏高，则降Fe2O3（降的量为差值*160/112），若Fe偏低，则增加Fe2O3（增的量为差值*160/112）。
            NFCA.comp_Fe2O3 -= DiffFe * 160 / 112;// 不论偏高还是偏低，DiffFe的符号都会自动变化
            if (NFCA.comp_Fe2O3 < 0) NFCA.comp_Fe2O3 = 0;
            Mines.Add(NFCA);
            // 改变当前完成精矿计数（2016.11.24）
            addCurrentItems();
        }
        // 2016.11.20更新此精矿数据
        private void calc_BOLO()
        {
            double[] mink = new double[] { 8, 1, 3, 4, 0, 10 };
            double[] maxk = new double[] { 17, 10, 11, 13, 8, 21 };
            int D = 6;
            double[,] x = new double[,] { { 0.347826087,0.8,0.634920635,0,0,0.576576577 },
                                          { 0.304347826,0,0.111111111,0.7,0.466666667,0 },
                                          { 0.347826087,0.2,0.253968254,0,0.533333333,0 } };
            double[] y = new double[] { BOLO.Cu - 3.153661282, BOLO.Fe - 0.709655172, BOLO.S - 0.271013216 };
            // SOMA执行
            SOMA soma = new SOMA(D, mink, maxk, x, y);
            Result_K rk = soma.startMigrate();
            double[] k = rk.K;
            BOLO.comp_CuFeS2 = k[0];
            BOLO.comp_Cu2S = k[1];
            BOLO.comp_Cu5FeS4 = k[2];
            BOLO.comp_Fe2O3 = k[3];
            BOLO.comp_FeS2 = k[4];
            BOLO.comp_Cu2_OH_2CO3 = k[5];
            BOLO.comp_Al2O3 = BOLO.Al2O3 - 0.495;
            BOLO.comp_Mg6Si8O20_OH_4 = BOLO.MgO * 3.15;
            BOLO.comp_SiO2 = BOLO.SiO2 - 1.75 - BOLO.comp_Mg6Si8O20_OH_4 * 0.635;
            BOLO.comp_CaO = BOLO.CaO;
            BOLO.comp_CuS = 0.24;
            BOLO.comp_Cu4SO4_OH_6 = 2.71;
            BOLO.comp_KAlSi3O8 = 2.7;
            BOLO.comp_KAl2_AlSi3O10__OH_2 = 0.08;
            BOLO.comp_C = 14.75;
            BOLO.comp_Cu2O = 1.12;
            BOLO.comp_Cu = 0.47;
            BOLO.comp_Fe3O4 = 0.98;
            // 计算实际值和模拟值的差异，调整混合成分（2016.11.20）
            // 模拟值-实际值
            double DiffCu = soma.yfit[0] - soma.yreal[0];
            // 拟合完成后，若Cu偏高，则降Cu2(OH)2CO3（降的量为差值*111/64），若偏低，则增Cu2(OH)2CO3物相的量（增的量为差值*111/64）
            BOLO.comp_Cu2_OH_2CO3 -= DiffCu * 111 / 64;
            k[5] = BOLO.comp_Cu2_OH_2CO3;
            double DiffS = soma.yfit[2] - soma.yreal[2];
            // 若S偏高，则降FeS2的量（降的量为差值 * 120 / 64），若S偏低，则增加S2（增量为差值）
            if (DiffS > 0)
            {
                // 自己理解：降低了FeS2，将会同步降低Fe，降的量为上述差值*56/64
                if (BOLO.comp_FeS2 > DiffS * 120 / 64)
                {
                    // 大于的话才能减，小于等于置为0
                    BOLO.comp_FeS2 -= DiffS * 120 / 64;
                    // 同步降低Fe
                    soma.yfit[1] -= DiffS * 56 / 64;
                }
                else
                {
                    BOLO.comp_FeS2 = 0;
                    // FeS2比差值要小，那么就把FeS2变成0为止，此时的差值便为FeS2的量
                    soma.yfit[1] -= BOLO.comp_FeS2 * 56 / 64;
                }
            }
            else
                BOLO.comp_S2 -= DiffS;// 偏低时，DiffS为负数
            double DiffFe = soma.yfit[1] - soma.yreal[1];
            // 若Fe偏高，则降Fe2O3（降的量为差值 * 160 / 112），若Fe偏低，则增加Fe2O3（增的量为差值 * 160 / 112）
            BOLO.comp_Fe2O3 -= DiffFe * 160 / 112;// 不论偏高还是偏低，DiffFe的符号都会自动变化
            if (BOLO.comp_Fe2O3 < 0) BOLO.comp_Fe2O3 = 0;
            Mines.Add(BOLO);
            // 改变当前完成精矿计数（2016.11.24）
            addCurrentItems();
        }
        private void calc_CCS()
        {
            // 局部变量定义
            double[] mink = new double[] { 3, 7, 10, 0, 15 };
            double[] maxk = new double[] { 15, 16, 20, 5, 28 };
            int D = 5;
            double[,] x = new double[,] { { 0.347826087, 0.8, 0.634920635, 1, 0 },
                                          { 0.304347826, 0, 0.111111111, 0, 0.549019608 },
                                          { 0.347826087, 0.2, 0.253968254, 0, 0 } };
            double[] y = new double[] { CCS.Cu - 0.333333333, CCS.Fe - 8.237586207, CCS.S - 0.166666667 };
            // SOMA执行
            SOMA soma = new SOMA(D, mink, maxk, x, y);
            Result_K rk = soma.startMigrate();
            double[] k = rk.K;
            CCS.comp_CuFeS2 = k[0];
            CCS.comp_Cu2S = k[1];
            CCS.comp_Cu5FeS4 = k[2];
            CCS.comp_Cu = k[3];
            CCS.comp_Fe2SiO4 = k[4];
            CCS.comp_CuS = 0.500;
            CCS.comp_Fe2O3 = 0.300;
            CCS.comp_SiO2 = CCS.SiO2 - 3.432 - k[4] * 0.2941;
            CCS.comp_KAlSi3O8 = (CCS.Al2O3 - 0.972) * 5.451;
            CCS.comp_Fe3O4 = 8.600;
            CCS.comp_Fe = 1.800;
            CCS.comp_CaO = CCS.CaO;
            CCS.comp_Al2O3 = 0.972;
            CCS.comp_K2O = 0.700;
            CCS.comp_MgO = CCS.MgO;
            // 防负数判定
            if (CCS.comp_SiO2 < 0) CCS.comp_SiO2 = 0;
            if (CCS.comp_KAlSi3O8 < 0) CCS.comp_KAlSi3O8 = 0;
            // 计算实际值和模拟值的差异，调整混合成分（2016.10.20）
            // 模拟值-实际值
            double DiffCu = soma.yfit[0] - soma.yreal[0];
            // 若Cu偏高，则降Cu5FeS4（降的量为差值*504/320），若偏低，则增Cu物相的量。
            if (DiffCu > 0)
            {
                CCS.comp_Cu5FeS4 -= DiffCu * 504 / 320;
                k[2] = CCS.comp_Cu5FeS4;
            }
            else
                CCS.comp_Cu -= DiffCu;// 偏低时，DiffCu为负数
            // 重新计算S和Fe
            soma.yfit[2] = soma.calcFit(k, x, 2);
            soma.yfit[1] = soma.calcFit(k, x, 1);
            double DiffS = soma.yfit[2] - soma.yreal[2];
            // 若S偏高，暂无方法，若S偏低，则增加S2（增量为差值）
            if (DiffS < 0)
                CCS.comp_S2 -= DiffS;// 偏低时，DiffS为负数
            double DiffFe = soma.yfit[1] - soma.yreal[1];
            // 若Fe偏高，则降Fe2SiO4（降的量为差值*204/112），若Fe偏低，则增Fe2SiO4（增的量为差值*204/112）。
            CCS.comp_Fe2SiO4 -= DiffFe * 204 / 112;// 不论偏高还是偏低，DiffFe的符号都会自动变化
            if (CCS.comp_Fe2SiO4 < 0) CCS.comp_Fe2SiO4 = 0;
            Mines.Add(CCS);
            // 改变当前完成精矿计数（2016.11.24）
            addCurrentItems();
        }
        #endregion

        #region  熔炼模拟——用户修改推荐值之后再调用FactSage进行计算
        private void predict_Click(object sender, EventArgs e)
        {
            // 获取用户修改过后的氧料比、补硅、补钙、补煤的推荐值
            suggestOxyRatio = Convert.ToDouble(this.textBox65.Text.ToString());
            suggestAddSi = Convert.ToDouble(this.textBox61.Text.ToString());
            suggestAddCa = Convert.ToDouble(this.textBox60.Text.ToString());
            suggestAddCoal = Convert.ToDouble(this.textBox58.Text.ToString());
            //将上述4个变量写入Excel
            ModifyExcelFileRecommend(suggestOxyRatio, suggestAddSi, suggestAddCa, suggestAddCoal);
            // 初始化progress窗体
            calcprogress = new CalcProgress();
            progressTXT = calcprogress.textBox;
            progressBAR = calcprogress.progressBar;
            progressPERCENT = calcprogress.label;
            calcprogress.Show();
            progressWindowChange("准备使用FactSage进行迭代计算...");
            //调用FactSage计算
            BackgroundWorker bwFactSagePredict = new BackgroundWorker();
            bwFactSagePredict.DoWork += new DoWorkEventHandler(bwFactSagePredict_DoWork);
            bwFactSagePredict.WorkerSupportsCancellation = true;
            bwFactSagePredict.WorkerReportsProgress = true;
            bwFactSagePredict.ProgressChanged += new ProgressChangedEventHandler(bwFactSagePredict_ProgressChanged);
            bwFactSagePredict.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwFactSagePredict_RunWorkerCompleted);
            bwFactSagePredict.RunWorkerAsync();
        }
        private void bwFactSagePredict_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bwFactSagePredict = sender as BackgroundWorker;
            CallFactSageEquilib(bwFactSagePredict);
        }
        private void bwFactSagePredict_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }
        private void bwFactSagePredict_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressWindowChange("FactSage计算完毕，已根据修改的推荐值重新计算！");
            //ExportProcessToExcel();
            progressBarChange(100);
            calcprogress.startTimer();
            fillResult();
        }
        // 用户修改氧料比，补硅，补钙，补煤的推荐值，写入Excel
        private void ModifyExcelFileRecommend(double OxygenRatio, double SupplementSi, double SupplementCa, double SupplementCoal)
        {
            string tempPath = commonFilePath;
            HSSFWorkbook wk = null;
            using (FileStream fs = File.Open(tempPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                wk = new HSSFWorkbook(fs);
                fs.Close();
            }
            //打开第四个工作表
            ISheet tb = wk.GetSheetAt(3);
            tb.ForceFormulaRecalculation = true;
            ICell cell = null;
            //修改文件中氧料比的值
            cell = tb.GetRow(33).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(OxygenRatio);
            //修改文件中补硅的值
            cell = tb.GetRow(37).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(SupplementSi);
            //修改文件中补钙的值
            cell = tb.GetRow(38).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(SupplementCa);
            //修改文件中补煤的值
            cell = tb.GetRow(39).GetCell(2);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(SupplementCoal);
            // 设置前两个表的公式自动计算
            tb = wk.GetSheetAt(0);
            tb.ForceFormulaRecalculation = true;
            tb = wk.GetSheetAt(1);
            tb.ForceFormulaRecalculation = true;
            // 保存文件
            using (FileStream fs = File.Open(tempPath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                wk.Write(fs);
                fs.Close();
            }
        }
        #endregion

        #region  导出配料单（2018.04添加）
        public static string SelectPath()
        {
            string path = string.Empty;
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                path = fbd.SelectedPath;
            }
            return path;
        }
        private void Export_file_Click(object sender, EventArgs e)
        {
            //复制配料单模板
            string id = this.comboBox1.Text.ToString().Split('等')[0];
            string name = id + "-配料单.xls";
            string LocalFile = Thread.GetDomain().BaseDirectory + "配料单模板.xls";//要复制的文件路径
            string SaveFile = SelectPath() + "\\" + name;//指定存储的路径

            if (File.Exists(LocalFile))//必须判断要复制的文件是否存在
            {
                File.Copy(LocalFile, SaveFile, true);//三个参数分别是源文件路径，存储路径，若存储路径有相同文件是否替换
            }


            //打开文件
            HSSFWorkbook wk = null;
            using (FileStream fs = File.Open(SaveFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                wk = new HSSFWorkbook(fs);
                fs.Close();
            }
            ISheet tb = wk.GetSheetAt(0);
            IRow row = null;
            ICell cell = null;

            cell = tb.GetRow(1).GetCell(1);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(DateTime.Now.ToString());

            cell = tb.GetRow(1).GetCell(3);
            cell.SetCellType(CellType.Blank);
            cell.SetCellValue(id);


            //【上部】精矿+数据
            for (int i = 0; i < 7; i++)
            {
                row = tb.GetRow(i + 3);
                for (int j = 0; j < 9; j++)
                {
                    cell = row.GetCell(j);
                    cell.SetCellType(CellType.Blank);
                    cell.SetCellValue(Mdata[i, j + 1]);
                }
            }

            //【上部】总成分
            row = tb.GetRow(10);
            cell = row.GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox46.Text));
            cell = row.GetCell(2);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox47.Text));
            cell = row.GetCell(3);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox48.Text));
            cell = row.GetCell(4);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox49.Text));
            cell = row.GetCell(5);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox50.Text));
            cell = row.GetCell(6);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox51.Text));
            cell = row.GetCell(7);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox52.Text));
            cell = row.GetCell(8);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox53.Text));

            //【中部】初始参数——精矿下料量、富氧浓度、氧气纯度、喷枪端压
            cell = tb.GetRow(12).GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox25.Text));
            cell = tb.GetRow(13).GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox55.Text));
            cell = tb.GetRow(14).GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox54.Text));
            cell = tb.GetRow(15).GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox23.Text));

            //【中部】目标参数——冰铜品位（%）、Fe3O4、SiO2/Fe、SiO2/CaO
            cell = tb.GetRow(12).GetCell(3);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox1.Text));
            cell = tb.GetRow(13).GetCell(3);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox2.Text));
            cell = tb.GetRow(14).GetCell(3);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox3.Text));
            cell = tb.GetRow(15).GetCell(3);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox4.Text));

            //【中部】推荐参数——氧料比、风量、氧气流量、补硅、补钙、补煤
            cell = tb.GetRow(12).GetCell(5);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox65.Text));
            cell = tb.GetRow(13).GetCell(5);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox63.Text));
            cell = tb.GetRow(14).GetCell(5);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox62.Text));
            cell = tb.GetRow(15).GetCell(5);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox61.Text));
            cell = tb.GetRow(16).GetCell(5);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox60.Text));
            cell = tb.GetRow(17).GetCell(5);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox58.Text));

            //【中部】预测值——冰铜品位（%）、Fe3O4、SiO2/Fe、SiO2/CaO
            cell = tb.GetRow(12).GetCell(7);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox67.Text));
            cell = tb.GetRow(13).GetCell(7);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox68.Text));
            cell = tb.GetRow(14).GetCell(7);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox69.Text));
            cell = tb.GetRow(15).GetCell(7);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(Convert.ToDouble(this.textBox70.Text));

            //打开参数推荐文件，从里面选取需要的数据复制到配料单
            HSSFWorkbook wk2 = null;
            string path = "D:\\ExpertSystem\\" + id + ".xls";

            using (FileStream fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                wk2 = new HSSFWorkbook(fs);
                fs.Close();
            }
            ISheet tb2 = wk2.GetSheetAt(3);
            //【下部】——气体
            row = tb.GetRow(24);
            cell = row.GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(4).GetCell(2).NumericCellValue);
            cell = row.GetCell(2);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(4).GetCell(3).NumericCellValue);
            cell = row.GetCell(3);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(4).GetCell(4).NumericCellValue);
            cell = row.GetCell(4);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(4).GetCell(5).NumericCellValue);
            cell = row.GetCell(5);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(4).GetCell(6).NumericCellValue);
            cell = row.GetCell(6);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(4).GetCell(7).NumericCellValue);
            cell = row.GetCell(7);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(4).GetCell(8).NumericCellValue);
            cell = row.GetCell(8);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(4).GetCell(9).NumericCellValue);
            cell = row.GetCell(9);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(4).GetCell(10).NumericCellValue);
            cell = row.GetCell(10);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(4).GetCell(11).NumericCellValue);

            row = tb.GetRow(25);
            cell = row.GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(5).GetCell(2).NumericCellValue);
            cell = row.GetCell(2);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(5).GetCell(3).NumericCellValue);
            cell = row.GetCell(3);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(5).GetCell(4).NumericCellValue);
            cell = row.GetCell(4);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(5).GetCell(5).NumericCellValue);
            cell = row.GetCell(5);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(5).GetCell(6).NumericCellValue);
            cell = row.GetCell(6);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(5).GetCell(7).NumericCellValue);
            cell = row.GetCell(7);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(5).GetCell(8).NumericCellValue);
            cell = row.GetCell(8);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(5).GetCell(9).NumericCellValue);
            cell = row.GetCell(9);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(5).GetCell(10).NumericCellValue);
            cell = row.GetCell(10);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(5).GetCell(11).NumericCellValue);

            row = tb.GetRow(26);
            cell = row.GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(6).GetCell(2).NumericCellValue);
            cell = row.GetCell(2);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(6).GetCell(3).NumericCellValue);
            cell = row.GetCell(3);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(6).GetCell(4).NumericCellValue);
            cell = row.GetCell(4);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(6).GetCell(5).NumericCellValue);
            cell = row.GetCell(5);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(6).GetCell(6).NumericCellValue);
            cell = row.GetCell(6);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(6).GetCell(7).NumericCellValue);
            cell = row.GetCell(7);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(6).GetCell(8).NumericCellValue);
            cell = row.GetCell(8);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(6).GetCell(9).NumericCellValue);
            cell = row.GetCell(9);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(6).GetCell(10).NumericCellValue);
            cell = row.GetCell(10);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(6).GetCell(11).NumericCellValue);
            //冰铜
            row = tb.GetRow(28);
            cell = row.GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(9).GetCell(1).NumericCellValue);
            cell = row.GetCell(2);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(9).GetCell(10).NumericCellValue);
            cell = row.GetCell(3);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(9).GetCell(11).NumericCellValue);
            cell = row.GetCell(4);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(9).GetCell(12).NumericCellValue);

            row = tb.GetRow(29);
            cell = row.GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(10).GetCell(1).NumericCellValue);
            cell = row.GetCell(2);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(10).GetCell(10).NumericCellValue);
            cell = row.GetCell(3);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(10).GetCell(11).NumericCellValue);
            cell = row.GetCell(4);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(10).GetCell(12).NumericCellValue);

            cell = tb.GetRow(30).GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(50).GetCell(1).NumericCellValue);

            //熔渣
            row = tb.GetRow(32);
            cell = row.GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(23).GetCell(1).NumericCellValue);
            cell = row.GetCell(2);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(23).GetCell(2).NumericCellValue);
            cell = row.GetCell(3);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(23).GetCell(3).NumericCellValue);
            cell = row.GetCell(4);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(23).GetCell(4).NumericCellValue);
            cell = row.GetCell(5);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(23).GetCell(5).NumericCellValue);
            cell = row.GetCell(6);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(23).GetCell(6).NumericCellValue);
            cell = row.GetCell(7);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(23).GetCell(7).NumericCellValue);
            cell = row.GetCell(8);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(23).GetCell(8).NumericCellValue);
            cell = row.GetCell(9);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(23).GetCell(9).NumericCellValue);
            cell = row.GetCell(10);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(23).GetCell(10).NumericCellValue);
            cell = row.GetCell(11);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(23).GetCell(11).NumericCellValue);
            cell = row.GetCell(12);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(23).GetCell(12).NumericCellValue);

            row = tb.GetRow(33);
            cell = row.GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(24).GetCell(1).NumericCellValue);
            cell = row.GetCell(2);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(24).GetCell(2).NumericCellValue);
            cell = row.GetCell(3);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(24).GetCell(3).NumericCellValue);
            cell = row.GetCell(4);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(24).GetCell(4).NumericCellValue);
            cell = row.GetCell(5);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(24).GetCell(5).NumericCellValue);
            cell = row.GetCell(6);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(24).GetCell(6).NumericCellValue);
            cell = row.GetCell(7);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(24).GetCell(7).NumericCellValue);
            cell = row.GetCell(8);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(24).GetCell(8).NumericCellValue);
            cell = row.GetCell(9);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(24).GetCell(9).NumericCellValue);
            cell = row.GetCell(10);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(24).GetCell(10).NumericCellValue);
            cell = row.GetCell(11);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(24).GetCell(11).NumericCellValue);
            cell = row.GetCell(12);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(24).GetCell(12).NumericCellValue);

            cell = tb.GetRow(34).GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(51).GetCell(1).NumericCellValue);
            cell = tb.GetRow(35).GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(52).GetCell(1).NumericCellValue);
            cell = tb.GetRow(36).GetCell(1);
            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(tb2.GetRow(53).GetCell(1).NumericCellValue);

            using (FileStream fs = File.Open(SaveFile, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                wk.Write(fs);
                fs.Close();
                MessageBox.Show("配料单导出成功！");
            }

        }
        #endregion
    }
}
