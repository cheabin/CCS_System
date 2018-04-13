using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace CCS_System
{
    public class SOMA
    {
        double PathLength, Step, PRT, MinDiv, EPS;
        int P, D, M;
        // 拟合参数的范围
        double[] mink = new double[10];
        double[] maxk = new double[10];
        double[][] kstart = new double[20][];
        // 种群所有个体当前位置的函数值
        double[] fstart = new double[20];
        // 定义Leader
        double[] k = new double[10];
        // 最优函数值
        double fbest;
        // 查找过程中的最优函数值
        double ftbest;
        // 查找过程中的最优函数值对应的k值
        double[] kt = new double[10];
        // 最差函数值
        double fworst;
        double[,] x = new double[3,10];
        double[] y = new double[3];
        // 定义实际y和拟合y（2016.10.20更新）
        public double[] yreal = new double[3];
        public double[] yfit = new double[3];

        Result_K rk = new Result_K();

        public SOMA()
        {
            init();
        }

        public SOMA(int D,double[] mink,double[] maxk,double[,] x,double[] y)
        {
            this.D = D;
            this.mink = mink;
            this.maxk = maxk;
            this.x = x;
            this.y = y;
            init();
        }

        private void init()
        {
            Random r = new Random();
            PathLength = 2;
            P = 20;
            Step = 0.11;
            PRT = 1;
            M = 1000;
            MinDiv = 0.000001;
            EPS = 0.001;
            double ftmp;

            for (int j = 0; j < P; j++)
            {
                kstart[j] = new double[10];
                // 初始化种群
                for (int i = 0; i < D; i++)
                {
                    kstart[j][i] = mink[i] + (maxk[i] - mink[i]) * r.NextDouble();
                }
                ftmp = f(kstart[j], x, y);
                // 初始化Leader为整个种群中的最优值
                if (j == 0)
                {
                    fbest = ftmp;
                    k = kstart[j];
                }
                else
                {
                    if (ftmp < fbest)
                    {
                        fbest = ftmp;
                        k = kstart[j];
                    }
                }
            }
            // 初始化搜索过程最优值
            ftbest = fbest;
            for (int i = 0; i < D; i++)
            {
                kt[i] = k[i];
            }
        }

        private int PRTVector()
        {
            Random r = new Random();
            if (r.NextDouble() < PRT) return 1;
            else return 0;
        }

        // 供外部调用的函数
        public Result_K startMigrate()
        {
            rk.Result = 0;
            bool find = false;
            for (int i = 0; i < M; i++)
            {
                if (migrate() == 1)
                {
                    find = true;
                    break;
                }
            }
            // 未找到满足条件的最优值，输出找的过程中最优的那个
            if (!find)
            {
                fbest = ftbest;
                k = kt;
            }
            // 存储实际值和拟合值（2016.10.20更新）
            this.yreal[0] = y[0];
            this.yreal[1] = y[1];
            this.yreal[2] = y[2];
            this.yfit[0] = y1(k, x, 0);
            this.yfit[1] = y1(k, x, 1);
            this.yfit[2] = y1(k, x, 2);
            rk.K = this.k;
            return rk;
        }

        // 迁移算法
        private int migrate()
        {
            // 当前步长
            double t;
            // 暂时计算出来的优化函数值
            double ftmp = 0;
            // 迁移过程中的某一步
            double[] ktmp = new double[10];
            // 本次迁移过程的最优解
            double[] kpbest = new double[10];
            // 本次迁移过程中的最优值
            double fpbest;
            // 本次迁移过程中最佳个体标志
            int ibest = 0;
            // 混合迁移参数
            double r = 0.3;

            for (int i = 0; i < P; i++)
            {
                t = Step;
                kpbest = kstart[i];
                fpbest = f(kpbest, x, y);
                
                while (t <= PathLength)
                {
                    for (int j = 0; j < D; j++)
                    {
                        Random randh = new Random();
                        if (randh.NextDouble() < r)
                        {
                            // 满足混合迁移概率后，不向Leader迁移，而随机个体相反的方向迁移
                            Random randi = new Random();
                            int itmp = randi.Next(20);
                            // 向随机个体迁移的公式
                            ktmp[j] = kstart[i][j] + (kstart[i][j] - kstart[itmp][j]) * t * PRTVector();
                        }
                        else
                        {
                            // 正常迁移
                            ktmp[j] = kstart[i][j] + (k[j] - kstart[i][j]) * t * PRTVector();
                        }
                        
                        // k不能超过范围
                        if (ktmp[j] < mink[j])
                        {
                            ktmp[j] = mink[j];
                        }
                        if (ktmp[j] > maxk[j])
                        { 
                            ktmp[j] = maxk[j]; 
                        }
                    }
                    // 计算本次迁移的函数值
                    ftmp = f(ktmp, x, y);
                    // 保存本次迁移的最优值
                    if (ftmp < fpbest)
                    {
                        fpbest = ftmp;
                        for (int l = 0; l < D; l++)
                        {
                            kpbest[l] = ktmp[l];
                        }
                    }
                    t += Step;
                }
                // 完成这一个个体的迁移
                for (int j = 0; j < D; j++)
                {
                    kstart[i][j] = kpbest[j];
                }
                // 保存当前迁移个体的函数值
                fstart[i] = fpbest;
            }
            // 初始化当前最坏值
            fworst = fbest;
            // 一次迁移全部完成，选出新Leader
            for (int i = 0; i < P; i++)
            {
                if (fstart[i] < fbest)
                {
                    ibest = i;
                    // 存储新Leader
                    fbest = fstart[i];
                    k = kstart[i];
                }
                // 存储最坏值
                if (fstart[i] > fworst) fworst = fstart[i];
            }

            // 最坏和最好的差值小于指定精度，并且EPS满足要求
            if (Math.Abs(fworst - fbest) < MinDiv)
            {
                if (fbest < EPS) return 1;
                else
                {
                    modify();
                }
            } 
            return 0;
        }

        // 调整种群
        private void modify()
        {
            // 调整前，存储本轮最优值
            if (fbest < ftbest)
            {
                ftbest = fbest;
                for (int i = 0; i < D; i++)
                {
                    kt[i] = k[i];
                }
            }
            Random r = new Random();
            for (int i = 0; i < P; i++)
            {
                for (int j = 0; j < D; j++)
                {
                    // 重来
                    kstart[i][j] = mink[j] + (maxk[j] - mink[j]) * r.NextDouble();
                }
                double ftmp = f(kstart[i], x, y);
                // 初始化Leader为整个种群中的最优值
                if (i == 0)
                {
                    fbest = ftmp;
                    k = kstart[i];
                }
                else
                {
                    if (ftmp < fbest)
                    {
                        fbest = ftmp;
                        k = kstart[i];
                    }
                }
            }
        }
       
        // 目标函数的计算
        private double f(double[] k, double[,] x, double[] y)
        {
            double target = 0;
            // 求所有三组数据中输出值与拟合值差值的加和
            for (int j = 0; j < 3; j++)
            {
                for (int i = 0; i < D; i++)
                {
                    target += Math.Pow((y[j] - y1(k, x, j)), 2);
                }
            }
            return target;
        }

        // 用于计算拟合值的函数
        private double y1(double[] k, double[,] x, int i)
        {
            double result = 0;
            for (int j = 0; j < D; j++)
            {
                result += k[j] * x[i,j];
            }
            return result;
        }

        // 供外部调用的计算拟合值的函数
        public double calcFit(double[] k, double[,] x, int i)
        {
            return y1(k, x, i);
        }
    }
}
