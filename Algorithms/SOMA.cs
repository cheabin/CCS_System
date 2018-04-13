using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Algorithms
{
    public class SOMA
    {
        // 定义最大迁移路径，步长参数，PRT参数，差异值
        double PathLength, Step, PRT, MinDiv;
        // 定义种群数量，问题维数，最大迁移次数
        int P, D, M;
        // 定义拟合参数的范围
        double[] mink = new double[10];
        double[] maxk = new double[10];
        // 定义所有种群个体(交错数组)
        double[][] kstart = new double[20][];
        // 定义种群所有个体当前位置算出的函数值
        double[] fstart = new double[20];
        // 定义最终输出的参数，即Leader，从主调函数接收地址
        double[] k = null;
        // 定义最优函数值
        double fbest;
        // 定义最差函数值
        double fworst;
        // 定义X和Y
        double[,] x = new double[3,10];
        double[] y = new double[3];

        public SOMA()
        {
            init();
        }
        // 含参构造函数
        public SOMA(int D,double[] mink,double[] maxk,double[,] x,double[] y,double[] k)
        {
            this.D = D;
            this.mink = mink;
            this.maxk = maxk;
            this.x = x;
            this.y = y;
            this.k = k;
            init();
        }
        // 初始化
        private void init()
        {
            Random r = new Random();
            PathLength = 2;
            P = 20;
            Step = 0.11;
            PRT = 0.4;
            MinDiv = 0.000001;
            // 暂时计算出来的优化函数值
            double ftmp;
            for (int j = 0; j < P; j++)
            {
                // 初始化交错数组的每一行
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
            
        }

        private int PRTVector()
        {
            Random r = new Random();
            if (r.NextDouble() < PRT) return 1;
            else return 0;
        }

        // 供外部调用的函数
        public int startMigrate()
        {
            return migrate();
        }

        // 迁移算法
        private int migrate()
        {
            // 当前步长
            double t;
            // 暂时计算出来的优化函数值
            double ftmp;
            // 迁移过程中的某一步
            double[] ktmp = new double[10];
            // 本次迁移过程的最优解（局部最优）
            double[] kpbest = new double[10];
            // 本次迁移过程中的最优值
            double fpbest;
            // 外层循环，依次遍历每个个体
            for (int i = 0; i < P; i++)
            {
                // 每一个个体迁移前初始化一些参数
                t = Step;
                kpbest = kstart[i];
                fpbest = f(kpbest, x, y);
                while (t <= PathLength)
                {
                    // 内层循环，个体中的每一个维都进行运算
                    for (int j = 0; j < D; j++)
                    {
                        // 个体迁移公式
                        ktmp[j] = kstart[i][j] + (k[j] - kstart[i][j]) * t * PRTVector();
                    }
                    // 计算本次迁移的函数值
                    ftmp = f(ktmp, x, y);
                    // 保存本次迁移的最优值
                    if (ftmp < fpbest)
                    {
                        fpbest = ftmp;
                        kpbest = ktmp;
                    }
                    t += Step;
                }
                // 完成这一个个体的迁移
                kstart[i] = kpbest;
                // 保存当前迁移个体的函数值
                fstart[i] = fpbest;
            }
            // 初始化当前最坏值，这个值在重新寻找后一定比fbest大
            fworst = fbest;
            // 种群的一次迁移全部完成，选出中群众的新Leader和新的最优值
            for (int i = 0; i < 20; i++)
            {
                if (fstart[i] < fbest)
                {
                    // 存储新的Leader
                    fbest = fstart[i];
                    k = kstart[i];
                }
                // 存储最坏值
                if (fstart[i] > fworst) fworst = fstart[i];
            }
            // 最坏和最好的差值小于指定精度，满足要求，返回1，说明已经找到
            if (Math.Abs(fworst - fbest) < MinDiv) return 1;
            return 0;
        }

        // 目标函数的计算
        private double f(double[] k, double[,] x, double[] y)
        {
            double target = 0;
            // J(k1,k2,...,kn)=sigma((y-xi*ki)^2)
            // 求所有三组数据中最优函数值的加和
            for (int j = 0; j < 3; j++)
            {
                for (int i = 0; i < D; i++)
                {
                    target += Math.Pow((y[j]-x[j,i]*k[i]),2);
                }
            }
            return target;
        }
    }
}
