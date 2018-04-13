using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CCS_System
{
    public class Mineral
    {
        // 基础元素
        public double Cu = 0.0;
        public double Fe = 0.0;
        public double S = 0.0;
        public double SiO2 = 0.0;
        public double CaO = 0.0;
        public double MgO = 0.0;
        public double Al2O3 = 0.0;
        // 化合物
        public double comp_CuFeS2 = 0.0;
        public double comp_CuS = 0.0;
        public double comp_Cu2S = 0.0;
        public double comp_Cu5FeS4 = 0.0;
        public double comp_Cu4SO4_OH_6 = 0.0;
        public double comp_Cu2_OH_2CO3 = 0.0;
        public double comp_FeS2 = 0.0;
        public double comp_Fe2O3 = 0.0;
        public double comp_SiO2 = 0.0;
        public double comp_Mg6Si8O20_OH_4 = 0.0;
        public double comp_KAlSi3O8 = 0.0;
        public double comp_KAl2_AlSi3O10__OH_2 = 0.0;
        public double comp_CaMg_CO3_2 = 0.0;
        public double comp_C = 0.0;
        public double comp_Cu2O = 0.0;
        public double comp_Cu = 0.0;
        public double comp_Fe3O4 = 0.0;
        public double comp_Fe2SiO4 = 0.0;
        public double comp_Fe = 0.0;
        public double comp_CaO = 0.0;
        public double comp_Al2O3 = 0.0;
        public double comp_K2O = 0.0;
        public double comp_MgO = 0.0;
        public double comp_S2 = 0.0;
        // 石英砂专属变量
        public double QS_CaCO3 = 0.0;
        // 用量
        public double dosage = 0.0;

        public Mineral()
        {
        }

        // 初始化单质量
        public Mineral(double Cu, double Fe, double S, double SiO2, double CaO, double MgO, double Al2O3)
        {
            this.Cu = Cu;
            this.Fe = Fe;
            this.S = S;
            this.SiO2 = SiO2;
            this.CaO = CaO;
            this.MgO = MgO;
            this.Al2O3 = Al2O3;
        }
    }
}
