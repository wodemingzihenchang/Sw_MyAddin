using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Sw_toolkit
{ 
    /// <summary>
    /// 方程式
    /// </summary>
    class Equation
    {
       
        public static void Add(ModelDoc2 swDoc)//修改方程式
        {
            //方程式管理器对象
            EquationMgr equationMgr = swDoc.GetEquationMgr();
            if (equationMgr != null)
            {
                equationMgr = swDoc.GetEquationMgr();
                //全局变量
                //equationMgr.Delete(0);
                //equationMgr.Delete(1);
                //equationMgr.Delete(2);
                //equationMgr.Delete(3);
                //equationMgr.Delete(4);
                //equationMgr.Delete(5);
                //equationMgr.Delete(6);
                ////equationMgr.Delete(7);
                equationMgr.Add(0, "\"Rough_Weight\" =\"展开长度\"*\"展开宽度\"*\"展开厚度\"*\"SW-密度\"");
                equationMgr.Add(1, "\"A\" = \"图号代码\"");
                equationMgr.Add(2, "\"B\" = \"名称代码\"");

                //特征
                //equationMgr.Add(-1, "\"切除 - 拉伸1\" = \"suppressed\"");
                //方程式
                //equationMgr.Add(-1, "\"D1@草图1\" = 0.05in+0.5in");
            }
            //方程式输入TXT文件
            //equationMgr.FilePath = @"E:\20-LeonLin\[06]Learn\#测试模型\2022\其他模型\equations.txt";
            //equationMgr.LinkToFile = false;
            //equationMgr.Add(-1, "图号代码");

        }
    }
}
