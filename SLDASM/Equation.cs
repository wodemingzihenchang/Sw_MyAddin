using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Sw_MyAddin.SLDASM
{
    /// <summary>
    /// 方程式
    /// </summary>
    class Equation
    {
        public static void Function(ISldWorks swApp)
        {
            ModelDoc2 swModel = swApp.ActiveDoc;
            if (swModel.GetType() != 2) { MessageBox.Show("当前文档不是装配体"); return; }
            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            Add_codename(swAssy, 0);
            MessageBox.Show("遍历完成");
        }
        static void Add(SldWorks swApp)//修改方程式
        {
            ModelDoc2 swDoc = swApp.ActiveDoc; Console.WriteLine(swDoc.GetPathName());
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
                //equationMgr.Delete(7);
                //equationMgr.Add(0, "\"Rough_Weight\" =\"展开长度\"*\"展开宽度\"*\"展开厚度\"*\"SW-密度\"");
                equationMgr.Add(1, "\"A\" = \"代号代码\"");
                equationMgr.Add(2, "\"B\" = \"名称代码\"");

                //特征
                //equationMgr.Add(-1, "\"切除 - 拉伸1\" = \"suppressed\"");
                //方程式
                //equationMgr.Add(-1, "\"D1@Sketch1\" = 0.05in+0.5in");
            }
            //方程式输入TXT文件
            //equationMgr.FilePath = @"E:\20-LeonLin\[06]Learn\#测试模型\2022\其他模型\equations.txt";
            //equationMgr.LinkToFile = false;
            //equationMgr.Add(-1, "图号代码");

        }
        static void Add_codename(AssemblyDoc swAssembly, int level)
        {
            string code = $"swModel.Extension.CustomPropertyManager(\"\").Set(\"代号\", Left(Part.GetTitle, InStr(Part.GetTitle, \" \")))";
            string name = $"swModel.Extension.CustomPropertyManager(\"\").Set(\"名称\", Left(Right(Part.GetTitle, Len(Part.GetTitle) - InStr(Part.GetTitle, \" \")), Len(Right(Part.GetTitle, Len(Part.GetTitle) - InStr(Part.GetTitle, \" \"))) - 7))";
            //进度条
            进度条 form_asmcount = new 进度条();
            form_asmcount.文字显示("装配体零部件数量提取中......"); form_asmcount.Show();
            //遍历零部件
            List<string> strList = new List<string>();
            object[] components = swAssembly.GetComponents(false); form_asmcount.Close();
            //进度条
            进度条 form1 = new 进度条(); form1.Text = "零件写入属性和方程式"; form1.Show();
            form1.progressBar1.Value = 0; ;
            form1.progressBar1.Maximum = components.Length;

            if (components == null) { return; }
            for (int i = 0; i < components.Length; i++)
            {
                Component2 swComponent = (Component2)components[i];
                ModelDoc2 swModel = swComponent.GetModelDoc2();

                //判断是否存在
                string SaveAs_path = Path.GetFileNameWithoutExtension(swComponent.GetPathName());
                if (!strList.Contains(SaveAs_path))
                {
                    strList.Add(SaveAs_path);
                    if ((int)swModel.GetType() == (int)1)//处理零件
                    {
                        swModel.AddCustomInfo3("", "代号", 1, "");
                        swModel.AddCustomInfo3("", "名称", 1, "");
                        CustomPropertyManager swCustPropMgr = swModel.Extension.CustomPropertyManager[""];//自定义属性  //CustomPropertyManager swCustPropMgr = swModel.Extension.CustomPropertyManager[swModel.GetActiveConfiguration().Name];//配置特定属性
                        swCustPropMgr.Add3("代号代码", 30, code, 1);
                        swCustPropMgr.Add3("名称代码", 30, name, 1);

                        string equationStr;
                        EquationMgr swEquationMgr = swModel.GetEquationMgr();
                        equationStr = $"\"1\" = \"代号代码\"";
                        swEquationMgr.Add(1, equationStr);
                        equationStr = $"\"2\" = \"名称代码\"";
                        swEquationMgr.Add(2, equationStr);
                        swModel.ForceRebuild3(true);
                    }
                    //进度条
                    form1.progressBar1.Value += 1; i += 1;
                    form1.数字显示(i, components.Length);
                }
                #region////递归处理子装配体里的零件
                //if (swComponent.GetSuppression() != 0)
                //{
                //    ModelDoc2 swSubAssembly = swComponent.GetModelDoc2();
                //    if (swSubAssembly != null)
                //    {
                //        if ((int)swSubAssembly.GetType() == (int)2)
                //        {
                //            Add_codename((AssemblyDoc)swSubAssembly, level + 1);
                //        }
                //    }
                //}
                #endregion
            }
            form1.Close();
            swAssembly.EditRebuild();
        }
    }
}
