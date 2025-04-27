using System;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Sw_toolkit
{
    /// <summary>
    /// 边界框尺寸
    /// </summary>
    class SW_BoundingSize
    {
        //计算包络尺寸子程序

        public static string Get_BoundingSize_value(SldWorks swApp)
        {
            ModelDoc2 swDoc = swApp.ActiveDoc;

            // 获取零件或装配体边界框信息
            double[] Corners = null;
            if (swDoc.GetType() == 1) { PartDoc Part = (PartDoc)swDoc; Corners = Part.GetPartBox(true); }
            if (swDoc.GetType() == 2) { AssemblyDoc Asm = (AssemblyDoc)swDoc; Corners = Asm.GetBox(0); }
            //if (swDoc.GetType() == 2) 
            //{ 
            //    AssemblyDoc Asm = (AssemblyDoc)swDoc;
            //    object [] objs= Asm.GetComponents(true);
            //    Component com = (Component)objs[0];
            //    Corners = com.GetBox(false,false); 
            //}
            if (swDoc.GetType() != 1 && swDoc.GetType() != 2) { MessageBox.Show("此功能只对零件或者装配体有效"); }
            //
            Console.WriteLine(Corners[0]);
            Console.WriteLine(Corners[1]);
            Console.WriteLine(Corners[2]);
            Console.WriteLine(Corners[3]);
            Console.WriteLine(Corners[4]);
            Console.WriteLine(Corners[5]);


            //获取用户单位设置
            short[] UserUnits = swDoc.GetUnits();
            //根据单位类型设置转换因子
            double ConvFactor = GetConvFactor(UserUnits[0]);

            //计算边界框的长度、宽度和高度
            object Length = (Corners[3] - Corners[0]) * ConvFactor;// X轴
            object Width = (Corners[5] - Corners[2]) * ConvFactor; // Y轴
            object Height = (Corners[4] - Corners[1]) * ConvFactor;// Z轴

            ////边界框添加补偿
            //double  AddFactor=0;  //输入补偿值   
            //Length = Math.Abs(((Corners[3] - Corners[0]) * ConvFactor) + AddFactor);// X轴
            //Width = Math.Abs(((Corners[5] - Corners[2]) * ConvFactor) + AddFactor); // Y轴
            //Height = Math.Abs(((Corners[4] - Corners[1]) * ConvFactor) + AddFactor);// Z轴

            //    // 如果单位是英尺英寸，并且为分数形式，进行特殊处理
            //    If (UserUnits(0) = swFEETINCHES Or UserUnits(0) = swINCHES) And UserUnits(1) = 2 Then
            //        Height = DecimalToFeetInches(Height, CInt(UserUnits(2)))
            //        Width = DecimalToFeetInches(Width, CInt(UserUnits(2)))
            //        Length = DecimalToFeetInches(Length, CInt(UserUnits(2)))
            //    End If

            string s = Length + "×" + Width + "×" + Height; Console.WriteLine(s);
            return s;
        }

        //获取单位转换因子函数
        private static Double GetConvFactor(int UnitType)
        {
            double GetConvFactor = 1;
            switch (UnitType)
            {
                case 0: GetConvFactor = 1000; break;      // 毫米
                case 1: GetConvFactor = 100; break;   // 厘米
                case 2: GetConvFactor = 1; break;   // 米
                case 3: GetConvFactor = 39.37007874; break;  // 英寸
                case 4: GetConvFactor = 3.280839895; break;   // 英尺
                case 5: GetConvFactor = 39.37007874; break;   // 英尺英寸
                case 6: GetConvFactor = 10000000000; break;  // 埃
                case 7: GetConvFactor = 1000000000; break;    // 纳米
                case 8: GetConvFactor = 1000000; break;  // 微米
                case 9: GetConvFactor = 39370.07874; break;    // 密尔
                case 10: GetConvFactor = 39370078.74; break; // 微英寸}
            }
            return GetConvFactor;

        }
    }
}
