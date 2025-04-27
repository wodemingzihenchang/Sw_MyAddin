using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Windows.Forms;

namespace Sw_toolkit
{
    class Selcet_SamecolorFace
    {
        public static void Function(SldWorks swApp)
        {
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swApp.GetDocumentCount() == 0 || swModel.GetType() != 1) { MessageBox.Show("请打开零件"); return; }

            SelectionMgr swSelMgr = swModel.ISelectionManager;
            Face2 swFace = (Face2)swSelMgr.GetSelectedObject(1);
            if (swSelMgr.GetSelectedObjectCount() == 1)
            {
                if (swSelMgr.GetSelectedObjectType2(1) == 2)
                {
                    //获得所选面的颜色信息
                    if (swFace != null) GetFaceColor(swFace);
                    else { Console.WriteLine("swFace没有获得对象"); return; }
                    //比较颜色获得相同颜色的面
                    swModel.EditRebuild3();
                    SameFaceColor(swApp);
                    //获得相同颜色实体面
                    //double[] body_Values = Body2.GetMaterialPropertyValues
                    //Body2.GetFaces
                    //获得相同颜色特征面
                    //double[] fea_Values = Feature.GetMaterialPropertyValues
                    // Feature.GetFaces
                    //测量
                    Measure swMeasure = (Measure)swModel.Extension.CreateMeasure();
                    swMeasure.ArcOption = 0;
                    bool status = swMeasure.Calculate(null);
                    if ((status)) { Console.WriteLine("Total area: " + swMeasure.TotalArea); }
                    //写入属性
                    string porp_name = "所选面的表面积";
                    string porp_value1 = (swMeasure.Area * 1000000).ToString();
                    string porp_value2 = (swMeasure.TotalArea * 1000000).ToString();
                    if (porp_value1 != "-1000000") { swModel.Extension.CustomPropertyManager[""].Add3(porp_name, (int)swCustomInfoType_e.swCustomInfoText, porp_value1, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd); }
                    if (porp_value2 != "-1000000") { swModel.Extension.CustomPropertyManager[""].Add3(porp_name, (int)swCustomInfoType_e.swCustomInfoText, porp_value2, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd); }
                }
                else { MessageBox.Show("所选中的对象不是面"); return; }
            }
            else MessageBox.Show("请选中零件表面的1个面");
        }
        static double refR;
        static double refG;
        static double refB;
        private static void GetFaceColor(Face2 swFace)//获得所选面的颜色信息
        {
            double[] vMatValues = (double[])swFace.MaterialPropertyValues;
            if (vMatValues != null)
            {
                refR = vMatValues[0] * 255;
                refG = vMatValues[1] * 255;
                refB = vMatValues[2] * 255;
            }
            else { refR = 0; refG = 0; refB = 0; }
        }
        private static void SameFaceColor(SldWorks swApp)//遍历并对比所选面的颜色信息
        {
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            PartDoc swPrt = (PartDoc)swModel;
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            double[] vMatValues; double[] body_Values;
            double myR; double myG; double myB;

            //遍历实体
            object[] Bodies = (object[])swPrt.GetBodies2(-1, false);
            for (int i = 0; i < Bodies.Length; i++)
            {
                Body2 swBody = (Body2)Bodies[i];
                body_Values = (double[])swBody.MaterialPropertyValues2;

                ////获得相同颜色实体面
                //if (body_Values != null)
                //{
                //    myR = body_Values[0] * 255; myG = body_Values[1] * 255; myB = body_Values[2] * 255;
                //}
                //else { myR = 0; myG = 0; myB = 0; }
                //if (myR == refR && myG == refG && myB == refB)
                //{
                //    object[] body_Faces = (object[])swBody.GetFaces();
                //    for (int j = 0; j < body_Faces.Length; j++)
                //    {
                //        Face swFace = (Face)body_Faces[j];
                //        swSelMgr.AddSelectionListObject(swFace, swSelMgr.CreateSelectData());
                //    }
                //}


                //遍历面对象
                object[] Faces = (object[])swBody.GetFaces();
                for (int j = 0; j < Faces.Length; j++)
                {
                    Face swFace = (Face)Faces[j];
                    vMatValues = (double[])swFace.MaterialPropertyValues;
                    if (vMatValues != null)
                    {
                        myR = vMatValues[0] * 255;
                        myG = vMatValues[1] * 255;
                        myB = vMatValues[2] * 255;
                    }
                    else { myR = 0; myG = 0; myB = 0; }
                    if (myR == refR && myG == refG && myB == refB)
                    {
                        //Console.WriteLine("1");//swFace.Select(true); 有问题
                        swSelMgr.AddSelectionListObject(swFace, swSelMgr.CreateSelectData());
                    }
                }
            }
        }
    }
}
