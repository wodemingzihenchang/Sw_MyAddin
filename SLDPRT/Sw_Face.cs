using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Windows.Forms;
using System;

namespace Sw_toolkit
{
    class Sw_Face
    {
        private static SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();


        public static void Selcet_SamecolorFace()
        {
            ModelDoc2 swModel = swApp.ActiveDoc;
            PartDoc swPrt = (PartDoc)swModel;
            if (swApp.GetDocumentCount() == 0) { MessageBox.Show("请在零件状态下运行"); return; }
            if (swModel.GetType() != 1) { MessageBox.Show("仅支持零件状态下运行"); return; };

            SelectionMgr swSelMgr = swModel.ISelectionManager;
            Face2 swFace = swSelMgr.GetSelectedObject(1);
            if (swSelMgr.GetSelectedObjectCount() == 1)
            {
                if (swSelMgr.GetSelectedObjectType2(1) == 2)
                {
                    //获得所选面的颜色信息
                    if (swFace != null) GetFaceColor(swFace);
                    else { Console.WriteLine("swFace没有获得对象"); return; }
                    //比较颜色获得相同颜色的面
                    swModel.EditRebuild3();
                    SameFaceColor(swPrt);
                    //测量
                    Measure swMeasure = (Measure)swModel.Extension.CreateMeasure();
                    swMeasure.ArcOption = 0;
                    bool status = swMeasure.Calculate(null);
                    if ((status))
                    {
                        Console.WriteLine("Total area: " + swMeasure.TotalArea);
                    }
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
        private static void SameFaceColor(PartDoc swPrt)//遍历并对比所选面的颜色信息
        {
            ModelDoc2 swDoc = swApp.ActiveDoc; SelectionMgr swSelMgr = swDoc.SelectionManager;
            double[] vMatValues;
            double myR; double myG; double myB;

            //遍历实体
            object[] Bodies = swPrt.GetBodies2(-1, false);
            for (int i = 0; i < Bodies.Length; i++)
            {
                Body2 swBody = (Body2)Bodies[i];
                //遍历面对象
                object[] Faces = swBody.GetFaces();
                for (int j = 0; j < Faces.Length; j++)
                {
                    Face swFace = (Face)Faces[j];
                    vMatValues = swFace.MaterialPropertyValues;
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

        public static void GetFace()
        {
            ModelDoc2 swModel = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            FeatureManager swFeatureManager = default(FeatureManager);
            MidSurface3 swMidSurfaceFeature = default(MidSurface3);
            Feature swFeature = default(Feature);
            SelectionMgr swSelectionMgr = default(SelectionMgr);
            Face2 swFace = default(Face2);
            bool status = false;
            int errors = 0;
            int warnings = 0;
            string fileName = null;
            int count = 0;
            object[] faces = null;
            int i = 0;

            fileName = "C:\\Users\\Public\\Documents\\SOLIDWORKS\\SOLIDWORKS 2018\\samples\\tutorial\\api\\box.sldprt";
            swModel = (ModelDoc2)swApp.OpenDoc6(fileName, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

            swModelDocExt = (ModelDocExtension)swModel.Extension;
            status = swModelDocExt.SelectByID2("", "FACE", -0.0533080255641494, 0.0299999999999727, 0.0131069871973182, true, 0, null, 0);
            status = swModelDocExt.SelectByID2("", "FACE", -0.0370905424398416, 0, 0.0289438729892595, true, 0, null, 0);
            swFeatureManager = (FeatureManager)swModel.FeatureManager;
            swFeatureManager.InsertMidSurface(null, null, 0.0, false);
            status = swModelDocExt.SelectByID2("Surface-MidSurface1", "REFSURFACE", 0, 0, 0, false, 0, null, 0);
            swSelectionMgr = (SelectionMgr)swModel.SelectionManager;
            swFeature = (Feature)swSelectionMgr.GetSelectedObject6(1, -1);
            swMidSurfaceFeature = (MidSurface3)swFeature.GetSpecificFeature2();
            count = swMidSurfaceFeature.GetFaceCount();
            //Debug.Print("Number of faces for midsurface feature: " + count);
            faces = (object[])swMidSurfaceFeature.GetFaces();
            for (i = faces.GetLowerBound(0); i <= faces.GetUpperBound(0); i++)
            {
                swFace = (Face2)faces[i];
                //Debug.Print("Area of face " + i + " of midsurface feature: " + swFace.GetArea());
            }

        }
    }
}
