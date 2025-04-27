using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Windows.Forms;

namespace Sw_MyAddin
{
    /// <summary>
    /// 边界框
    /// </summary>
    class Bounding
    {
        public static void Get_Bounding(ISldWorks swApp)
        {
            ModelDoc2 swDoc = swApp.ActiveDoc;
            if (swDoc != null)
            {
                Feature swFeat = GetBoundingBoxFeature(swDoc);
                if (swFeat != null)
                {
                    Sketch swSketch = swFeat.GetSpecificFeature2();
                    object[] objs = swSketch.GetSketchSegments();
                    ConvertSegmentsIntoSketch(swDoc, objs);
                }
                else { MessageBox.Show("获得包围框失败"); }
            }
            else { MessageBox.Show("打开文件"); }
        }
        private static Feature GetBoundingBoxFeature(ModelDoc2 swDoc)
        {
            //检查是否有边界框
            Feature swFeat = FindBoundingBoxFeature(swDoc);
            if (swFeat == null)
            {//为空则生成边界框
                int status;
                swDoc.FeatureManager.InsertGlobalBoundingBox((int)swGlobalBoundingBoxFitOptions_e.swBoundingBoxType_BestFit, false, false, out status);
                swFeat = FindBoundingBoxFeature(swDoc);
            }
            return swFeat;

        }
        private static Feature FindBoundingBoxFeature(ModelDoc2 swDoc)//获得边界框特征
        {
            Feature swFeat = swDoc.FirstFeature();
            while (swFeat != null)
            {
                //Console.WriteLine(swFeat.Name);
                //Console.WriteLine(swFeat.GetTypeName2());
                if (swFeat.GetTypeName2() == "BoundingBoxProfileFeat") { return swFeat; }
                swFeat = swFeat.GetNextFeature();
            }
            return null;
        }
        private static void ConvertSegmentsIntoSketch(ModelDoc2 swDoc, object[] segs)//插入3D草图
        {
            if (swDoc.SketchManager.ActiveSketch == null) { swDoc.SketchManager.Insert3DSketch(true); }
            else if (swDoc.SketchManager.ActiveSketch.Is3D() == false) { MessageBox.Show("", "支持3D草图"); }

            swDoc.ClearSelection2(true);
            for (int i = 0; i < segs.Length; i++)
            {
                SketchSegment swSkSeg = (SketchSegment)segs[i]; swSkSeg.Select4(true, null);
            }
            swDoc.SketchManager.SketchUseEdge3(false, false);
            swDoc.SketchManager.Insert3DSketch(true);
        }
    }
}
