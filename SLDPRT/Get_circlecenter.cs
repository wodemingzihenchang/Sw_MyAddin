using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Sw_MyAddin.SLDPRT
{
    /// <summary>
    /// 绘制草图圆心点
    /// </summary>
    class Get_circlecenter
    {
        public static void circlecenter(SldWorks swApp)
        {
            var swModel = (ModelDoc2)swApp.ActiveDoc;
            var swSelMgr = (SelectionMgr)swModel.SelectionManager;

            //获取选择边
            var swEdge = (Edge)swSelMgr.GetSelectedObject6(1, -1);

            var swCurve = (Curve)swEdge.GetCurve();

            if (swCurve.IsCircle())
            {//判断是否是圆边
                var edgeParams = (double[])swCurve.CircleParams;
                double x = edgeParams[0];
                double y = edgeParams[1];
                double z = edgeParams[2];

                Console.WriteLine("X坐标是" + x + "、Y坐标是" + y + "、Z坐标是" + z);

                SketchPoint skPoint = null;
                skPoint = ((SketchPoint)(swModel.SketchManager.CreatePoint(x, y, z)));

            }
        }
        public static void GetCyclePoint(ModelDoc2 swModel)//绘制草图圆心
        {
            //获取选择边
            //ISketch sketch = (ISketch)swSelMgr.GetSelectedObject6(1, -1);
            //var CenterPoint = arc.GetCenterPoint();
            //Console.WriteLine(arc.GetCenterPoint()[0] + arc.GetCenterPoint()[1]);
            //激活草图
            SketchManager sketchMgr = swModel.SketchManager;
            Sketch sketch = sketchMgr.ActiveSketch;
            //获得草图圆弧包括：圆、圆弧、U型槽、
            var arcs = sketch.GetArcs();
            for (int i = 0; i < sketch.GetArcCount(); i++)
            {
                SketchPoint skPoint = ((SketchPoint)(swModel.SketchManager.CreatePoint(arcs[7 + 11 * i], arcs[8 + 11 * i], 0)));
            }
        }

    }
}
