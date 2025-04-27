using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Drawing;

namespace Sw_MyAddin.SLDDRW
{
    class SW_Layer
    {
        private static ModelDoc2 swDoc;
        private static DrawingDoc swDraw;

        /// <summary>
        /// 工程图线条改颜色分层
        /// </summary>
        public static void Layered(ISldWorks swApp)
        {
            swDoc = swApp.ActiveDoc;
            try { swDraw = (DrawingDoc)swDoc; }
            catch (Exception) { System.Windows.Forms.MessageBox.Show("请先打开工程图"); return; }

            //获取当前工程图对象
            Sheet drwSheet = (Sheet)swDraw.GetCurrentSheet();
            object[] views = (object[])drwSheet.GetViews();
            View view = (View)views[0];
            //获取当前工程图总装配体对象
            DrawingComponent swDrawComp0 = view.RootDrawingComponent;
            //获取当前工程图子装配体对象
            object[] childrencomps = (object[])swDrawComp0.GetChildren();
            //进度条
            进度条 form1 = new 进度条(); form1.Show();
            form1.progressBar1.Value = 0; ;
            form1.progressBar1.Maximum = childrencomps.Length;
            for (int i = 0; i < childrencomps.Length; i++)
            {
                //遍历工程图零部件
                DrawingComponent swDrawComp = (DrawingComponent)childrencomps[i];
                Component2 swComp = (Component2)swDrawComp.Component;
                //统一同名零件
                string samnename = swComp.Name.Substring(0, swComp.Name.LastIndexOf('-'));
                //新建图层
                NewLayer(samnename);
                //选择路径
                string selectname = swDrawComp0.Name + "@" + view.Name + "/" + swComp.Name;
                //设置图层
                SetLayer(selectname, samnename);
                //进度条
                form1.progressBar1.Value += 1;
            }
            form1.progressBar1.Value = form1.progressBar1.Maximum; form1.Close();
        }
        private static void NewLayer(string Layername)//新建图层
        {
            int[] num = File_edit.StrToint(Layername);
            //颜色定义方法,ToArgb()方法转成32进制
            Color color = Color.FromArgb(10, num[0], num[1], num[2]);
            int colorInt = color.ToArgb();
            //删除图层，以便重新新建图层（不然设置不会变）
            //LayerMgr layerMgr = (LayerMgr)swDoc.GetLayerManager();
            //layerMgr.DeleteLayer(Layername);
            //新建图层(图名，说明，颜色，线型，线粗，可见，可打印)
            swDraw.CreateLayer2(Layername, "说明", (int)colorInt, (int)swLineStyles_e.swLineCONTINUOUS, (int)swLineWeights_e.swLW_NORMAL, true, true);
        }
        private static void SetLayer(string Compname, string Layername) //改变零部件线型，Compname选择零部件路径，Layername添加图层名字
        {
            //获得所选择的对象
            swDoc.Extension.SelectByID2(Compname, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            SelectionMgr swSelMgr = (SelectionMgr)swDoc.SelectionManager;
            DrawingComponent swDrawComp = (DrawingComponent)swSelMgr.GetSelectedObjectsComponent4(1, 0);
            //关闭默认文档线型
            swDrawComp.UseDocumentDefaults = false;
            //线型
            swDrawComp.SetLineStyle((int)swDrawingComponentLineFontOption_e.swDrawingComponentLineFontVisible, (int)swLineStyles_e.swLineCONTINUOUS);
            //线粗
            swDrawComp.SetLineThickness((int)swDrawingComponentLineFontOption_e.swDrawingComponentLineFontVisible, (int)swLineWeights_e.swLW_CUSTOM, 0.0003);
            //图层
            swDrawComp.Layer = Layername; swDraw.ChangeComponentLayer(Layername, true); //（图层名称，是否应用所有视图）
        }
        private static void GetLayer()//获得图层属性
        {
            //获取当前图层对象
            var swLayerMgr = (LayerMgr)swDoc.GetLayerManager();
            var layCount = swLayerMgr.GetCount();
            String[] layerList = (String[])swLayerMgr.GetLayerList();

            foreach (var lay in layerList) //遍历图层
            {
                //图层类型赋值，用于后续获取图层属性
                Layer currentLayer = (Layer)swLayerMgr.GetLayer(lay);
                if (currentLayer != null)
                {
                    var currentName = currentLayer.Name;//图名
                    var currentColor = currentLayer.Color;//颜色
                    var currentDesc = currentLayer.Description;//说明
                    var currentStype = Enum.GetName(typeof(swLineStyles_e), currentLayer.Style);//线
                    var currentWidth = currentLayer.Width;//线粗

                    #region //颜色是Ref值结果转int，得到对应的RGB值
                    int refcolor = currentColor;
                    int blue = refcolor >> 16 & 255;
                    int green = refcolor >> 8 & 255;
                    int red = refcolor & 255;
                    int colorARGB = 255 << 24 | (int)red << 16 | (int)green << 8 | (int)blue;
                    Color ARGB = Color.FromArgb(colorARGB);
                    #endregion

                    Console.WriteLine($"图层名称：{currentName}");
                    Console.WriteLine($"图层颜色：R {ARGB.R},G {ARGB.G} ,B {ARGB.B}");
                    Console.WriteLine($"图层描述：{currentDesc}");
                    Console.WriteLine($"图层线型：{currentStype}");
                    Console.WriteLine($"-------------------------------------");
                }
            }
        }

        /*public static void s()
        {
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            DrawingDoc swDraw = (DrawingDoc)swModel;
            Sheet drwSheet = (Sheet)swDraw.GetCurrentSheet();       //获取当前工程图对象
            object[] views = (object[])drwSheet.GetViews();
            View view = (View)views[0];

            DrawingComponent swDrawComp0 = view.RootDrawingComponent; //获取当前工程图总装配体对象
            object[] childrencomps = (object[])swDrawComp0.GetChildren();//获取当前工程图子装配体对象

            for (int i = 0; i < childrencomps.Length; i++)//遍历工程图零部件
            {
                DrawingComponent swDrawComp = (DrawingComponent)childrencomps[i];
                Component2 swComp = (Component2)swDrawComp.Component;


                string samnename = swComp.Name.Substring(0, swComp.Name.LastIndexOf('-'));//统一同名零件
                                                                                          //新建图层
                SwModelDoc.NewLayer(samnename);
                //选择路径
                string selectname = swDrawComp0.Name + "@" + view.Name + "/" + swComp.Name;
                //设置图层
                SwModelDoc.SetLayer(selectname, samnename);
            }
            // iSwApp.SendMsgToUser(view.Name);
            //SwDrawing.SetLayer(); 
        }
        */
    }
}
