using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace Sw_MyAddin
{
    class SwModelDoc
    {
        //封拆箱控制swApp的读写性，获取当前运行的程序   
        public static SldWorks swApp { get; private set; }
        public static ISldWorks ConnectToSolidWorks()
        {

            string str1 = "SldWorks.Application";
            string str2 = "SldWorks.Application";
            if (swApp != null) { return swApp; }
            else
            {
                for (int i = 20; i < 100; i++)
                {
                    str2 = str1 + "." + i;
                    //利用try - catch检查程序对象和SW版本,只装一个版本的可以直接GetActiveObject("SldWorks.Application")  
                    try { swApp = (SldWorks)Marshal.GetActiveObject(str2); return swApp; }
                    catch (COMException) { swApp = null; }
                }
                MessageBox.Show(" 请先打开SOLIDWORKS程序 ");
                swApp.CommandInProgress = true;//告诉SW现在是用外部程序调用命令（优化性能）
                return swApp;
            }
        }
        //通过文件路径打开SW文件（1零件，2装配体，3工程图）
        public static ModelDoc2 swDoc = ConnectToSolidWorks().ActiveDoc;
        public static ModelDoc2 OpenDoc(string str)
        {
            string postfix = str.Substring(str.LastIndexOf('.'));

            if (postfix == ".sldprt" || postfix == ".SLDPRT")
            {
                swDoc = SwModelDoc.swApp.OpenDoc(str, 1);
            }
            if (postfix == ".sldasm" || postfix == ".SLDASM")
            {
                swDoc = SwModelDoc.swApp.OpenDoc(str, 2);
            }
            if (postfix == ".slddrw" || postfix == ".SLDDRW")
            {
                swDoc = SwModelDoc.swApp.OpenDoc(str, 3);
            }
            return swDoc;
        }

        public static void Function(SldWorks swApp)//自定义属性同步到所有配置特定属性
        {
            //删除配合
            ModelDoc2 swModel = swApp.ActiveDoc;

            //获得特征,//找到配合特征
            Feature swFeat = (Feature)swModel.FirstFeature(); Feature swMateFeat = null;
            while ((swFeat != null))
            {
                if ("MateGroup" == swFeat.GetTypeName()) { swMateFeat = (Feature)swFeat; break; }
                swFeat = (Feature)swFeat.GetNextFeature();
            }

            //获得配合的第一子特征
            Feature swSubFeat = (Feature)swMateFeat.GetFirstSubFeature();
            while ((swSubFeat != null))
            {
                //Debug.Print("    " + swSubFeat.Name);
                swModel.Extension.SelectByID2(swSubFeat.Name, "MATE", 0, 0, 0, true, 0, null, 0);
                //获得配合的下一子特征
                swSubFeat = (Feature)swSubFeat.GetNextSubFeature();
            }
            swModel.EditDelete();

            ////固定零部件
            swModel.Extension.SelectAll();
            AssemblyDoc assembly = (AssemblyDoc)swModel;
            assembly.FixComponent();

            swModel.Save();
        }





        public static void ASM_Function(ModelDoc2 swDoc)//功能测试
        {
            // 获取当前活动的文档
            AssemblyDoc swAssy = (AssemblyDoc)swApp.ActiveDoc;

            // 获取零部件，true仅获取顶层组件，false是全部，//遍历部件
            object[] vComponents = swAssy.GetComponents(false);
            foreach (Component2 SingleComponent in vComponents)
            {
                //判断压缩
                if (SingleComponent.IsSuppressed())
                {
                    //打开同名零件
                    Console.WriteLine(SingleComponent.Name);
                    SingleComponent.GetImportedPath();

                    // SingleComponent.FindAttribute(AttributeDef,WhichOne);
                    //Console.WriteLine(SingleComponent.ComponentReference);
                    //SingleComponent.ComponentReference = @"C:\Users\LeonLin\Desktop\新建文件夹 (2)\2018\零件2018.SLDPRT";
                    //选择部件
                    //swDoc.Extension.SelectByID2(SingleComponent.Name, "COMPONENT", 0, 0, 0, false, 0, null, 0);
                    swDoc.EditUnsuppress2();
                    //定位新位置
                    swDoc.ClearSelection2(true);
                    swDoc.ClearSelection2(true);
                }
            }
        }


        public static void Select_Function(ModelDoc2 swDoc)//功能测试
        {
            // 获取当前活动的文档
            AssemblyDoc swAssy = (AssemblyDoc)swApp.ActiveDoc;
            ModelDoc2 swModel = default(ModelDoc2);
            SelectionMgr swSelMgr = default(SelectionMgr);
            Component2 swComp = default(Component2);

            swModel = (ModelDoc2)swApp.ActiveDoc;
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swComp = (Component2)swSelMgr.GetSelectedObjectsComponent3(1, 0);
            if (swComp == null)
            {
                Debug.Print("Select a component and run the macro again.");
                return;

            }
            else
            {
                // swUserPreferenceToggle_e.swExtRefUpdateCompNames must be set to
                // false to change the name of a component using IComponent2::Name2
                swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swExtRefUpdateCompNames, false);

                // Print original name of component
                Debug.Print("  Original name of component = " + swComp.Name2);

                // Change name of component
                swComp.Name2 = "SW";

                // Print new name of component
                Debug.Print("  New name of component      = " + swComp.Name2);
            }
        }


        public static void Getall_Dimension(ModelDoc2 swModel)//获得零件全部尺寸
        {
            Feature feature;
            DisplayDimension displayDimension;
            Dimension dimension;

            swModel = swApp.ActiveDoc; if (swModel.GetType() != 1) { Console.WriteLine("仅支持SW零件格式"); }
            feature = swModel.FirstFeature();

            while (feature != null)
            {
                displayDimension = feature.GetFirstDisplayDimension();
                while (displayDimension != null)
                {
                    dimension = displayDimension.GetDimension();
                    Console.WriteLine(dimension.Value);
                    //dimension.GetSystemValue2("");
                    displayDimension = feature.GetNextDisplayDimension(displayDimension);
                }
                feature = feature.GetNextFeature();
            }
        }
        public static void Document_prop(SldWorks swApp)//查看文档的版本
        {
            ModelDoc2 swDoc = swApp.ActiveDoc;
            swDoc.VersionHistory();
        }
        public static void Option_config()//零件配置选项设置
        {

            ISldWorks swApp = SwModelDoc.ConnectToSolidWorks(); if (swApp == null) { return; }
            //  if (textBox1.Text == "请选择文件" || textBox1.Text == "" || textBox2.Text == "请选择文件" || textBox2.Text == "") { MessageBox.Show(@"请选择文件"); return; }

            ModelDoc2 swDoc = ((ModelDoc2)(swApp.ActiveDoc));
            ModelView myModelView = null;
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.FrameState = ((int)(swWindowState_e.swWindowMaximized));
            swDoc.Extension.SelectByID2("默认", "CONFIGURATIONS", 0, 0, 0, false, 0, null, 0);
            swDoc.EditConfiguration3("默认", "默认", " ", " ", 9181);

            IConfiguration config = swDoc.GetConfigurationByName("默认");

            config.ChildComponentDisplayInBOM = 1;
        }


    }

}

#region
/*
//——————工程图——————//
public static DrawingDoc swDraw = ConnectToSolidWorks().ActiveDoc;
/// <summary>
/// 获得工程图对象（图纸、视图、零部件）
/// </summary>
public static void GetSheetNames()
{
    Sheet drwSheet = (Sheet)swDraw.GetCurrentSheet();       //获取当前工程图对象

    object[] sheetNames = (object[])swDraw.GetSheetNames(); //获取当前工程图中的所有图纸名称

    object[] views = (object[])drwSheet.GetViews();         //获取当前工程图中的所有图纸视图

    foreach (object sheetName in sheetNames)
    {
        Debug.Print((String)sheetName);
    }

    foreach (View view in views)//遍历工程图零部件,输入选择视图,输出零部件名
    {
        //选择视图激活
        DrawingComponent comp = view.RootDrawingComponent; Debug.Print(comp.Name);
        //获得子件对象
        object[] childrencomps = (object[])comp.GetChildren();
        //遍历工程图零部件
        for (int i = childrencomps.GetLowerBound(0); i <= childrencomps.GetUpperBound(0); i++)
        {
            swDoc.ClearSelection2(true);
            Debug.Print("零部件是" + ((DrawingComponent)childrencomps[i]).Name);
        }
    }

}
//获得可见实体对象
public static void GetVisibleEntity()
{
    View swView = (View)((SelectionMgr)swDoc.SelectionManager).GetSelectedObject6(1, -1);

    DrawingComponent swDrawingComponent = swView.RootDrawingComponent;

    Component2 component = swDrawingComponent.Component; //如果是零件则为空，装配体看零部件

    Debug.Print(swDrawingComponent.Name);
    Debug.Print("Number of edges found: " + swView.GetVisibleEntityCount2(null, 1));
    Debug.Print("Number of Vertex found: " + swView.GetVisibleEntityCount2(null, 2));
    Debug.Print("Number of Face found: " + swView.GetVisibleEntityCount2(null, 3));
    Debug.Print("Number of SilhouetteEdge found: " + swView.GetVisibleEntityCount2(null, 4));
}
//表格对象
public static void GetTable()
{
    Feature f = (Feature)swDraw.FeatureByName("材料明细表1");
    BomFeature Bom = (BomFeature)f.GetSpecificFeature();
    Debug.Print(Bom.Name);
    object[] swBomAnn = (object[])Bom.GetTableAnnotations();

    //获取一般表特性的注释数据
    TableAnnotation swTableAnnotation = (TableAnnotation)swBomAnn[0];
    bool anchorAttached = swTableAnnotation.Anchored;
    Debug.Print("Table anchored        = " + anchorAttached);
    int anchorType = swTableAnnotation.AnchorType;
    Debug.Print("Anchor type           = " + anchorType);
    int nbrColumns = swTableAnnotation.ColumnCount;
    Debug.Print("Number of columns     = " + nbrColumns);
    int nbrRows = swTableAnnotation.RowCount;
    Debug.Print("Number of rows        = " + nbrRows);

}
*/

#endregion
