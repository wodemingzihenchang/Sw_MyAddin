using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace Sw_toolkit.Part
{
    class SW__SheetMetal
    {
        static SldWorks swApp;

        //——————钣金——————//       
        public static void SheetMetalFeatureData(SldWorks swApp)
        {
            //bool UseGaugeTable = true;
            //string GaugeTablePath = @"D:\Program Files\SolidWorks 2022\SOLIDWORKS\lang\chinese-simplified\Sheet Metal Gauge Tables\JHL-钢材规格表-按厚度排序.XLS";

            ModelDoc2 swDoc = swApp.ActiveDoc;
            //进入钣金特征
            FeatureManager swFeatMgr = swDoc.FeatureManager;
            SheetMetalFolder sheetMetalFolder = swFeatMgr.GetSheetMetalFolder();//钣金参数文件夹
            Feature feature = sheetMetalFolder.GetFeature();
            SheetMetalFeatureData sheetMetalFeatureData = feature.GetDefinition();

            //进入访问钣金数据
            //sheetMetalFeatureData.AccessSelections(swDoc, null);
            //折弯系数
            CustomBendAllowance customBendAllowance= sheetMetalFeatureData.GetCustomBendAllowance();
            Console.WriteLine(customBendAllowance.BendAllowance*1000);

            //属性 
            //bool UseGaugeTable;
            //sheetMetalFeatureData.GetOverrideDefaultParameter(out UseGaugeTable);
            //Console.WriteLine(sheetMetalFeatureData.AutoReliefType);    //获取或设置钣金折弯止裂槽类型。
            //Console.WriteLine(sheetMetalFeatureData.BendAllowance);     //Obsolete. Superseded by SetCustomBendAllowance
            //Console.WriteLine(sheetMetalFeatureData.BendAllowanceType); //Obsolete. Superseded by SetCustomBendAllowance
            //Console.WriteLine(sheetMetalFeatureData.BendRadius);        //转弯半径
            //Console.WriteLine(sheetMetalFeatureData.BendTableFile);     //Obsolete. Superseded by SetCustomBendAllowance
            //Console.WriteLine(sheetMetalFeatureData.FixedReference);  //获取或设置此钣金特征的固定面或边。
            //sheetMetalFeatureData.KFactor = 5;
            Console.WriteLine(sheetMetalFeatureData.KFactor * 1000);           //K因子，Obsolete. Superseded by SetCustomBendAllowance
            //Console.WriteLine(sheetMetalFeatureData.ReliefRatio = 0.005);       //获取或设置此钣金特征的起伏比。
            //Console.WriteLine(sheetMetalFeatureData.Thickness = 0.003);         //厚度
            //Console.WriteLine(sheetMetalFeatureData.UseAutoRelief);     //使用自动起伏	获取或设置此钣金特征是否使用自动起伏。
            //Console.WriteLine(sheetMetalFeatureData.UseMaterialSheetMetalParameters = true);//获取或设置是否使用创建此钣金特征时应用的材料的属性。
            //方法
            //bool overParameter = true;
            //sheetMetalFeatureData.GetOverrideDefaultParameter(out overParameter);//获取是否在多实体钣金零件的此钣金特征中覆盖指定的缺省参数。
            //sheetMetalFeatureData.SetOverrideDefaultParameter( overParameter);//设置是否覆盖多实体钣金零件中此钣金特征中的指定缺省参数。
            //sheetMetalFeatureData.GetCustomBendAllowance();//获取此钣金特征的自定义折弯余量。
            //sheetMetalFeatureData.ReleaseSelectionAccess();//释放对用于定义钣金特征的选择的访问。
            //sheetMetalFeatureData.SetCustomBendAllowance();//设置此钣金特征的自定义折弯余量。
            //设置是否使用钣金特征仪表表。
            //sheetMetalFeatureData.GetUseGaugeTable(out UseGaugeTable, out GaugeTablePath);
            //sheetMetalFeatureData.SetUseGaugeTable(UseGaugeTable, GaugeTablePath);

            //完成修改
            feature.ModifyDefinition(sheetMetalFeatureData, swDoc, null);
        }
        public static void test()
        {
            //参考SOLIDWORKS API Help中的IExportToDWG2 Method (IPartDoc) 和 Export Part to DWG Example (C#)
            //首先在钣金中绘制三维草图，长边代表导出dxf图的X方向，短边代表dxf图的Y方向，用于限定钣金拉丝方向(默认是X方向)，防止排版错误

            ModelDoc2 swDoc = (ModelDoc2)swApp.ActiveDoc; //获取当前已打开的零件
            if (swDoc != null)
            {
                PartDoc swPart = (PartDoc)swDoc;
                string swDocName = swDoc.GetPathName();
                string swDxfName = swDocName.Substring(0, swDocName.Length - 6) + "dxf";
                double[] dataAlignment = new double[12];
                dataAlignment[0] = 0.0;
                dataAlignment[1] = 0.0;
                dataAlignment[2] = 0.0;
                dataAlignment[3] = 1.0;
                dataAlignment[4] = 0.0;
                dataAlignment[5] = 0.0;
                dataAlignment[6] = 0.0;
                dataAlignment[7] = 1.0;
                dataAlignment[8] = 0.0;
                dataAlignment[9] = 0.0;
                dataAlignment[10] = 0.0;
                dataAlignment[11] = 1.0;
                //Array[0], Array[1], Array[2] - XYZ coordinates of new origin
                //Array[3], Array[4], Array[5] - coordinates of new x direction vector
                //Array[6], Array[7], Array[8] - coordinates of new y direction vector
                //判断XYAXIS，长边作为X轴，短的作为Y轴，用于限定拉丝方向
                bool status = swDoc.Extension.SelectByID2("XYAXIS", "SKETCH", 0, 0, 0, false, 0, null, 0);
                if (status)
                {
                    SelectionMgr swSelectionMgr = swDoc.SelectionManager;
                    Feature swFeature = swSelectionMgr.GetSelectedObject6(1, -1);
                    Sketch swSketch = swFeature.GetSpecificFeature2();
                    var swSketchPoints = swSketch.GetSketchPoints2();//获取草图中的所有点
                                                                     //用这三个点抓取直线，并判断长度，长边作为X轴，画3D草图的时候一次性画出两条线，不能分两次画出，否则会判断错误
                    SketchPoint p0 = swSketchPoints[0];//最先画的点
                    SketchPoint p1 = swSketchPoints[1];//作为坐标原点
                    SketchPoint p2 = swSketchPoints[2];//最后画的点
                    dataAlignment[0] = p1.X * 1000;
                    dataAlignment[1] = p1.Y * 1000;
                    dataAlignment[2] = p1.X * 1000;
                    double l1 = Math.Sqrt(Math.Pow(p0.X - p1.X, 2) + Math.Pow(p0.Y - p1.Y, 2) + Math.Pow(p0.Z - p1.Z, 2));
                    double l2 = Math.Sqrt(Math.Pow(p2.X - p1.X, 2) + Math.Pow(p2.Y - p1.Y, 2) + Math.Pow(p2.Z - p1.Z, 2));
                    if (l1 > l2)
                    {
                        dataAlignment[3] = p0.X * 1000 - p1.X * 1000;
                        dataAlignment[4] = p0.Y * 1000 - p1.Y * 1000;
                        dataAlignment[5] = p0.Z * 1000 - p1.Z * 1000;
                        dataAlignment[6] = p2.X * 1000 - p1.X * 1000;
                        dataAlignment[7] = p2.Y * 1000 - p1.Y * 1000;
                        dataAlignment[8] = p2.Z * 1000 - p1.Z * 1000;
                    }
                    else
                    {
                        dataAlignment[3] = p2.X * 1000 - p1.X * 1000;
                        dataAlignment[4] = p2.Y * 1000 - p1.Y * 1000;
                        dataAlignment[5] = p2.Z * 1000 - p1.Z * 1000;
                        dataAlignment[6] = p0.X * 1000 - p1.X * 1000;
                        dataAlignment[7] = p0.Y * 1000 - p1.Y * 1000;
                        dataAlignment[8] = p0.Z * 1000 - p1.Z * 1000;
                    }
                }
                object varAlignment = dataAlignment;

                //Export sheet metal to a single drawing file将钣金零件导出单个dxf文件
                //include flat-pattern geometry，倒数第二位数字1代表钣金展开，options = 1;
                swPart.ExportToDWG2(swDxfName, swDocName, (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal, true, varAlignment, false, false, 4095, null);
            }
        }
        public static void SetParameters()//更新参数
        {
            //bool UseGaugeTable = true;
            //string GaugeTablePath = @"D:\Program Files\SolidWorks 2022\SOLIDWORKS\lang\chinese-simplified\Sheet Metal Gauge Tables\JHL-钢材规格表-按厚度排序.XLS";

            ModelDoc2 swDoc = swApp.ActiveDoc;
            PartDoc part = (PartDoc)swDoc;
            //获得基体法兰对象
            Feature swFeat = null;
            for (int i = 1; i < 100; i++)
            {
                swFeat = part.FeatureByName("基体-法兰" + i); if (swFeat != null) { break; }
                swFeat = part.FeatureByName("Base Flange" + i); if (swFeat != null) { break; }
            }
            //编辑特征
            BaseFlangeFeatureData swBaseFlangeFeatData = swFeat.GetDefinition();

            Debug.WriteLine(swBaseFlangeFeatData.BendRadius);

            swBaseFlangeFeatData.AccessSelections(swDoc, null);
            //钣金规格
            //swBaseFlangeFeatData.GaugeTablePath = GaugeTablePath;
            //钣金参数
            if (swBaseFlangeFeatData.UseGaugeTable)//是否使用规格表
            {
                string[] vThicknesses = swBaseFlangeFeatData.GetTableThicknesses();
                swBaseFlangeFeatData.ThicknessTableName = vThicknesses[4];
            }
            //swBaseFlangeFeatData.OverrideRadius = true;
            //swBaseFlangeFeatData.Thickness = 0.006;                           //选择厚度
            //swBaseFlangeFeatData.BendRadius = 0.006;                          //选择钣金

            //完成修改
            swFeat.ModifyDefinition(swBaseFlangeFeatData, swDoc, null);
            swBaseFlangeFeatData.KFactor = 0.5;
            Debug.WriteLine(swBaseFlangeFeatData.KFactor);
            MessageBox.Show("完成");
        }
        private void Debug_Click2()//更新规格表
        {

            ModelDoc2 swDoc = swApp.ActiveDoc;
            PartDoc part = (PartDoc)swDoc;

            //删除旧规格表
            ModelDocExtension SwModelDocExt = swDoc.Extension;
            int DeleteOption = (int)swDeleteSelectionOptions_e.swDelete_Absorbed;
            SwModelDocExt.SelectByID2("规格表", "SM ENVIRONMENT TABLE", 0, 0, 0, false, 0, null, 0);
            bool longstatus = SwModelDocExt.DeleteSelection2(DeleteOption);

            //设置新规格表（这里使用材料钣金属性的规格表-是2019新功能）
            string standard = @"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2022\自定义材料\JHL-自定义材料.sldmat";
            part.SetMaterialPropertyName2("默认", standard, "2.0钢板"); part.EditRebuild();

            //设置规格参数
            Feature swFeat = null;
            for (int i = 1; i < 100; i++)
            {
                swFeat = part.FeatureByName("基体-法兰" + i); if (swFeat != null) { break; }
                swFeat = part.FeatureByName("Base Flange" + i); if (swFeat != null) { break; }
            }
            //编辑特征
            BaseFlangeFeatureData swBaseFlangeFeatData = swFeat.GetDefinition();
            swBaseFlangeFeatData.AccessSelections(swDoc, null);
            //钣金参数
            if (swBaseFlangeFeatData.UseGaugeTable)//是否使用规格表
            {
                string[] vThicknesses = swBaseFlangeFeatData.GetTableThicknesses();
                //swBaseFlangeFeatData.ThicknessTableName = vThicknesses[3];
                swBaseFlangeFeatData.ThicknessTableName = "2.0 不锈钢板|覆铝锌板|钢板|铝板|无磁不锈钢板";
            }
            swBaseFlangeFeatData.OverrideThickness = false;
            swBaseFlangeFeatData.OverrideRadius = true;
            //swBaseFlangeFeatData.Thickness = 0.006;                    
            swBaseFlangeFeatData.BendRadius = 0.0002;
            //完成修改
            swFeat.ModifyDefinition(swBaseFlangeFeatData, swDoc, null);

            //KillProgram("EXCEL");
            GC.Collect();

            // MessageBox.Show("完成"); this.Close();
        }


        private void KillProgram(string str)//关闭任务管理器程序
        {
            /*if (book != null)
                book = null;  // WorkBook 的实例欢畅
            if (app != null)
                app.Quit(); // Microsoft.Office.Interop.Excel  的实例对象 
            */
            GC.Collect();  // 回收资源
            System.Diagnostics.Process[] excelProcess = System.Diagnostics.Process.GetProcessesByName(str);//str为程序名，例如：Excel;
            foreach (var item in excelProcess)
            {
                item.Kill();
            }
        }
        public static void writeLogException(Exception ex) //异常信息写入日志
        {
            //获取异常信息的类、行号、异常 信息
            string exceptionStr =
            ex.StackTrace.ToString().Substring(ex.StackTrace.ToString().LastIndexOf('\\') + 1)
             + "  " + ex.Message;
            exceptionStr = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "  " + exceptionStr;
            //自己定义一个存储日志文件的位置
            string sFilePath = "C:\\Logs";
            string sFileName = DateTime.Now.ToString("yyyy-MM-dd") + ".log";
            sFileName = sFilePath + @"\\" + sFileName; //文件
            if (!Directory.Exists(sFilePath))
            {
                Directory.CreateDirectory(sFilePath);
            }
            FileStream fs;
            StreamWriter sw;
            if (System.IO.File.Exists(sFileName))
            {
                fs = new FileStream(sFileName, FileMode.Append, FileAccess.Write);
            }
            else
            {
                fs = new FileStream(sFileName, FileMode.Create, FileAccess.Write);
            }
            sw = new StreamWriter(fs);
            sw.WriteLine(exceptionStr);
            sw.Close();
            fs.Close();
        }

    }
}
