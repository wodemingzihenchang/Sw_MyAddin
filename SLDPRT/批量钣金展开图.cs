using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Windows.Forms;

namespace Sw_toolkit
{
    public partial class Sw_Export_DWG : Form
    {
        private static SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();

        public string[] filepath = null;           //
        int Addcount = 0;                          //用于记录多次导入选择文件的计数累计（没有这项，在导入时会覆盖之前的行数据）

        //————————————//
        public Sw_Export_DWG() { InitializeComponent(); }
        private void Export_Load(object sender, EventArgs e)//许可加载
        {
            //if (License.License_Time("2023-06-30") || License.License_Registry()) { return; }
            //MessageBox.Show("程序更新，请联系https://space.bilibili.com/12254884", "注意", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk); System.Environment.Exit(0);
            //识别电脑许可
            
            if (License.License_Machine(Program.promname, Program.license)) { return; }
            System.Environment.Exit(0);
        }
        private void button1_Click(object sender, EventArgs e)//选择文件
        {
            //选择文件
            filepath = File.Addfiles(true);
            //进度条
            progressBar1.Value = 0; label1.Text = "进度： ";
            progressBar1.Maximum = filepath.Length;
            //设置表格行列
            dataGridView1.RowCount = filepath.Length + Addcount + 1;
            dataGridView1.ColumnCount = 2;
            dataGridView1.Columns[0].HeaderText = "文件位置"; int i;

            for (i = 0; i < filepath.Length; i++)
            {//写入表格内容
                this.dataGridView1.Rows[i + Addcount].Cells[0].Value = filepath[i];
                //进度条
                progressBar1.Value += 1;
            }
            this.dataGridView1.Rows[i + Addcount].Cells[0].Value = "/";
            Addcount += i;

            //判断是否选择文件，
            if (filepath.Length == 0) { return; }

            //进度条
            label1.Text = "进度：完成"; progressBar1.Value = progressBar1.Maximum;

        }
        private void button2_Click(object sender, EventArgs e)//运行操作
        {
            //进度条
            progressBar1.Value = 0; label1.Text = "进度： ";
            progressBar1.Maximum = dataGridView1.RowCount;

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                //打开文件
                ModelDoc2 swDoc = SwModelDoc.OpenDoc((string)dataGridView1.Rows[i].Cells[0].Value); swDoc.EditRebuild3();
                //操作内容
                ExportToDWG2(SwModelDoc.swDoc);
                //关闭文件
                SwModelDoc.swApp.CloseDoc((string)dataGridView1.Rows[i].Cells[0].Value);
                //进度条
                this.dataGridView1.Rows[i].Cells[1].Value = "完成";
                progressBar1.Value += 1;
            }
            //进度条
            label1.Text = "进度：完成"; progressBar1.Value = progressBar1.Maximum;
            MessageBox.Show(@"完成"); this.Close();
        }
        private void button3_Click(object sender, EventArgs e)//Debug
        {
            //SwModelDoc.btnGetDimensionInfo_Click(swApp);
        }


        /*————  导出操作  ————*/
        public static void ExportToDWG2(ModelDoc2 swDoc)
        {
            //获取当前已打开的零件
            PartDoc swPart = (PartDoc)swDoc;
            //获取保存路径
            string swDocName = swDoc.GetPathName();
            string swDxfName = swDocName.Substring(0, swDocName.Length - 6) + "dxf";
            //double[] dataAlignment = new double[12]; object varAlignment = dataAlignment;
            swPart.ExportToDWG2(swDxfName, swDocName, (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal, true, null, false, false, 2149, null);
            #region
            /*
            FilePath
            导出的 DXF/DWG 文件的路径和文件名
            ModelName
            活动部件文档的路径和文件名
            Action
            swExportToDWG_e中定义的导出操作
            ExportToSingleFile
            如果为 True，则另存为一个文件;假以另存为多个文件
            Alignment 
            进程内非托管C++：指向包含与输出对齐相关的信息的 12 个双精度值数组的指针（请参阅备注) 
            VBA、VB.NET、C# 和 C++/CLI：不支持                
            IsXDirFlipped
            如果为 true，则翻转 x 方向; 否则为假
             IsYDirFlipped
            如果为 true，则翻转 y 方向; 否则为假
             SheetMetalOptions
            包含钣金件导出选项的位掩码; 仅当操作 = swExportToDWG_e时有效.swExportToDWG_ExportSheetMetal（请参阅备注)
            ViewsCount
            要导出的注释视图的数量; 仅当操作 = swExportToDWG_e 时才有效.swExportToDWG_ExportAnnotationViews
             Views
            进程内非托管C++：指向要导出的注释视图名称数组的指针; 仅当操作 = swExportToDWG_e 时才有效.swExportToDWG_ExportAnnotationViews
             VBA、VB.NET、C# 和 C++/CLI：不支持
            有关此类方法的详细信息，请参阅进程内方法。
            */
            #endregion
        }
        public static void ExportToPDF()
        {
            int errors = 0;
            int warnings = 0;

            // 定义工程图文件夹路径和PDF输出文件夹路径
            string drawingFolder = @"E:\20-LeonLin\[06]Learn\#测试模型\2022";
            string pdfFolder = @"E:\20-LeonLin\[06]Learn\#测试模型\2022";

            ExportPdfData swExportPDFData = default(ExportPdfData);
            swExportPDFData = (ExportPdfData)swApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);

            ModelDocExtension swModExt = default(ModelDocExtension);

            // 循环遍历工程图文件夹中的所有SOLIDWORKS文件
            foreach (string filename in Directory.GetFiles(drawingFolder))
            {
                if (Path.GetExtension(filename).ToLower() == ".slddrw")
                {
                    // 打开SOLIDWORKS文件
                    ModelDoc2 swDoc = swApp.OpenDoc6(filename, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);

                    // 获取PDF输出文件路径
                    string pdfFilename = Path.GetFileNameWithoutExtension(filename) + ".pdf";
                    string pdfFilepath = Path.Combine(pdfFolder, pdfFilename);

                    // 导出为PDF文件
                    swModExt = (ModelDocExtension)swDoc.Extension;
                    swModExt.SaveAs(pdfFilepath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, swExportPDFData, ref errors, ref warnings);

                    // 关闭SOLIDWORKS文件
                    swApp.CloseDoc(swDoc.GetTitle());
                }
            }
        }


        /*————  进度条操作  ————*/
        public void loading_Initialize(int timelength)  //初始化进度条
        {
            progressBar1.Maximum = timelength;      //设置最大长度值
            progressBar1.Value = 0;                 //设置当前值
            progressBar1.Step = 1;                  //设置没次增长多少
        }
        public void loading()                           //更新进度条
        {
            System.Threading.Thread.Sleep(100);    //暂停1秒
            progressBar1.Value += progressBar1.Step;//让进度条增加一次
        }
        //————————————//
       
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.bilibili.com/read/cv23871167");
        }
    }
}
