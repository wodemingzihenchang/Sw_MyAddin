using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;

namespace Sw_toolkit
{
    public partial class 批量修改材质 : Form
    {
        public string sldmatpath = null;
        public string[] filepath = null;

        public 批量修改材质() { InitializeComponent(); }

        private void button1_Click(object sender, EventArgs e)//选择材质库
        {
            sldmatpath = File.Addfiles(false)[0];
            button1.Text = System.IO.Path.GetFileNameWithoutExtension(sldmatpath);
        }

        private void button2_Click(object sender, EventArgs e)//选择文件
        {
            //选择文件
            filepath = File.Addfiles(true);
            //进度条
            progressBar1.Value = 0; label1.Text = "进度： ";
            progressBar1.Maximum = filepath.Length;
            //判断是否选择文件，
            if (filepath.Length == 0) { return; }
            //设置表格行列
            dataGridView1.RowCount = filepath.Length;
            dataGridView1.ColumnCount = 3;
            dataGridView1.Columns[0].HeaderText = "文件位置";
            dataGridView1.Columns[1].HeaderText = "文件配置";
            dataGridView1.Columns[2].HeaderText = "零件材质";
            for (int i = 0; i < filepath.Length; i++)
            {//写入表格内容
                this.dataGridView1.Rows[i].Cells[0].Value = Path.GetFileName(filepath[i]);//文件名称
                this.dataGridView1.Rows[i].Cells[1].Value = filepath[i];//文件路径
                progressBar1.Value += 1;//进度条
            }
            //进度条
            label1.Text = "进度：完成"; progressBar1.Value = progressBar1.Maximum;
        }

        private void button3_Click(object sender, EventArgs e)//运行修改
        {
            //进度条
            progressBar1.Value = 0; label1.Text = "进度： ";
            progressBar1.Maximum = dataGridView1.RowCount;
            //
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                //打开文件
                ModelDoc2 swDoc = SwModelDoc.swApp.OpenDoc(dataGridView1.Rows[i].Cells[1].Value.ToString(), 1);
                PartDoc part = (PartDoc)swDoc; string configname = swDoc.GetActiveConfiguration().Name;
                //空值跳过
                if (dataGridView1.Rows[i].Cells[2].Value == null) { continue; }
                //修改材质
                part.SetMaterialPropertyName2(configname, sldmatpath, dataGridView1.Rows[i].Cells[2].Value.ToString()); part.EditRebuild();
                //保存关闭
                swDoc.Save(); SwModelDoc.swApp.CloseDoc(dataGridView1.Rows[i].Cells[1].Value.ToString());
                //进度条
                progressBar1.Value += 1;
            }
            //进度条
            label1.Text = "进度：完成"; progressBar1.Value = progressBar1.Maximum;
            MessageBox.Show("完成"); this.Close();
        }

        private void Debug_Click(object sender, EventArgs e)//Debug
        {
            ModelDoc2 swDoc = SwModelDoc.swApp.ActiveDoc;
            PartDoc part = (PartDoc)swDoc;
            string configname = swDoc.GetActiveConfiguration().Name;
            part.SetMaterialPropertyName2(configname, sldmatpath, "S45C"); part.EditRebuild();

            //part.SetMaterialPropertyName2("默认<按加工>", "C:/ProgramData/SolidWorks/SOLIDWORKS 2018/自定义材料/治具常用材料-A.sldmat", "S45C"); part.EditRebuild();
        }





        //——————★Excel操作——————//
        private void Save_Excel_Click(object sender, EventArgs e)//保存Excel数据
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel|*.xls|Excel|*.xlsx"
            };
            if (saveFileDialog.ShowDialog() == DialogResult.OK && saveFileDialog.FileName != "")
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null) { MessageBox.Show("无Excel"); return; }

                Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
                Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets[1];

                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    worksheet.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
                }
                //写入数值
                for (int r = 0; r < dataGridView1.RowCount; r++)
                {
                    progressBar1.Value = dataGridView1.RowCount - 1;
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        worksheet.Cells[r + 2, i + 1] = dataGridView1.Rows[r].Cells[i].Value;
                    }
                }
                worksheet.Columns.EntireColumn.AutoFit();
                MessageBox.Show("资料保存成功");
                workbook.Saved = true;
                workbook.SaveCopyAs(saveFileDialog.FileName);
                xlApp.Quit();

                //进度条
                label1.Text = "进度：完成"; progressBar1.Value = progressBar1.Maximum;
            }
        }
        private void Open_Excel_Click(object sender, EventArgs e)//打开Excel数据
        {
            OpenFileDialog fd = new OpenFileDialog
            {
                Filter = "Excel(*.xls)|*.xls|Excel(*.xlsx)|*.xlsx"
            };
            string strpath;
            if (fd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    strpath = fd.FileName;

                    //string strCon = "provider=microsoft.jet.oledb.4.0;data source=" + strpath + ";extended properties=excel 8.0";//32位平台
                    //string strCon = "provider=microsoft.ACE.OLEDB.12.0" + strpath + ";extended properties=excel 8.0";//需要安装Access Data engin数据库引擎
                    string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strpath + ";Extended Properties=\"Excel 12.0;HDR=YES\"";//64位平台可用

                    OleDbConnection Con = new OleDbConnection(strCon);
                    string strSql = "select* from [Sheet1$]";
                    OleDbCommand cmd = new OleDbCommand(strSql, Con);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "(需使用Excel 2003.xls版本)");
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
    }
}
