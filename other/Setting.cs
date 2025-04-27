using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;


namespace Sw_MyAddin.other
{
    public partial class Setting : Form
    {
        public static int[] cmd_nums = Read_Settingfile();

        public Setting() { InitializeComponent(); }
        private void load(object sender, EventArgs e)//加载设置
        {
            foreach (int item in Read_Settingfile())
            {
                checkedListBox1.SetItemChecked(item, true);
            }
        }
        private void button1_Click(object sender, EventArgs e)//确认按钮
        {
            string s = null;
            List<string> cmd_check = new List<string>();
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {

                if (checkedListBox1.GetItemChecked(i))
                {
                    cmd_check.Add(i.ToString());
                    s += i.ToString() + "\n";
                }
            }
            Creat_Settingfile(s);
            this.Close();
        }
        private void button2_Click(object sender, EventArgs e)//取消按钮
        {
            this.Close();
        }


        private static void Creat_Settingfile(string text)       //新建设置文件
        {
            //清空内容
            FileStream fs = new FileStream(System.Environment.CurrentDirectory + @"\Sw_MyAddinSettingfile.txt", FileMode.OpenOrCreate);
            fs.Seek(0, SeekOrigin.Begin);
            fs.SetLength(0);
            fs.Close();
            //添加内容
            StreamWriter sw = new StreamWriter(System.Environment.CurrentDirectory + @"\Sw_MyAddinSettingfile.txt", true, System.Text.Encoding.GetEncoding("gb2312"));
            sw.WriteLine(text);
            sw.Flush();
            sw.Close();
        }
        private static int[] Read_Settingfile()                  //打开设置文件
        {
            string settingfile = System.Environment.CurrentDirectory + @"\Sw_MyAddinSettingfile.txt";
            try
            {
                StreamReader textreader = new StreamReader(settingfile);//实例化文件流对象
                List<int> sign = new List<int>(); string textline;

                //循环获得txt内容
                while ((textline = textreader.ReadLine()) != null)
                {
                    if (textline == "") { continue; }
                    sign.Add(int.Parse(textline));
                }
                int[] vs = sign.ToArray();
                textreader.Close();
                return vs;
            }
            catch (Exception)
            {
                int[] vs = { 0 };
                Creat_Settingfile("0");
                return vs;
            }


        }
    }
}
