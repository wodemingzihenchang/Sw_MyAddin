using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Sw_MyAddin
{
    public static class File_edit
    {
        /*———— 获得文件对象操作  ————*/
        /// <summary>
        /// 添加文件，多选true单选false
        /// </summary>
        public static string[] Addfiles(bool isMulti)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = isMulti;
            fileDialog.Filter = "零件|*.sldprt|装配体|*.sldasm|工程图|*.slddrw|所有文件|*.*";
            fileDialog.ShowDialog();

            string[] filenames = fileDialog.FileNames;

            //string[] blank_error= { "请选择文件" };
            //if (filenames.Length == 0) { return blank_error; }
            return filenames;
        }
        public static string[] Addfiles(bool isMulti, string filter)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = isMulti;
            fileDialog.Filter = filter;
            fileDialog.ShowDialog();

            string[] filenames = fileDialog.FileNames;
            return filenames;
        }
        /// <summary>
        /// 输出文件位置
        /// </summary>
        /// <returns></returns>
        public static string Outfile()
        {
            //该对话框自带的UI不太好，可改用Ookii.Dialogs，
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            dialog.ShowDialog();
            string Path = dialog.SelectedPath; return Path;
        }
        /// <summary>
        /// 利用后缀判断文件格式
        /// </summary>
        public static void SWDocumentManager(string FileName)
        {
            if (FileName.EndsWith("sldprt")) { return; } else { return; }
        }


        /*———— txt文件操作  ————*/
        /// <summary>
        /// 把txt内容输入到窗口文本框
        /// </summary>
        /// <returns>txt字符串</returns>
        public static string Readertxt()
        {
            StreamReader textreader = new StreamReader(File_edit.Addfiles(false)[0]);//实例化文件流对象
            string text = ""; string textline;
            while ((textline = textreader.ReadLine()) != null) { text += textline + "\n"; }//循环获得txt内容
            Debug.Print(text); return text;
        }
        /// <summary>
        /// 把字符串数据输出成txt文件
        /// </summary>
        public static void Writertxt(string text)
        {
            //清空内容
            FileStream fs = new FileStream(@"C:\WINDOWS\Temp\无标题.txt", FileMode.OpenOrCreate);
            fs.Seek(0, SeekOrigin.Begin);
            fs.SetLength(0);
            fs.Close();

            //添加内容
            StreamWriter sw = new StreamWriter(@"C:\WINDOWS\Temp\无标题.txt", true, Encoding.GetEncoding("gb2312"));
            sw.WriteLine(text);
            sw.Flush();
            sw.Close();

            /*
            using (FileStream fs = new FileStream(@"C:\temp\无标题.txt", FileMode.OpenOrCreate))
            {
                string txt = text;
                byte[] buffer = Encoding.Default.GetBytes(txt);
                fs.Write(buffer, 0, buffer.Length);
            }*/
            Process.Start(@"C:\WINDOWS\Temp\无标题.txt");
        }



        /*———— string字符串操作  ————*/
        //移除/保留文字
        public static void str(string str)
        {
            str = "中1234";
            string str_num = System.Text.RegularExpressions.Regex.Replace(str, @"[\u4e00-\u9fa5]", ""); //去除汉字
            string str_text = System.Text.RegularExpressions.Regex.Replace(str, @"[^\u4e00-\u9fa5]", ""); //只留汉字
        }
        ///<summary>
        /// 移除前缀字符串
        ///</summary>
        ///<param name="val">原字符串</param>
        ///<param name="str">前缀字符串</param>
        ///<returns></returns>
        public static string GetRemovePrefixString(string val, string str)
        {
            string strRegex = @"^(" + str + ")";
            return System.Text.RegularExpressions.Regex.Replace(val, strRegex, "");
        }
        ///<summary>
        /// 移除后缀字符串
        ///</summary>
        ///<param name="val">原字符串</param>
        ///<param name="str">后缀字符串</param>
        ///<returns></returns>
        public static string GetRemoveSuffixString(string val, string str)
        {
            string strRegex = @"(" + str + ")" + "$";
            return System.Text.RegularExpressions.Regex.Replace(val, strRegex, "");
        }
        /// <summary>
        /// 截取指定字符后面信息
        /// </summary>
        /// <param name="meterial"></param>
        /// <param name="a"></param>
        public static void str(string meterial, char x)
        {
            meterial = "80X8铜排";
            x = 'X';
            meterial = System.Text.RegularExpressions.Regex.Replace(meterial, @"[\u4e00-\u9fa5]", ""); //去除汉字，结果：80X8
            string Thickness = meterial.Substring(meterial.LastIndexOf(x) + 1);//结果：8
        }

        /*———— 其他操作  ————*/
        /// <summary>
        /// UTF-8字符转数字
        /// </summary>
        /// <param name="str"></param>
        public static string Utf8toint(string str)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(str);
            string[] strArr = new string[bytes.Length];

            string outstr = "";
            for (int i = 0; i < bytes.Length; i++)
            {
                strArr[i] = bytes[i].ToString("");
                outstr = outstr + strArr[i];
            }
            Console.WriteLine(outstr); return outstr;

            /*Console.WriteLine("从16进制转换回汉字：");
            for (int i = 0; i < strArr.Length; i++)
            {
                bytes[i] = byte.Parse(strArr[i], System.Globalization.NumberStyles.HexNumber);
            }
            string ret = Encoding.Unicode.GetString(bytes);
            Console.WriteLine(ret);*/
        }
        /// <summary>
        /// 利用字符用GetHashCode()方法串输出随机数0~255
        /// </summary>
        /// <returns></returns>
        public static int[] StrToint(string str)
        {
            int number = -414239725;
            //随机数方法1
            Random r = new Random();
            number = r.Next(0, 255);
            //随机数方法2
            int b = str.GetHashCode(); Debug.WriteLine(b);
            //将长字符转255输出
            int number1 = b % 1000; while (number1 < 0 || number1 > 255)
            {
                if (number1 > 255)
                {
                    number1 -= 255;
                }
                else if (number1 < 0)
                {
                    number1 += 255;
                }
            }
            int number2 = b / 1000 % 1000; while (number2 < 0 || number2 > 255)
            {
                if (number2 > 255)
                {
                    number2 -= 255;
                }
                else if (number2 < 0)
                {
                    number2 += 255;
                }
            }
            int number3 = b / 1000000 % 1000; while (number3 < 0 || number3 > 255)
            {
                if (number3 > 255)
                {
                    number3 -= 255;
                }
                else if (number3 < 0)
                {
                    number3 += 255;
                }
            }

            int[] numbers = { number1, number2, number3 };
            Debug.WriteLine(numbers[0] + "," + numbers[1] + "," + numbers[2]);
            return numbers;
        }
        //代码运行时间
        public static void Runtime()
        {
            System.Diagnostics.Stopwatch watch = new System.Diagnostics.Stopwatch();
            watch.Start();  //开始监视代码运行时间
                            //需要测试的代码
            watch.Stop();  //停止监视
            TimeSpan timespan = watch.Elapsed;  //获取当前实例测量得出的总时间
            System.Diagnostics.Debug.WriteLine("打开窗口代码执行时间：{0}(毫秒)", timespan.TotalMilliseconds);  //总毫秒数 
        }

        //类型判断
        public static void Type_Check()
        {
            object obj = null;
            string str = null;
            if (obj.GetType() != typeof(string)) { str = Convert.ToString(obj); }
        }

        //后台启动程序
        public static void ProStart_Hide(string proname)
        {
            Process pro = new Process();
            ProcessStartInfo proStartInfo = new ProcessStartInfo(proname);
            proStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            pro.StartInfo = proStartInfo;

            Process[] pro2 = Process.GetProcessesByName("Sw_toolkit");
            if (pro2.Length > 1 || pro2.Length > 1)
            {
                return;  //MessageBox.Show("目前已有一个安装程序在运行，请勿重复运行程序!", "提示");
            }
            //pro.Start();

            //Process[] pro1 = Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName);

        }

        class FileArchives
        {
            [DllImport("kernel32.dll", EntryPoint = "SetProcessWorkingSetSize")]
            public static extern int SetProcessWorkingSetSize(IntPtr process, int minSize, int maxSize);
        }
    }
}
