using System.IO;

namespace Sw_MyAddin
{
    class log
    {
        public static void Creatlogfile(string text)    //新建日志文件
        {
            //清空内容
            FileStream fs = new FileStream(@"C:\Windows\Temp\logfile.txt", FileMode.OpenOrCreate);
            fs.Seek(0, SeekOrigin.Begin);
            fs.SetLength(0);
            fs.Close();
            //添加内容
            StreamWriter sw = new StreamWriter(@"C:\Windows\Temp\logfile.txt", true, System.Text.Encoding.GetEncoding("gb2312"));
            sw.WriteLine(text);
            sw.Flush();
            sw.Close();
        }
        public static void Openlogfile()    //打开日志文件
        {
            //File.OpenText(@"C:\Windows\Temp\logfile.txt");
            System.Diagnostics.Process.Start(@"C:\Windows\Temp\logfile.txt");
        }
    }
}
