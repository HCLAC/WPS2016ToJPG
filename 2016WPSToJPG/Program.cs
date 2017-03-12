using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Configuration;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Imaging;
using System.Threading;

namespace WPSToJPG
{
    class Program
    {
        public static void RunCheckWPSApp()
        {
            try
            {

                Process[] process = Process.GetProcessesByName("CheckWPSApp");
                if (process.Length == 0)
                {
                    string exepath = System.Environment.CurrentDirectory;
                    ProcessStartInfo p = new ProcessStartInfo("CheckWPSApp.exe", null);
                    p.WorkingDirectory = exepath;                       //设置此外部程序所在windows目录
                    Console.WriteLine(DateTime.Now + " " + exepath + "\\CheckWPSApp.exe");
                    Process Proc = System.Diagnostics.Process.Start(p); //调用外部程序
                }
                else if (process.Length > 1)
                {
                    for (int i = 0; i < process.Length - 1; i++)
                    {
                        Process prtemp = process[i];
                        prtemp.Kill();
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(/*"[" + DateTime.Today + "]  " + */ex.Message.ToString());
                //m_nConvertStatus = EErrorType.OTP_EXCEPTION_FAILED;
            }
        }
        // 必须有这个，Acrobat只能在单线程中运行。
        [STAThread]
        public static void Main(string[] args)
        {
            // 输入的参数必须大于2个，
            // 格式：
            //    args[1]：需要转换的文档路径，
            //    args[2]：保存图片的目录，
            //    args[3]：图片质量等级（1-10），
            //    args[4]：如果转换成中间PDF，是否删除PDF文件
            if (args.Length < 2)
            {
                return;
            }
            CWPSToJPG.Definition nQuality = CWPSToJPG.Definition.Three;
            bool bDelePDF = false;
            if (!System.IO.File.Exists(args[0]))
            {// 需要转换的文件不存在，退出
#if DEBUG
                Console.WriteLine("[" + DateTime.Today + "]  " + "no exist :" + args[0]);
#else
                Console.WriteLine("ConvertStatus:3 17 0#");
#endif
                return;
            }
            if (!System.IO.Directory.Exists(args[1]))
            {// 保存图片的目录不存在，退出
#if DEBUG
                Console.WriteLine("[" + DateTime.Today + "]  " + "no exist :" + args[1]);
#else
                Console.WriteLine("ConvertStatus:3 17 0#");
#endif
                return;
            }
            if (args.Length > 2)
                nQuality = (CWPSToJPG.Definition)Math.Min(Math.Max(int.Parse(args[2]), 1), 10);
            if (args.Length > 3)
                bDelePDF = bool.Parse(args[3]);
            // 启动监控程序，监测WPS进程和Acrobat进程
            // RunCheckWPSApp();
            // 开始转换WPSToJPG
            CWPSToJPG WJ = new CWPSToJPG();
            WJ.OfficeToJPGEx(args[0], args[1], nQuality, bDelePDF);
        }
    }
}
