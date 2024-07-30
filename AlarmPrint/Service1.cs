using System;
using System.Configuration;
using System.ServiceProcess;
using System.Diagnostics;
using System.Threading;
using System.Net;
using System.Drawing.Printing;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace AlarmPrint
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        static HttpListener serverWeb;
        Thread WEBServer = new Thread(ThreadWEBServer);

        static PrintDocument printDoc = new PrintDocument();

        static string message = "";
        static bool statusWeb = true;
        static int logFileName = 0;
        static string namePrinter = "MPT-II";
        static string alarmPrinterPort = "8080";

        void LoadConfig()
        {
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\AlarmPrinter", true))
            {
                if (key.GetValue("FailureActions") == null)
                {
                    key.SetValue("FailureActions", new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00, 0x00, 0x00, 0x14, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x60, 0xea, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x60, 0xea, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x60, 0xea, 0x00, 0x00 });
                }
            }
            if (ConfigurationManager.AppSettings.Count != 0)
            {
                namePrinter = ConfigurationManager.AppSettings["NamePrinter"];
                alarmPrinterPort = ConfigurationManager.AppSettings["AlarmPrinterPort"];
            }
        }

        static void PrintMessage()
        {
            printDoc.PrinterSettings.PrinterName = namePrinter;
            printDoc.PrintPage += PrintPageHandler;
            printDoc.Print();
        }

        static void PrintPageHandler(object sender, PrintPageEventArgs e)
        {
            e.Graphics.DrawString(message, new Font("Arial", 7), Brushes.Black, 0, 0);
            e.Graphics.DrawString("***************************************************", new Font("Arial", 7), Brushes.Black, 0, (e.Graphics.DpiY /2) -10);

            LogWriteLine(message);
        }

        static void LogWriteLine(string message)
        {
            if (!(Directory.Exists(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\log")))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\log");
            }

            string logDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\log";

            string[] tempfiles = Directory.GetFiles(logDir, "-log.txt", SearchOption.AllDirectories);

            if (tempfiles.Count() != 0)
            {
                foreach (string file in tempfiles)
                {
                    string names = Path.GetFileName(file);
                    Regex regex = new Regex(@"\d{4}-");
                    if (regex.IsMatch(names))
                    {
                        int number = (int.Parse(names.Remove(names.IndexOf("-"))));
                        if (number > logFileName)
                        {
                            logFileName = number;
                        }
                    }
                }
            }

            string name = logFileName.ToString("0000");
            FileInfo fileInfo = new FileInfo(logDir + $"\\{name}-log.txt");
            using (StreamWriter sw = fileInfo.AppendText())
            {
                sw.WriteLine(String.Format("{0:yyMMdd hh:mm:ss}\r\n{1}\r\n\r\n", DateTime.Now.ToString(), message));
                sw.Close();
                if (fileInfo.Length > 20480)
                {
                    logFileName++;
                }

                string[] delTimefiles = Directory.GetFiles(logDir, "*", SearchOption.AllDirectories);
                foreach (string delTimefile in delTimefiles)
                {
                    FileInfo fi = new FileInfo(delTimefile);
                    if (fi.CreationTime < DateTime.Now.AddDays(-10)) { fi.Delete(); }
                }
            }
        }

        static void ThreadWEBServer()
        {
            serverWeb = new HttpListener();
            serverWeb.Prefixes.Add(@"http://+:" + alarmPrinterPort + "/");
            serverWeb.AuthenticationSchemes = AuthenticationSchemes.Anonymous;
            serverWeb.Start();
            while (statusWeb)
            {
                ProcessRequest();
            }
        }

        static void ProcessRequest()
        {
            var result = serverWeb.BeginGetContext(ListenerCallback, serverWeb);
            var startNew = Stopwatch.StartNew();
            result.AsyncWaitHandle.WaitOne();
            startNew.Stop();
        }

        static void ListenerCallback(IAsyncResult result)
        {
            var HttpResponse = serverWeb.EndGetContext(result);
            var context = HttpResponse.Request;
            using (var reader = new StreamReader(context.InputStream, System.Text.Encoding.UTF8 ))
            {
                message = reader.ReadToEnd();
            }
            PrintMessage();
            HttpResponse.Response.ContentType = "text/plain";
            HttpResponse.Response.StatusCode = 200;
            HttpResponse.Response.OutputStream.Write(new byte[] { }, 0, 0);
            HttpResponse.Response.Close();
        }

        protected override void OnStart(string[] args)
        {
            LoadConfig();
            WEBServer.Start();
        }

        protected override void OnStop()
        {
            statusWeb = false;
            WEBServer.Interrupt();
        }
    }
}
