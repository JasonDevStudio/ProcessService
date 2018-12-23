using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Text;
using System.Threading;

namespace WinProcessService
{
    public partial class MainService : ServiceBase
    {
        private Process Current { get; set; }
        private FileStream outputDataReceived = null;
        private FileStream errorDataReceived = null;
        private Timer timer = null;

        public MainService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            var dueTime = Convert.ToInt32(ConfigurationManager.AppSettings["dueTime"]);

            timer = new Timer(o =>
            {
                ProcessStart();
            }, null, dueTime, 10000);

        }

        protected override void OnStop()
        {
            timer.Dispose();
            Current.Kill();

            if (outputDataReceived != null)
            {
                outputDataReceived.Close();
                outputDataReceived.Dispose();
            }

            if (errorDataReceived != null)
            {
                errorDataReceived.Close();
                errorDataReceived.Dispose();
            }
        }

        public void WriteInfoLog(string msg)
        {
            try
            {
                var logpath = ConfigurationManager.AppSettings["log_Path"];
                var file = Path.Combine(logpath, $"Log.Info.{DateTime.Now:yyyyMMdd}.txt");

                if (outputDataReceived == null)
                {
                    outputDataReceived = new FileStream(file, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                }

                msg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss:ffff}] {msg}\t\n";
                var buffer = Encoding.Default.GetBytes(msg);

                outputDataReceived.Position = outputDataReceived.Length;
                outputDataReceived.Write(buffer, 0, buffer.Length);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void WriteErrorLog(string msg)
        {
            try
            {
                var logpath = ConfigurationManager.AppSettings["log_Path"];
                var file = Path.Combine(logpath, $"Log.Error.{DateTime.Now:yyyyMMdd}.txt");

                if (errorDataReceived == null)
                {
                    errorDataReceived = new FileStream(file, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                }

                msg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss:ffff}] {msg}\t\n";
                var buffer = Encoding.Default.GetBytes(msg);

                outputDataReceived.Position = outputDataReceived.Length;
                errorDataReceived.Write(buffer, 0, buffer.Length);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void ProcessStart()
        {
            try
            {
                if (Current == null)
                { 
                    var exec_file = ConfigurationManager.AppSettings["exec_file"];
                    var exec_args = ConfigurationManager.AppSettings["exec_args"];
                    var account = ConfigurationManager.AppSettings["account"];
                    var password = ConfigurationManager.AppSettings["password"];
                    var domain = ConfigurationManager.AppSettings["domain"];
                    var logpath = ConfigurationManager.AppSettings["log_Path"];

                    if (!Directory.Exists(logpath))
                    {
                        Directory.CreateDirectory(logpath);
                    }

                    Current = new Process
                    {
                        StartInfo = new ProcessStartInfo
                        {
                            FileName = exec_file,
                            Arguments = exec_args,
                            UseShellExecute = false,
                            RedirectStandardError = true,
                            RedirectStandardOutput = true,
                            RedirectStandardInput = true,
                            CreateNoWindow = true
                        }
                    };

                    Current.OutputDataReceived += (o, e) => WriteInfoLog(e.Data);
                    Current.ErrorDataReceived += (o, e) => WriteErrorLog(e.Data);

                    if (!string.IsNullOrWhiteSpace(account))
                    {
                        var pwds = password.ToCharArray();
                        var sstr = new System.Security.SecureString();
                        foreach (var c in pwds) sstr.AppendChar(c);

                        Current.StartInfo.Domain = domain;
                        Current.StartInfo.UserName = account;
                        Current.StartInfo.Password = sstr;

                    }

                    Current.Start();
                    Current.BeginOutputReadLine();
                    Current.BeginErrorReadLine();
                }
            }
            catch (Exception ex)
            {
                WriteErrorLog(ex.ToString());
                throw;
            }
        }
    }
}
