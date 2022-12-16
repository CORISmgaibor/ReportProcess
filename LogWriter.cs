using System;
using System.IO;
using System.Reflection;

namespace ReportProcess
{
    public class LogWriter
    {
        private string m_exePath = string.Empty;
        public static Properties properties = new Properties();

        public LogWriter() { }

        public LogWriter(string logMessage)
        {
            LogWrite(logMessage);
        }
        public void LogWrite(string logMessage)
        {
            //m_exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            m_exePath = properties.get("Process.Logger.Output.File", Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
            try
            {
                using (StreamWriter w = File.AppendText(m_exePath + "\\" + "log.txt"))
                {
                    Log(logMessage, w);
                }
            }
            catch (Exception)
            {
            }
        }

        public void Log(string logMessage, TextWriter txtWriter)
        {
            try
            {
                string time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                txtWriter.WriteLine(time + ": " + logMessage);
                Console.WriteLine(time + ": " + logMessage);
            }
            catch (Exception)
            {
            }
        }
    }
}