using Microsoft.Data.SqlClient;
using System;
using System.IO;
using System.Reflection;


namespace ReportProcess
{
    class Reporter
    {
        public static Properties properties;
        public static LogWriter logger;
        static void Main(string[] args)
        {
            init();
            while (true)
            {
                run();
                pause();
            }
        }

        private static void pause()
        {
            int valorTiempo = Int16.Parse(properties.get("Process.Pause.Minutes", "1"));
            int tiempoPausa = 1000 * 60 * valorTiempo; // de minutos a milisegundos
            logger.LogWrite("PausoProceso " + tiempoPausa + " minutos...");
            System.Threading.Thread.Sleep(tiempoPausa);
        }

        private static void run()
        {
            string strConn = properties.get("Process.Connection", "");
            SqlConnection conn = new SqlConnection(strConn);
            SqlCommand cmd = new SqlCommand("", conn);

            try
            {
                conn.Open();

            }
            catch (Exception e)
            {
                logger.LogWrite("ERROR: " + e.ToString());
            }
            finally
            {
                conn.Close();
                Console.ReadLine();
            }
        }

        private static void init()
        {
            properties = new Properties();
            logger = new LogWriter();
            logger.LogWrite("Inicio proceso automático de envio de Reportes...");
        }
    }
}
