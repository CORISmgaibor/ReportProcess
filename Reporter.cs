using Microsoft.Data.SqlClient;
using System;
using System.Data;

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
            string descripcion = "", nombreQuery = "", cadena = "", extension = "", repoSalida = "";
            SqlConnection conn = new SqlConnection(strConn);
            //SqlCommand cmd = new SqlCommand(@"SELECT A.DESCRIPCION , B.NOMBRE , D.CADENA , F.EXTENSION 
            //                                  FROM CAMPANA A JOIN QUERY B ON A.IDCAMPANA = B.IDCAMPANA
            //                                  JOIN CONEXIONES_QUERY C ON C.IDQUERY=B.IDQUERY 
            //                                  JOIN CONEXIONES D ON D.IDCONEXIONES = C.IDCONEXIONES
            //                                  JOIN FORMATO_QUERY E ON E.IDQUERY = B.IDQUERY
            //                                  JOIN FORMATOS F ON E.IDFORMATO = F.IDFORMATO
            //                                  WHERE A.ESTADO = 1 AND B.ESTADO = 1", conn);
            SqlCommand cmd = new SqlCommand(@"SELECT 'CAMPAÑA DE PRUEBA' AS DESCRIPCION , 'SP_TEST' AS NOMBRE , 'Data Source=192.168.0.185;Initial Catalog=Reportiador;Persist Security Info=True;User ID=sa;Password=Administrator1;TrustServerCertificate=True' AS CADENA, 'XLSX' AS EXTENSION, '' AS PARAMETROS , '\\192.168.10.4\Compartida temp\Reporteador\' as URLSALIDA", conn);

            try
            {
                conn.Open();
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    descripcion = (string)reader["DESCRIPCION"];
                    nombreQuery = (string)reader["NOMBRE"];
                    cadena = (string)reader["CADENA"];
                    extension = (string)reader["EXTENSION"];
                    repoSalida = (string)reader["URLSALIDA"];

                    SqlConnection tempConn = new SqlConnection(cadena);
                    SqlCommand tempCmd = new SqlCommand(nombreQuery, tempConn);
                    tempConn.Open();
                    try
                    {
                        tempCmd.CommandType = System.Data.CommandType.StoredProcedure;
                        SqlDataAdapter da = new SqlDataAdapter(tempCmd);
                        DataTable table = new DataTable();
                        da.Fill(table);
                        //table.ToCSV(repoSalida + descripcion + ".csv"  );
                        table.ToExcel(repoSalida + descripcion + "." + extension);

                    }
                    catch (Exception e)
                    {
                    }
                    finally
                    {
                        tempConn.Close();
                    }
                }

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
