using System;
using System.Data;
using System.IO;
using System.Reflection;
using ClosedXML.Excel;
namespace ReportProcess
{
    public static class ConvertUtility
    {
        public static void ToCSV(this System.Data.DataTable dtDataTable, string strFilePath)


        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers    
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }


        public static void ToExcel(this System.Data.DataTable dtDataTable, string strFilePath)
        {
            XLWorkbook wb = new XLWorkbook();
            //var imagePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Resources\coris.png");
            var pagina = wb.Worksheets.Add(dtDataTable, "WorksheetName");
            pagina.Row(1).InsertRowsAbove(5);
            pagina.Range("A1:C3");
            //pagina.AddPicture(imagePath).MoveTo(pagina.Cell("A1")).Scale(1.5);
            wb.SaveAs(strFilePath);

        }

    }
}
