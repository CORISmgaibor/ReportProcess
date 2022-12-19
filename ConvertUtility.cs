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
            string imagePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Properties\coris.png");
            var pagina = wb.Worksheets.Add(dtDataTable, "WorksheetName");
            pagina.Row(1).InsertRowsAbove(5);
            pagina.Range("A1:C4").Merge();
            pagina.Range("F1:K2").Merge();
            pagina.Range("A1:M5").Style.Fill.BackgroundColor = XLColor.White;
            pagina.Range("F1:K2").Style.Font.FontSize = 18;
            pagina.Range("F1:K2").Style.Font.Bold = true;
            pagina.Range("F1:K2").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            pagina.Range("F1:K2").Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            for (int i = 1; i < 50; i++)
            {
                pagina.Column(i).AdjustToContents();
            }
            pagina.AddPicture(imagePath).MoveTo(pagina.Cell("B2")).Scale(0.8);
            pagina.Cell("F1").Value = "NOMBRE DE LA CAMPAÑA";
            wb.SaveAs(strFilePath);

        }

    }
}
