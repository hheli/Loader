using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reactive.Linq;
using System.Text;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;

namespace LoaderModel
{
    public class  PsDataLoader
    {
        /// <summary>
        ///  Класс для загрузки профсуждений ГО
        /// </summary>

        public string filePath { set; get; }
        public bool isTestMode { set; get; }
        public PsDataLoader(string Path, bool istest)
        {
            filePath = Path;
            isTestMode = istest;
           // DeleteData();
        }
        public void UploadToDb()
        {
           // DeleteData();
            Upld();
        }
        private void Upld()
        {
            string cstr;
            DateTime dt;
            if (isTestMode)
            {
                cstr = ConfigurationManager.ConnectionStrings["TestString"].ConnectionString;

            }
            else
            {
                cstr = ConfigurationManager.ConnectionStrings["RealString"].ConnectionString;

            }
            StringBuilder query = new StringBuilder("Insert into z_ps  (onDate, OCR, ClientName, AccountName, AccountNumber, OstR, OFS, OOD, CQ, Norm, Res, Ob1, Ob2,  OtherFCR, cid, aid, IU) " +
                            "values ");
            using (MySqlConnection con = new MySqlConnection(cstr))
            {
            try
                {

                    List<string> Rows = new List<string>();
                    var excel = new Excel.Application();
                    var wb = excel.Workbooks.Open(filePath);
                    Excel._Worksheet worksheet1 = (Excel._Worksheet)WorkbookExtensions.GetWorksheetByName(wb, "Классификация клиентов");
                    if (worksheet1 != null)
                    {

                        dt = worksheet1.Cells[1, 3].value;
                        Excel._Worksheet worksheet = (Excel._Worksheet)WorkbookExtensions.GetWorksheetByName(wb, "ПрофСуж");
                        if (worksheet != null)
                        {
                            Excel.Range xlRange = worksheet.UsedRange;
                            int rowCount = xlRange.Rows.Count;
                            int colCount = xlRange.Columns.Count;
                            for (int i = 9; i <= rowCount; i++) //начинаем со строки 9
                            {
                                if (worksheet.Cells[i, 1].value == null || worksheet.Cells[i, 1].value == "") { continue; }
                                Rows.Add(string.Format("('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}')",
                                    MySqlHelper.EscapeString(dt.ToString("yyyy-MM-dd")),
                                    MySqlHelper.EscapeString("ГО"),
                                    MySqlHelper.EscapeString(worksheet.Cells[i, 1].value),
                                    MySqlHelper.EscapeString(worksheet.Cells[i, 2].value),
                                    MySqlHelper.EscapeString(worksheet.Cells[i, 3].value),
                                    MySqlHelper.EscapeString(Convert.ToString(worksheet.Cells[i, 4].value ?? 0.00)),
                                    MySqlHelper.EscapeString(worksheet.Cells[i, 5].value),
                                    MySqlHelper.EscapeString(worksheet.Cells[i, 6].value),
                                    MySqlHelper.EscapeString(Convert.ToString(worksheet.Cells[i, 7].value)),
                                    MySqlHelper.EscapeString(Convert.ToString(worksheet.Cells[i, 8].value ?? 0.00)),
                                    MySqlHelper.EscapeString(Convert.ToString(worksheet.Cells[i, 9].value ?? 0.00)),
                                    MySqlHelper.EscapeString(Convert.ToString(worksheet.Cells[i, 10].value ?? 0.00)),
                                    MySqlHelper.EscapeString(Convert.ToString(worksheet.Cells[i, 11].value ?? 0.00)),
                                    MySqlHelper.EscapeString(worksheet.Cells[i, 12].value),
                                    MySqlHelper.EscapeString(Convert.ToString(worksheet.Cells[i, 13].value ?? 0)),
                                    MySqlHelper.EscapeString(Convert.ToString(worksheet.Cells[i, 14].value ?? 0)),
                                    MySqlHelper.EscapeString(worksheet.Cells[i, 15].value)));

                            }
                            query.Append(string.Join(",", Rows));
                            query.Append(";");
                            wb.Close();
                            con.Open();
                            using (MySqlCommand com = new MySqlCommand(query.ToString(), con))
                            {
                                com.CommandType = CommandType.Text;
                                com.ExecuteNonQuery();
                            }

                        }
                    }

                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message);
                }
               
            }
        }

        private void DeleteData()
        {
            string cstr;
            if (isTestMode)
            {
                cstr = ConfigurationManager.ConnectionStrings["TestString"].ConnectionString;

            }
            else
            {
                cstr = ConfigurationManager.ConnectionStrings["RealString"].ConnectionString;

            }
            StringBuilder query = new StringBuilder("Delete from z_ps;");
            using (MySqlConnection con = new MySqlConnection(cstr))
            {
                con.Open();
                using (MySqlCommand com = new MySqlCommand(query.ToString(), con))
                {
                    try
                    {
                        com.CommandType = CommandType.Text;
                        com.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
            }
        }
    }
    public static class WorkbookExtensions
    {
        public static Excel.Worksheet GetWorksheetByName(this Excel.Workbook workbook, string name)
        {
            return workbook.Worksheets.OfType<Excel.Worksheet>().FirstOrDefault(ws => ws.Name == name);
        }
    }
}
