using MySql.Data.MySqlClient;
using System;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Text;
using System.Data;
using System.Configuration;

namespace LoaderModel
{
    class KpfDataLoader : XlDataLoader
    {
        public KpfDataLoader(string Path, bool Istest, string SheetName) : base(Path, Istest, SheetName)
        {
            
        }
        public override IEnumerable<XlRow> ReadXlsx()
        {
            // Get the file we are going to process
            var existingFile = new FileInfo(filePath);
            // Open and read the XlSX file.
            using (var package = new ExcelPackage(existingFile))
            {
                // Get the work book in the file
                var workBook = package.Workbook;
                if (workBook != null)
                {
                    if (workBook.Worksheets.Count > 0)
                    {
                        // Get the 'sheetName' worksheet
                        var currentWorksheet = workBook.Worksheets[sheetName];
                        if (currentWorksheet != null)
                        {
                            // read data from file
                            var start = currentWorksheet.Dimension.Start;
                            var end = currentWorksheet.Dimension.End;

                            for (int i = start.Row + 5; i <= end.Row; i++)
                            {
                                var rw = new XlRow();
                                rw.Dt = (DateTime)(currentWorksheet.Cells[3, 3].Value);
                                rw.Type = sheetName;
                                rw.Acc = (currentWorksheet.Cells[i, 2].Value == null ? string.Empty : currentWorksheet.Cells[i, 2].Value.ToString().Trim());
                                rw.OstRub = Helpers.ParseToDoubleEx((currentWorksheet.Cells[i, 3].Value == null ? "0.00" : currentWorksheet.Cells[i, 3].Value.ToString().Trim()),sheetName,i);
                                rw.ClientName = (currentWorksheet.Cells[i, 4].Value == null ? string.Empty : currentWorksheet.Cells[i, 4].Value.ToString().Trim());
                                rw.Inn = (currentWorksheet.Cells[i, 5].Value == null ? string.Empty : currentWorksheet.Cells[i, 5].Value.ToString().Trim());
                                rw.DealNumber = (currentWorksheet.Cells[i, 6].Value == null ? string.Empty : currentWorksheet.Cells[i, 6].Value.ToString().Trim());
                                rw.Cq = Helpers.CqConverter(currentWorksheet.Cells[i, 7].Value == null ? "0" : currentWorksheet.Cells[i, 7].Value.ToString().Trim());
                                rw.Norm = Helpers.ParseToFloat(currentWorksheet.Cells[i, 8].Value == null ? "-0.01" :  currentWorksheet.Cells[i, 8].Value.ToString().Trim());
                                rw.ResAcc = (currentWorksheet.Cells[i, 9].Value == null ? string.Empty : currentWorksheet.Cells[i, 9].Value.ToString().Trim());
                                rw.Reserv = Helpers.ParseToDoubleEx((currentWorksheet.Cells[i, 10].Value == null ? "0.00" : currentWorksheet.Cells[i, 10].Value.ToString().Trim()),sheetName,i);
                                rw.Ob1= Helpers.ParseToDoubleEx((currentWorksheet.Cells[i, 11].Value == null ? "0.00" : currentWorksheet.Cells[i, 11].Value.ToString().Trim()),sheetName, i);
                                rw.Ob2 = Helpers.ParseToDoubleEx((currentWorksheet.Cells[i, 12].Value == null ? "0.00" : currentWorksheet.Cells[i, 12].Value.ToString().Trim()), sheetName, i);
                                rw.Ocr = (currentWorksheet.Cells[1, 5].Value == null ? "Ошибка в файле" : currentWorksheet.Cells[1, 5].Value.ToString().Trim());
                                yield return rw;
                            }
                        }
                        else
                        {
                            throw new Exception(string.Format("В файле КПФ Нет листа '{0}'", sheetName));
                        }
                    }
                }
            }
        }

        public override void Upload(IList<XlRow> items)
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
            StringBuilder query = new StringBuilder("Insert into z_rzrv_from_ocr  (bydate, type, acc, ostrub, clientName, inn, dealNumber, cq, norm, ResAcc,  reserv, ob1, ob2,  ocr) " +
                           "values ");
            using (MySqlConnection con = new MySqlConnection(cstr))
            {
                List<string> Rows = new List<string>();
                foreach (var item in items)
                {
                   
                        Rows.Add(string.Format("('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}')",
                            MySqlHelper.EscapeString(item.Dt.ToString("yyyy-MM-dd")),
                            MySqlHelper.EscapeString(item.Type),
                            MySqlHelper.EscapeString(item.Acc),
                            MySqlHelper.EscapeString(item.OstRub.ToString()),
                            MySqlHelper.EscapeString(item.ClientName),
                            MySqlHelper.EscapeString(item.Inn),
                            MySqlHelper.EscapeString(item.DealNumber),
                            MySqlHelper.EscapeString(item.Cq.ToString()),
                            MySqlHelper.EscapeString(item.Norm.ToString()),
                            MySqlHelper.EscapeString(item.ResAcc),
                            MySqlHelper.EscapeString(item.Reserv.ToString()),
                            MySqlHelper.EscapeString(item.Ob1.ToString()),
                            MySqlHelper.EscapeString(item.Ob2.ToString()),
                            MySqlHelper.EscapeString(item.Ocr)));
                }
                query.Append(string.Join(",", Rows));
                query.Append(";");
                con.Open();
                using (MySqlCommand com = new MySqlCommand(query.ToString(), con))
                {
                    com.CommandType = CommandType.Text;
                    com.ExecuteNonQuery();
                }
            }
        }


       
    }
}

