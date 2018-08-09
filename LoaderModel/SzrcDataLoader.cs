using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using MySql.Data.MySqlClient;
using System.Reactive.Linq;
using System.Linq;
using System.Text;
using System.Windows;
using System.Data;
using System.Configuration;

namespace LoaderModel
{
    class SzrcDataLoader : XlDataLoader
    {
        public SzrcDataLoader(string Path, bool Istest, string SheetName) : base(Path, Istest, SheetName)
        {
           
        }
        override public IEnumerable<XlRow> ReadXlsx()
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
                                rw.Dt = (DateTime)(currentWorksheet.Cells[3, 2].Value);
                                rw.Acc = (currentWorksheet.Cells[i, 4].Value == null ? string.Empty : currentWorksheet.Cells[i, 4].Value.ToString());
                                rw.OstRub = Helpers.ParseToDouble(currentWorksheet.Cells[i, 5].Value == null ? "0.00" : currentWorksheet.Cells[i, 5].Value.ToString());
                                rw.ClientName = (currentWorksheet.Cells[i, 2].Value == null ? string.Empty : currentWorksheet.Cells[i, 2].Value.ToString().Replace("''","\""));
                                rw.Inn = (currentWorksheet.Cells[i, 1].Value == null ? string.Empty : currentWorksheet.Cells[i, 1].Value.ToString().Trim());
                                
                                rw.DealNumber = (currentWorksheet.Cells[i, 3].Value == null ? string.Empty : currentWorksheet.Cells[i, 3].Value.ToString());
                                rw.Cq = int.Parse(currentWorksheet.Cells[i, 6].Value == null ? "0" : currentWorksheet.Cells[i, 6].Value.ToString());
                                rw.Norm = Helpers.ParseToFloat(currentWorksheet.Cells[i, 7].Value == null ? "-0.01" : currentWorksheet.Cells[i, 7].Value.ToString());
                                rw.Reserv = Helpers.ParseToDouble(currentWorksheet.Cells[i, 8].Value == null ? "0.00" : currentWorksheet.Cells[i, 8].Value.ToString());
                                if (currentWorksheet.Cells[1, 1].Value.ToString() == "Ссудная задолженность СЗРЦ и филиалов СЗРЦ")
                                {
                                    rw.Type = "РВПС";
                                    rw.Ocr = "СЗРЦ_" + (currentWorksheet.Cells[i, 9].Value == null ? string.Empty : currentWorksheet.Cells[i, 9].Value.ToString());
                                    rw.ResAcc = "";
                                }
                                else
                                {
                                    rw.Type = "РВП";
                                    rw.ResAcc = (currentWorksheet.Cells[i, 9].Value == null ? string.Empty : currentWorksheet.Cells[i, 9].Value.ToString());
                                    rw.Ocr = "СЗРЦ_" + (currentWorksheet.Cells[i, 10].Value == null ? string.Empty : currentWorksheet.Cells[i, 10].Value.ToString());
                                    
                                }
                                yield return rw;
                            }
                        }
                        else
                        {
                            throw new Exception(string.Format("В файле СЗРЦ Нет листа '{0}'", sheetName));
                        }
                    }
                }
            }
        }



        override public void Upload(IList<XlRow> items)
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
            StringBuilder query = new StringBuilder("Insert into z_rzrv_from_ocr  (bydate, type, acc, ostrub, clientName, inn, dealNumber, cq, norm, ResAcc,  reserv,   ocr) " +
                           "values ");
            using (MySqlConnection con = new MySqlConnection(cstr))
            {
                List<string> Rows = new List<string>();
                foreach (var item in items)
                {
                    Rows.Add(string.Format("('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')",
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
                        MySqlHelper.EscapeString(item.Ocr)));
                }
                query.Append(string.Join(",", Rows));
                query.Append(";");
                con.Open();
                using (MySqlCommand com = new MySqlCommand(query.ToString(),con))
                {
                    com.CommandType= CommandType.Text;
                    com.ExecuteNonQuery();
                }
            }
               
        }
    }
}
