using System;
using System.Collections.Generic;
using OfficeOpenXml;
using System.IO;
using MySql.Data.MySqlClient;
using System.Reactive.Linq;
using System.Windows;
using System.Data;
using System.Text;
using System.Configuration;

namespace LoaderModel
{
    public class OfsDataLoader
    {
        public class OfsRow
        {
            public string inn { set; get; }
            public string clname { set; get; }
            public string fs { set; get; }
            public string activity { set; get; }
            public DateTime repDate { set; get; }
            public DateTime fsMadeDate { set; get; }
            public string reptype { set; get; }
            public string regionname { set; get; }
        }
        public string filePath { set; get; }
        public bool isTestMode { set; get; }
        public string sheetName { set; get; }

        public OfsDataLoader(string Path, bool istest, string sheetname = "Лист1")
        {
            filePath = Path;
            isTestMode = istest;
            sheetName = sheetname;
 

        }

        public IEnumerable<OfsRow> ReadXlsx()
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
                            if (currentWorksheet.Cells[1, 1].Value.ToString() == "Текущая отчетность клиентов для сверки")
                            {
                                // read data from file
                                var start = currentWorksheet.Dimension.Start;
                                var end = currentWorksheet.Dimension.End;
                                for (int i = start.Row + 3; i <= end.Row; i++)
                                {
                                    var rw = new OfsRow();
                                    rw.inn = (currentWorksheet.Cells[i, 3].Value == null ? string.Empty : currentWorksheet.Cells[i, 3].Value.ToString());
                                    rw.clname = (currentWorksheet.Cells[i, 4].Value == null ? string.Empty : currentWorksheet.Cells[i, 4].Value.ToString());
                                    rw.fs = (currentWorksheet.Cells[i, 5].Value == null ? string.Empty : currentWorksheet.Cells[i, 5].Value.ToString());
                                    rw.activity = (currentWorksheet.Cells[i, 6].Value == null ? string.Empty : currentWorksheet.Cells[i, 6].Value.ToString());
                                    rw.repDate = (DateTime)(currentWorksheet.Cells[i, 7].Value);
                                    rw.fsMadeDate = (DateTime)(currentWorksheet.Cells[i, 8].Value);
                                    rw.reptype = (currentWorksheet.Cells[i, 9].Value == null ? string.Empty : currentWorksheet.Cells[i, 9].Value.ToString());
                                    rw.regionname = (currentWorksheet.Cells[i, 10].Value == null ? string.Empty : currentWorksheet.Cells[i, 10].Value.ToString());
                                    yield return rw;
                                }
                            }
                            else throw new Exception(string.Format("на 'Лист1' в ячейке (1,1) должно быть указано 'Текущая отчетность клиентов для сверки'"));
                            
                        }
                        else
                        {
                            throw new Exception(string.Format("Похоже в книге нет листа '{0}', (там должны быть данные)", sheetName));
                        }
                    }
                }
            }
        }
        public void UploadToDb()
        {

            ReadXlsx().ToObservable().
            Buffer(1000).Subscribe(loadedData =>
            {
                Upload(loadedData);
            },
                exception => { MessageBox.Show(exception.Message); },
                () => { });
        }
        private void Upload(IList<OfsRow> items)
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
            StringBuilder query = new StringBuilder("Insert into z_ofs_from_ocr  (inn, clname, fs, activity, repDate, fsMadeDate, reptype, regionname) " +
                            "values ");
            using (MySqlConnection con = new MySqlConnection(cstr))
            {
                List<string> Rows = new List<string>();
                foreach (var item in items)
                {

                    Rows.Add(string.Format("('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')",
                        MySqlHelper.EscapeString(item.inn),
                        MySqlHelper.EscapeString(item.clname),
                        MySqlHelper.EscapeString(item.fs),
                        MySqlHelper.EscapeString(item.activity),
                        MySqlHelper.EscapeString(item.repDate.ToString("yyyy-MM-dd")),
                        MySqlHelper.EscapeString(item.fsMadeDate.ToString("yyyy-MM-dd")),
                        MySqlHelper.EscapeString(item.reptype),
                        MySqlHelper.EscapeString(item.regionname)));
                }
                query.Append(string.Join(",", Rows));
                query.Append(";");
                con.Open();
                using (MySqlCommand com = new MySqlCommand(query.ToString(), con))
                {
                    try
                    {
                        com.CommandType = CommandType.Text;
                        com.ExecuteNonQuery();
                    }
                   catch( Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
            }
        }
       
    }
}

