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
    public class CrmDataLoader
    {
        public class SlxRow
        {
            public string slxCode { set; get; }
            public string ogrn { set; get; }
            public string inn { set; get; }
            public string name { set; get; }
            public string category { set; get; }
            public string department { set; get; }
            public string upravlenie { set; get; }
            public string swift { set; get; }
            public string isactive { set; get; }
            public string status { set; get; }
            
         
        }
        public string filePath { set; get; }
        public bool isTestMode { set; get; }
        public string sheetName { set; get; }

        public CrmDataLoader(string Path, bool istest, string sheetname= "Все клиенты Банка")
        {
            filePath = Path;
            isTestMode = istest;
            sheetName = sheetname;
            //DeleteData();

        }

        public IEnumerable<SlxRow> ReadXlsx()
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
                            for (int i = start.Row + 1; i <= end.Row; i++)
                            {
                                var rw = new SlxRow();
                                rw.slxCode = (currentWorksheet.Cells[i, 1].Value==null? string.Empty: currentWorksheet.Cells[i, 1].Value.ToString());
                                rw.ogrn = (currentWorksheet.Cells[i, 2].Value == null ? string.Empty : currentWorksheet.Cells[i, 2].Value.ToString());
                                rw.inn = (currentWorksheet.Cells[i, 3].Value == null ? string.Empty : currentWorksheet.Cells[i, 3].Value.ToString());
                                rw.name = (currentWorksheet.Cells[i, 4].Value == null ? string.Empty : currentWorksheet.Cells[i, 4].Value.ToString());
                                rw.category = (currentWorksheet.Cells[i, 5].Value == null ? string.Empty : currentWorksheet.Cells[i, 5].Value.ToString());
                                rw.department = (currentWorksheet.Cells[i, 6].Value == null ? string.Empty : currentWorksheet.Cells[i, 6].Value.ToString());
                                rw.upravlenie = (currentWorksheet.Cells[i, 7].Value == null ? string.Empty : currentWorksheet.Cells[i, 7].Value.ToString());
                                rw.swift = (currentWorksheet.Cells[i, 8].Value == null ? string.Empty : currentWorksheet.Cells[i, 8].Value.ToString());
                                rw.isactive = (currentWorksheet.Cells[i, 9].Value == null ? string.Empty : currentWorksheet.Cells[i, 9].Value.ToString());
                                rw.status = (currentWorksheet.Cells[i, 10].Value == null ? string.Empty : currentWorksheet.Cells[i, 10].Value.ToString());
                                yield return rw;
                            }
                        }
                        else
                        {
                            throw new Exception(string.Format("Похоже в книге нет листа '{0}' (в нем должны быть данные).", sheetName));
                        }
                    }
                    else
                    {
                        throw new Exception(string.Format("В книге мало листов"));
                    }
                }
                else
                {
                    throw new Exception(string.Format("С книгой что-то не так"));
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
            () => {  });
        }
        private void Upload(IList<SlxRow> items)
        {
            string cstr;
            if (isTestMode)
            {
                cstr= ConfigurationManager.ConnectionStrings["TestString"].ConnectionString;
            }
            else
            {
                cstr = ConfigurationManager.ConnectionStrings["RealString"].ConnectionString;
            }
            StringBuilder query = new StringBuilder("Insert into z_AllClientViaSLX  (slxCode, ogrn, inn, slx_client_name, category, departament, upravlenie, swift, isactive, status) " +
                            "values ");
            using (MySqlConnection con = new MySqlConnection(cstr))
            {
                List<string> Rows = new List<string>();
                foreach (var item in items)
                {

                    Rows.Add(string.Format("('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}' )",
                        MySqlHelper.EscapeString(item.slxCode),
                        MySqlHelper.EscapeString(item.ogrn),
                        MySqlHelper.EscapeString(item.inn),
                        MySqlHelper.EscapeString(item.name),
                        MySqlHelper.EscapeString(item.category),
                        MySqlHelper.EscapeString(item.department),
                        MySqlHelper.EscapeString(item.upravlenie),
                        MySqlHelper.EscapeString(item.swift),
                        MySqlHelper.EscapeString(item.isactive),
                        MySqlHelper.EscapeString(item.status)
                        ));

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
            StringBuilder query = new StringBuilder("Delete from z_AllClientViaSLX;");
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
}
