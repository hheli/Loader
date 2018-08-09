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
    public class GroupsLoader
    {

        public class GroupRow
        {
            public string grName { set; get; }
            public string grNum { set; get; }
            public string clName { set; get; }
            public string Inn { set; get; }
            public string OGRN { set; get; }
            public string SLX{ set; get; }
            public string dscrtpn{ get; set; }
        }
        public string filePath { set; get; }
        public bool isTestMode { set; get; }
        public GroupsLoader(string Path, bool istest)
        {
            filePath = Path;
            isTestMode = istest;
            DeleteData();
        }

        public IEnumerable<GroupRow> ReadFirst()
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
                        var currentWorksheet = workBook.Worksheets["Перечень групп"];
                        if (currentWorksheet != null)
                        {
                            // read data from file
                            var start = currentWorksheet.Dimension.Start;
                            var end = currentWorksheet.Dimension.End;
                            string grname = "";
                            for (int i = start.Row + 8; i <= end.Row; i++) //начинаем со строки 8
                            {
                                var rw = new GroupRow();
                                rw.grName = ((currentWorksheet.Cells[i, 3].Value == null) || (currentWorksheet.Cells[i,3].Value.ToString()=="") ?  grname : grname = currentWorksheet.Cells[i, 3].Value.ToString());
                                rw.grNum = (currentWorksheet.Cells[i, 2].Value == null ? string.Empty : currentWorksheet.Cells[i, 2].Value.ToString());
                                rw.clName = (currentWorksheet.Cells[i, 5].Value == null ? string.Empty : currentWorksheet.Cells[i, 5].Value.ToString());
                                rw.Inn = (currentWorksheet.Cells[i, 6].Value == null ? string.Empty : currentWorksheet.Cells[i, 6].Value.ToString());
                                rw.OGRN = (currentWorksheet.Cells[i, 7].Value == null ? string.Empty : currentWorksheet.Cells[i, 7].Value.ToString());
                                rw.dscrtpn = (currentWorksheet.Cells[i, 8].Value == null ? string.Empty : currentWorksheet.Cells[i, 8].Value.ToString());
                                rw.SLX = (currentWorksheet.Cells[i, 11].Value == null ? string.Empty : currentWorksheet.Cells[i, 11].Value.ToString());
                                yield return rw;
                            }
                        }
                        else
                        {
                            throw new Exception(string.Format("Не обнаружен лист с названием 'Перечень групп'"));
                        }
                    }
                }
            }
        }

        public IEnumerable<GroupRow> ReadSecond()
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
                        var currentWorksheet = workBook.Worksheets["Связанные с Банком"];
                        if (currentWorksheet != null)
                        {
                            // read data from file
                            var start = currentWorksheet.Dimension.Start;
                            var end = currentWorksheet.Dimension.End;
                            for (int i = start.Row + 8; i <= end.Row; i++) //начинаем со строки 8
                            {
                                var rw = new GroupRow();
                                rw.grName =  "VTB";
                                rw.grNum = "0000000";
                                rw.clName = (currentWorksheet.Cells[i, 2].Value == null ? string.Empty : currentWorksheet.Cells[i, 2].Value.ToString());
                                rw.Inn = (currentWorksheet.Cells[i, 3].Value == null ? string.Empty : currentWorksheet.Cells[i, 3].Value.ToString());
                                rw.OGRN = (currentWorksheet.Cells[i, 4].Value == null ? string.Empty : currentWorksheet.Cells[i, 4].Value.ToString());
                                rw.dscrtpn = (currentWorksheet.Cells[i, 5].Value == null ? string.Empty : currentWorksheet.Cells[i, 5].Value.ToString());
                                rw.SLX = (currentWorksheet.Cells[i, 6].Value == null ? string.Empty : currentWorksheet.Cells[i, 6].Value.ToString());
                                yield return rw;
                            }
                        }
                        else
                        {
                            throw new Exception(string.Format("Не обнаружен лист с названием 'Связанные с Банком'"));
                        }
                    }
                }
            }
        }
        public void UploadToDb()
        {

            ReadFirst().ToObservable().
            Buffer(1000).Subscribe(loadedData =>
            {
                Upload(loadedData);
            },
                exception => { MessageBox.Show(exception.Message); },
                () => { });
            ReadSecond().ToObservable().
            Buffer(1000).Subscribe(loadedData =>
            {
                Upload(loadedData);
            },
            exception => { MessageBox.Show(exception.Message); },
            () => { });
        }
        private void Upload(IList<GroupRow> items)
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
            StringBuilder query = new StringBuilder("Insert into z_groups  (grNum, grName, clientName, inn, ogrn, description, slxCode) " +
                            "values ");
            using (MySqlConnection con = new MySqlConnection(cstr))
            {
                List<string> Rows = new List<string>();
                foreach (var item in items)
                {

                    Rows.Add(string.Format("('{0}','{1}','{2}','{3}','{4}','{5}','{6}')",
                        MySqlHelper.EscapeString(item.grNum),
                        MySqlHelper.EscapeString(item.grName),
                        MySqlHelper.EscapeString(item.clName),
                        MySqlHelper.EscapeString(item.Inn),
                        MySqlHelper.EscapeString(item.OGRN),
                        MySqlHelper.EscapeString(item.dscrtpn),
                        MySqlHelper.EscapeString(item.SLX)));
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
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
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
            StringBuilder query = new StringBuilder("Delete from z_groups;");
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
