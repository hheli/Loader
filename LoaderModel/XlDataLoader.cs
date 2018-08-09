using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reactive.Linq;
using System.Windows;

namespace LoaderModel
{
    abstract public class XlDataLoader
    {
        public string filePath { set; get; }
        public bool isTestMode { set; get; }
        public string sheetName { set; get; }
        public XlDataLoader(string Path, bool Istest, string SheetName )
        {
            filePath = Path;
            isTestMode = Istest;
            sheetName = SheetName;
        }
        public  abstract IEnumerable<XlRow> ReadXlsx();

        public abstract void Upload(IList<XlRow> items);

        public  void UploadToDb()
        {
           ReadXlsx().ToObservable().
           Buffer(1000).Subscribe(loadedData =>
           {
               Upload(loadedData);
           },
               exception => { MessageBox.Show(exception.Message); },
               () => { });
        }
    }
}
