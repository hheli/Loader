using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoaderModel
{
    public class MBankDataLoaderCreator : XlDataLoaderCreator
    {
        public override XlDataLoader CreateXlDataLoader(string Path, bool Istest, string SheetName)
        {
            return new MBankDataLoader(Path, Istest, SheetName);
        }
    }
}
