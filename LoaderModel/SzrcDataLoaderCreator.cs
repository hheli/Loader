using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoaderModel
{
    public class SzrcDataLoaderCreator : XlDataLoaderCreator
    {
        public override XlDataLoader CreateXlDataLoader(string Path, bool Istest, string SheetName)
        {
            return new SzrcDataLoader(Path, Istest, SheetName);
        }
    }
}
