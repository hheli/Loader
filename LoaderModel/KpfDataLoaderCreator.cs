using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoaderModel
{
    public class KpfDataLoaderCreator : XlDataLoaderCreator
    {
        public override XlDataLoader CreateXlDataLoader(string Path, bool Istestmode, string SheetName)
        {
            return new KpfDataLoader( Path,  Istestmode,  SheetName);
        }
    }
}
