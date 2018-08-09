using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoaderModel
{
    public abstract class XlDataLoaderCreator
    {
        public abstract XlDataLoader CreateXlDataLoader(string Path, bool Istestmode, string SheetName);
    }
}
