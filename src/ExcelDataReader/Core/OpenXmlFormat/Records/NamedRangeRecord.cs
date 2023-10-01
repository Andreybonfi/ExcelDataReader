using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class NamedRangeRecord : Record
    {
        public NamedRangeRecord(NamedRange range) 
        {
            Range = range;
        }

        public NamedRange Range { get; }
    }
}
