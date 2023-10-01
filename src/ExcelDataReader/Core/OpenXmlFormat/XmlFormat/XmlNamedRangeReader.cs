using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal sealed class XmlNamedRangeReader : XmlRecordReader
    {
        private const string NamedRangeGroup = "definedNames";
        private const string NamedRange = "definedName";
        private const string LocalSheetId = "localSheetId";
        private const string NamOfRange = "name";


        public XmlNamedRangeReader(XmlReader reader) 
            : base(reader)
        {
        }

        protected override IEnumerable<Record> ReadOverride()
        {
            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(NamedRangeGroup, ProperNamespaces.NsSpreadsheetMl))
                {                   
                    foreach (var namedrange in ReadNamedRange(ProperNamespaces.NsSpreadsheetMl))
                         yield return new NamedRangeRecord(namedrange);
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }
        }

        protected IEnumerable<NamedRange> ReadNamedRange(string nsSpreadsheetMl)
        { 
            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }
            
            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(NamedRange, nsSpreadsheetMl))
                {
                    var name = Reader.GetAttribute(NamOfRange);
                    bool global = Reader.GetAttribute(LocalSheetId) == null;

                    XmlReaderHelper.ReadFirstContent(Reader);
                    var value = Reader.Value;

                    yield return new NamedRange(name, value, global);
                    XmlReaderHelper.SkipContent(Reader);
                }
                else if (XmlReaderHelper.IsEndElement(Reader, NamedRange, nsSpreadsheetMl))
                {
                    XmlReaderHelper.SkipContent(Reader);
                    continue;
                }
                else if (XmlReaderHelper.IsEndElement(Reader, NamedRangeGroup, nsSpreadsheetMl))
                {
                    break;
                }
            }
        }
    }
}
