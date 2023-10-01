using System;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelDataReader.Core;

namespace ExcelDataReader
{
    /// <summary>
    /// A range for cells using 0 index positions. 
    /// </summary>
    public sealed class NamedRange
    {
        private Regex _regex = new Regex(@"(?<sheetname>[\w|\W]+!)(?<first>[\w|\W]+\d:|[\w|\W]+\d)(?<second>[\w|\W]*\d|)", RegexOptions.Compiled);

        internal NamedRange() { }

        internal NamedRange(string name, string range , bool global)
        {
            Match match = _regex.Match(range);
            
            SheetName = !string.IsNullOrEmpty(match.Groups["sheetname"].Value) ? match.Groups["sheetname"].Value.Replace("!", string.Empty) : null;
            RangeName = name;
            Global = global;

            if (!string.IsNullOrEmpty(match.Groups["first"].Value))
            {
                var first = Regex.Replace(match.Groups["first"].Value, @"\$|:", string.Empty);
                ReferenceHelper.ParseReference(first, out int column, out int row);

                // 0 indexed vs 1 indexed
                FromColumn = column - 1;
                FromRow = row - 1;

                if (!string.IsNullOrEmpty(match.Groups["second"].Value))
                {
                    ReferenceHelper.ParseReference(match.Groups["second"].Value.Replace("$", string.Empty), out column, out row);

                    // 0 indexed vs 1 indexed
                    ToColumn = column - 1;
                    ToRow = row - 1;
                }
            }
        }

        internal NamedRange(string name,bool global, string sheetName, int fromColumn, int fromRow, int toColumn, int toRow)
        {
            RangeName = name;
            SheetName = sheetName;
            FromColumn = fromColumn;
            FromRow = fromRow;
            ToColumn = toColumn;
            ToRow = toRow;
            Global = global;
        }


        /// <summary>
        /// Gets the Sheet where range located. Null if range global.
        /// </summary>
        public string SheetName { get; }

        /// <summary>
        /// Gets the range name.
        /// </summary>
        public string RangeName { get; }

        /// <summary>
        /// Return true if named range located in Workbook name splace. 
        /// </summary>
        public bool Global { get; }

        /// <summary>
        /// Gets the column the range starts in.
        /// </summary>
        public int FromColumn { get; }

        /// <summary>
        /// Gets the row the range starts in.
        /// </summary>
        public int FromRow { get; }

        /// <summary>
        /// Gets the column the range ends in.
        /// </summary>
        public int? ToColumn { get; }

        /// <summary>
        /// Gets the row the range ends in.
        /// </summary>
        public int? ToRow { get; }

        /// <inheritsdoc/>
        public override string ToString() => $"{SheetName},{RangeName}, {FromRow}, {ToRow}, {FromColumn}, {ToColumn}, {Global}";
    }
}