using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;

namespace ExcelUtil._06_Model
{

    /// <summary>
    ///     水位线
    /// </summary>
    [ExcelImporter(HeaderRowIndex = 0)]
    public class ImportWatermark
    {
        /// <summary>
        ///     流     水位(厘米)
        /// </summary>
        [ImporterHeader(Name = " 流     水位(厘米)")]
        public string Watermark { get; set; }

        /// <summary>
        ///     0
        /// </summary>
        [ImporterHeader(Name = "0")]
        public string A { get; set; }

        /// <summary>
        ///     1
        /// </summary>
        [ImporterHeader(Name = "1")]
        public string B { get; set; }

        /// <summary>
        ///     2
        /// </summary>
        [ImporterHeader(Name = "2")]
        public string C { get; set; }

        /// <summary>
        ///     3
        /// </summary>
        [ImporterHeader(Name = "3")]
        public string D { get; set; }

        /// <summary>
        ///     4
        /// </summary>
        [ImporterHeader(Name = "4")]
        public string E { get; set; }

        /// <summary>
        ///     5
        /// </summary>
        [ImporterHeader(Name = "5")]
        public string F { get; set; }

        /// <summary>
        ///     6
        /// </summary>
        [ImporterHeader(Name = "6")]
        public string G { get; set; }

        /// <summary>
        ///     7
        /// </summary>
        [ImporterHeader(Name = "7")]
        public string H { get; set; }

        /// <summary>
        ///     8
        /// </summary>
        [ImporterHeader(Name = "8")]
        public string I { get; set; }

        /// <summary>
        ///     9
        /// </summary>
        [ImporterHeader(Name = "9")]
        public string J { get; set; }
    }
}