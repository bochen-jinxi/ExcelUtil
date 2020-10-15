namespace ExcelUtil
{
    /// <summary>
    /// Excel 表头
    /// </summary>
    public class ExcelHeader
    {
        public ExcelHeader() { }

        public ExcelHeader(string code, string displayName)
        {
            Code = code;
            DisplayName = displayName;
        }

        /// <summary>
        /// 表头字段编码
        /// </summary>
        public string Code { get; set; }
        /// <summary>
        /// 表头字段显示名
        /// </summary>
        public string DisplayName { get; set; }
    }



}
