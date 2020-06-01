using Magicodes.ExporterAndImporter.Core;

namespace ExcelUtil.Test.Api.Dto
{
    /// <summary>
    /// 雇员
    /// </summary>
    public class EmployeeDto
    {
        /// <summary>
        /// 雇员编号
        /// </summary>
        [ImporterHeader(Name = "雇员编号")]
        public int EmployeeID { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        [ImporterHeader(Name = "姓名")]
        public string Name { get; set; }

        /// <summary>
        /// 隶属部门
        /// </summary>
        [ImporterHeader(Name = "隶属部门")]
        public int DeptID { get; set; }
    }
}