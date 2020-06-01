using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using System;
using System.ComponentModel.DataAnnotations;

namespace ExcelUtil.Test.Api.Dto
{
    /// <summary>
    /// 表1
    /// 1.表格对应Sheet名称，一定要对，不然报错【Object reference not set to an instance of an object.】
    /// 2.Excel独占打开
    /// 3.输出Sheet名，用户可以自定义
    /// </summary>
    public class Table1
    {
        /// <summary>
        /// Sheet
        /// </summary>
        [ExcelImporter(IsLabelingError = true, SheetName = "部门数据")]//todo:特性类型不能可空类型, EndColumnCount = 2，, HeaderRowIndex = 2未生效
        public DepartmentDto Sheet1 { get; set; }

        /// <summary>
        /// Sheet
        /// </summary>
        [ExcelImporter(IsLabelingError = true, SheetName = "编制员工")]
        public EmployeeDto Sheet2 { get; set; }

        /// <summary>
        /// Sheet
        /// </summary>
        [ExcelImporter(IsLabelingError = true, SheetName = "临时工")]
        public EmployeeDto Sheet3 { get; set; }
    }

    /// <summary>
    /// 表2
    /// </summary>
    public class Table2
    {
        /// <summary>
        /// Sheet
        /// </summary>
        [ExcelImporter(IsLabelingError = true, SheetName = "编制人员")]
        public EmployeeDto Sheet2 { get; set; }

        /// <summary>
        /// Sheet
        /// </summary>
        [ExcelImporter(IsLabelingError = true, SheetName = "临时工")]
        public EmployeeDto Sheet1 { get; set; }
    }
}