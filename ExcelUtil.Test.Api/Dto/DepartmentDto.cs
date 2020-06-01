using Magicodes.ExporterAndImporter.Core;
using System;

namespace ExcelUtil.Test.Api.Dto
{
    /// <summary>
    /// 部门
    /// </summary>
    public class DepartmentDto
    {
        /// <summary>
        /// 部门编号
        /// </summary>
        [ImporterHeader(Name = "部门编号")]
        public int DepartmentID2 { get; set; }

        /// <summary>
        /// 部门
        /// </summary>
        [ImporterHeader(Name = "部门")]
        public string Name { get; set; }

        /// <summary>
        /// 简介
        /// </summary>
        [ImporterHeader(Name = "简介")]
        public string Desc { get; set; }

        /// <summary>
        /// 成立时间
        /// </summary>
        [ImporterHeader(Name = "成立时间")]
        public DateTime DT { get; set; }
    }
}