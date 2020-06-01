using ExcelUtil.Test.Api.Dto;
using System.Collections.Generic;

namespace ExcelUtil.Test.Api.DataCache
{
    public class ListCache
    {
        /// <summary>
        /// 初始化数据
        /// </summary>
        static ListCache()
        {
            DepartmentCache.Add(new DepartmentDto()
            {
                DepartmentID2 = 1,
                Name = "软件",
                Desc = "完成软件设计和研发"
            });

            DepartmentCache.Add(new DepartmentDto()
            {
                DepartmentID2 = 1,
                Name = "市场",
                Desc = "完成市场开拓和客户维护"
            });
        }

        /// <summary>
        /// 部门
        /// </summary>
        public static List<DepartmentDto> DepartmentCache = new List<DepartmentDto>();

        /// <summary>
        /// 雇员
        /// </summary>
        public static List<EmployeeDto> EmployeeCache = new List<EmployeeDto>();
    }
}