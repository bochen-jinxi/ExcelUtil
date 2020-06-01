using ExcelUtil.Test.Api.DataCache;
using ExcelUtil.Test.Api.Dto;
using Magicodes.ExporterAndImporter.Excel;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelUtil.Test.Api.Controllers
{
    /// <summary>
    /// 部门管理
    /// </summary>
    [ApiController]
    [Route("[controller]")]
    public class DepartmentAndEmployeeController : ControllerBase
    {
        public IExcel oIExcel { get; }

        /// <summary>
        /// 测试
        /// </summary>
        public DepartmentAndEmployeeController(IExcel importer)
        {
            oIExcel = importer;
        }

        /// <summary>
        ///
        /// </summary>
        [HttpGet("GenerateImportTemplate")]
        public async void GenerateImportTemplate()
        {
            await oIExcel.GenerateImportTemplateAsync<EmployeeDto>(Path.Combine(Environment.CurrentDirectory, "DepartmentAndEmployee", "template.xlsx"));
        }

        /// <summary>
        /// Excel导入1
        /// </summary>
        /// <returns></returns>
        [HttpGet("Import")]
        public async Task<object> Import()
        {
            var result = await oIExcel.ImportAsync<DepartmentDto>(Path.Combine(Environment.CurrentDirectory, "DepartmentAndEmployee", "import.xlsx"));
            return result;//json错误-解决循环引用
        }

        /// <summary>
        /// Excel导入2.1
        /// </summary>
        /// <returns></returns>
        [HttpGet("ImportMultipleSheet")]
        public async Task<object> ImportMultipleSheet()
        {
            var result = await oIExcel.ImportMultipleSheetAsync<Table1>(Path.Combine(Environment.CurrentDirectory, "DepartmentAndEmployee", "import.xlsx"));
            return result;
        }

        /// <summary>
        /// Excel导入2.2
        /// 导入转换到数据字典数据字典
        /// </summary>
        /// <returns></returns>
        [HttpGet("ImportMultipleSheetLightAsync")]
        public async Task<object> ImportMultipleSheetLightAsync()
        {
            var result = await oIExcel.ImportMultipleSheetLightAsync<Table1>(Path.Combine(Environment.CurrentDirectory, "DepartmentAndEmployee", "import.xlsx"));

            #region 示例取值

            var ret1 = result.Get<DepartmentDto>("部门数据");
            var ret2 = result.Get<EmployeeDto>("临时工");
            ListCache.DepartmentCache.AddRange(ret1);
            ListCache.EmployeeCache.AddRange(ret2);

            #endregion 示例取值

            return result;
        }

        /// <summary>
        /// Excel导入3
        /// </summary>
        /// <returns></returns>
        [HttpGet("ImportSameSheets")]
        public async Task<object> ImportSameSheets()
        {
            var result = await oIExcel.ImportSameSheetsAsync<Table2, EmployeeDto>(Path.Combine(Environment.CurrentDirectory, "DepartmentAndEmployee", "import.xlsx"));
            return result;
        }
    }
}