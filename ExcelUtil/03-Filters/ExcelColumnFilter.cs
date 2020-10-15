using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelUtil
{
    /// <summary>
    /// Excel 表头字段过滤器
    /// </summary>
    internal class ExcelHeaderFilter : IExporterHeaderFilter
    {
        public ExcelHeaderFilter(List<ExcelHeader> displayHeaders)
        {
            DisplayHeaders = displayHeaders;
        }

        private List<ExcelHeader> DisplayHeaders { get; }

        /// <summary>
        /// 表头筛选器（修改名称）
        /// </summary>
        /// <param name="exporterHeaderInfo"></param>
        /// <returns></returns>
        public ExporterHeaderInfo Filter(ExporterHeaderInfo exporterHeaderInfo)
        {
            if (!DisplayHeaders.Any()) return exporterHeaderInfo;
            foreach (var displayHeader in DisplayHeaders)
            {
                if (displayHeader.Code.Equals(exporterHeaderInfo.PropertyName, StringComparison.OrdinalIgnoreCase))
                {
                    exporterHeaderInfo.DisplayName = displayHeader.DisplayName;
                    exporterHeaderInfo.ExporterHeaderAttribute.IsIgnore = false;
                    break;
                }
                exporterHeaderInfo.ExporterHeaderAttribute.IsIgnore = true;
            }
            return exporterHeaderInfo;
        }
    }
}
