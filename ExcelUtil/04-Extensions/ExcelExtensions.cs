using ExcelUtil;
using Magicodes.ExporterAndImporter.Excel;

namespace Microsoft.Extensions.DependencyInjection
{
    /// <summary>
    /// Excel 静态扩展
    /// </summary>
    public static class ExcelExtensions
    {
        /// <summary>
        /// 注册 Excel 服务
        /// </summary>
        /// <param name="services">ServiceCollection</param>
        /// <returns></returns>
        public static IServiceCollection AddExcelOperator(this IServiceCollection services)
        {
            services.AddScoped<IExcelOperator, ExcelOperator>();
            services.AddScoped<IExcelImporter, ExcelImporter>();
            services.AddScoped<IExcelExporter, ExcelExporter>();
            return services;
        }
    }
}