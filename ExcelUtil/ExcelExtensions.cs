using Magicodes.ExporterAndImporter.Excel;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelUtil
{
    public static class ExcelExtensions
    {
        /// <summary>
        /// 注册Excel服务
        /// </summary>
        /// <param name="services"></param>
        /// <returns></returns>
        public static IServiceCollection AddExcelService(this IServiceCollection services)
        {
            services.AddScoped<IExcel, Excel>();
            services.AddScoped<IExcelImporter, ExcelImporter>();
            services.AddScoped<IExcelExporter, ExcelExporter>();

            return services;
        }
    }
}