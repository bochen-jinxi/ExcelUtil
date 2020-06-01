using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using ExcelUtil.Model;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace ExcelUtil.Test
{
    internal class Program
    {
        private static async Task Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var host = new HostBuilder()

               .ConfigureServices((hostContext, services) =>
               {
                   //注册后台普通服务
                   services.AddExcelService();
               })
               .ConfigureLogging((hostContext, configLogging) =>
               {
                   configLogging.AddConsole();
                   configLogging.AddDebug();
               })
               .UseConsoleLifetime()
               .Build();

            // 获取服务
            var oExcel = host.Services.GetRequiredService<IExcel>();

            #region 导出模板

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "产品生成模板.xlsx");
            ExportFileInfo template = await oExcel.GenerateTemplateAsync<ExporterProductDto>(filePath);
            //返回的文件名
            string oFileName = template.FileName;
            //返回的文件类型
            string oFileType = template.FileType;

            //生成字节数据
            byte[] bytes = await oExcel.GenerateTemplateBytesAsync<ExporterProductDto>();
            //生成流
            Stream stream = await oExcel.GenerateTemplateStreamAsync<ExporterProductDto>();

            #endregion 导出模板

            #region 生成一个excel一个Sheel

            var filePath2 = Path.Combine(Directory.GetCurrentDirectory(), "产品单sheel导出模板.xlsx");
            var list = new List<ExporterProductDto>()
            {
                new ExporterProductDto { Name="2", Code="0000000002", BarCode="《XX从入门到放弃》", Type=0},
                new ExporterProductDto { Name="3", Code="0000000003", BarCode="《XX从入门到放弃》", Type=0},
            };
            var list2 = new List<ExporterProductDto2>()
            {
                new ExporterProductDto2 { Name="2", Code="0000000002", BarCode="《XX从入门到放弃》", Type=0},
                new ExporterProductDto2 { Name="3", Code="0000000003", BarCode="《XX从入门到放弃》", Type=0},
            };
            var list3 = new List<ExporterProductDto3>()
            {
                new ExporterProductDto3 { Name="2", Code="0000000002", BarCode="《XX从入门到放弃》", Type=0},
                new ExporterProductDto3 { Name="3", Code="0000000003", BarCode="《XX从入门到放弃》", Type=0},
            };
            var reult2 = await oExcel.ExportSingSheelAsync(filePath2, list);

            //导出二进制
            var sb = await oExcel.ExportSingSheelAsByteAsync(list);

            #endregion 生成一个excel一个Sheel

            #region 生成一个excel多个Sheel 最多5个sheel

            var filePath4 = Path.Combine(Directory.GetCurrentDirectory(), "产品2不同sheel导出模板.xlsx");
            var reult4 = oExcel.ExportMultSheelAsync<ExporterProductDto, ExporterProductDto2>(filePath4, list, list2);
            ////导出二进制
            var reult5 = oExcel.ExportMultSheelAsByteAsync<ExporterProductDto, ExporterProductDto2>(list, list2);
            var reult45 = oExcel.ExportMultSheelAsByteAsync<ExporterProductDto, ExporterProductDto2>(list, list2);
            var filePath6 = Path.Combine(Directory.GetCurrentDirectory(), "产品3不同sheel导出模板.xlsx");
            var reult6 = oExcel.ExportMultSheelAsync<ExporterProductDto, ExporterProductDto2, ExporterProductDto3>(filePath6, list, list2, list3);
            if (reult6.Exception != null)
            {
                Console.WriteLine("导出有错");
            }

            #endregion 生成一个excel多个Sheel 最多5个sheel

            #region 依赖模板导出数据

            //导出模板字节数据
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "2020年春季教材订购明细样表.xlsx");
            var data = new TextbookOrderInfo("湖南心莱信息科技有限公司", "湖南长沙岳麓区", "雪雁", "1367197xxxx", null,
                    DateTime.Now.ToLongDateString(), "https://docs.microsoft.com/en-us/media/microsoft-logo-dark.png",
                    new List<BookInfo>()
                    {
                          new BookInfo(1, "0000000001", "《XX从入门到放弃》", null, "机械工业出版社", "3.14", 100, "备注")
                          {
                              Cover = Path.Combine("TestFiles", "ExporterTest.png")
                          },
                          new BookInfo(2, "0000000002", "《XX从入门到放弃》", "张三", "机械工业出版社", "3.14", 100, null),
                          new BookInfo(3, null, "《XX从入门到放弃》", "张三", "机械工业出版社", "3.14", 100, "备注")
                          {
                              Cover = Path.Combine("TestFiles", "ExporterTest.png")
                          }
                    });
            var byts = oExcel.ExportSingByTemplateAsByteAsync<TextbookOrderInfo>(data, tplPath);

            var filePath9 = Path.Combine(Directory.GetCurrentDirectory(), "导出模板数据.xlsx");
            var result = await oExcel.ExportSingByTemplateAsync(filePath9, data, tplPath);

            #endregion 依赖模板导出数据

            host.Run();
        }
    }
}