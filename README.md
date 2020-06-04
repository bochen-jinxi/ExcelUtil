# ExcelUtil
基于Magicodes.IE.Excel，操作Excel导入导出
使用步骤


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

            var filePath2 = Path.Combine(Directory.GetCurrentDirectory(), "产品单sheel导出模板.xlsx");
            var list = new List<ExporterProductDto>()
            {
                new ExporterProductDto { Name="2", Code="0000000002", BarCode="《XX从入门到放弃》", Type=0},
                    new ExporterProductDto { Name="3", Code="0000000003", BarCode="《XX从入门到放弃》", Type=0}
            };
              /设置需要的字段
            var dic = new Dictionary<string, string>(){ { "name","名称3"},{"code","编码3"} }; 
            //ExporterProductDto.Filter(dic);
             var reult = await oExcel.ExportSingSheelAsync(filePath2, list);
