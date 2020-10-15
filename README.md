				var host = new HostBuilder()
                .ConfigureServices((hostContext, services) =>
                {
                    //注册后台普通服务
                    services.AddExcelOperator();
                })
                .ConfigureLogging((hostContext, configLogging) =>
                {
                    configLogging.AddConsole();
                    configLogging.AddDebug();
                })
                .UseConsoleLifetime()
                .Build();

            // 获取服务
            var oExcel = host.Services.GetRequiredService<IExcelOperator>();
            //===================================导入  水位-库容-面积曲线模版

             //动态映射EXCEL列
            var filePath1 = Path.Combine(Directory.GetCurrentDirectory(), "水位-库容-面积曲线模版.xlsx");
            using (var stream = new FileStream(filePath1, FileMode.Open))
            {
                var unused = await oExcel.ImportAsync<ReservoirWaterLevelStorageCurve>(stream, 5, 
                    null
                    ,"ZeroArea", "ZeroStorageNumber"
                    , "OneArea", "OneStorageNumber"
                    );
                Console.WriteLine(unused.Data.Count());
            }
            Console.ReadLine();
            //动态起始列头设置

            //var filePath = Path.Combine(Directory.GetCurrentDirectory(), "国控二期水位流量关系曲线数值2.xlsx");
            //using (var stream = new FileStream(filePath, FileMode.Open))
            //{
            //    var result = await oExcel.ReadHydrologicalRelationCurveAsync<ImportWatermark>(stream,3);
            //}
            ////====================================导出
            //var filePath2 = Path.Combine(Directory.GetCurrentDirectory(), "产品单sheel导出模板.xlsx");
            //var list = new List<ExporterProductDto>()
            //    {
            //        new ExporterProductDto { Name="2", Code="0000000002", BarCode="《XX从入门到放弃》", Type=0},
            //        new ExporterProductDto { Name="3", Code="0000000003", BarCode="《XX从入门到放弃》", Type=0}
            //    };
            //// 设置需要的字段
            //var heads = new List<ExcelHeader>();
            //heads.Add(new ExcelHeader {Code = "Name", DisplayName = "自定义名称"});
            //heads.Add(new ExcelHeader { Code = "Code", DisplayName = "自定义编码" });
            //var reult = await oExcel.ExportFileAsync(list, filePath2);