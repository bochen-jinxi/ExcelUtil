using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using ExcelUtil._05_Utils;
using ExcelUtil._06_Model;
using ExcelUtil.Model;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Excel.Utility;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using RestSharp;
using SaiLing.LuXian.BaseData.Api.Services.ExportDtos.BasicInfo;

namespace ExcelUtil.Test
{

   
    internal class Program
    {
        private static async Task Main(string[] args)
        {
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
            //===================================导入  水位线



            //动态映射EXCEL列
            //var filePath1 = Path.Combine(Directory.GetCurrentDirectory(), "水位-库容-面积曲线模版.xlsx");
            //using (var stream = new FileStream(filePath1, FileMode.Open))
            //{
            //    var unused = await oExcel.ImportAsync<ReservoirWaterLevelStorageCurve>(stream, 5

            //    //x => new object[] {null, x.Altitude});
            //        , "ZeroArea", "ZeroStorageNumber"
            //        , "OneArea", "OneStorageNumber",
            //    "TwoArea", "TwoStorageNumber",
            //    "ThreeArea", "ThreeStorageNumber",
            //    "FourArea", "FourStorageNumber",
            //    "FiveArea", "FiveStorageNumber",
            //    "SixArea", "SixStorageNumber",
            //    "SevenArea", "SevenStorageNumber",
            //    "EightArea", "EightStorageNumber",
            //    "NineArea", "NineStorageNumber", null);
            //    Console.WriteLine(unused.Data.Count());
            //}
            //Console.ReadLine();
            // 动态起始列头设置

            //var filePath = Path.Combine(Directory.GetCurrentDirectory(), "国控二期水位流量关系曲线数值2.xlsx");
            //using (var stream = new FileStream(filePath, FileMode.Open))
            //{
            //    var result = await oExcel.ReadHydrologicalRelationCurveAsync<ImportWatermark>(stream, 3);
            //}



            var dicRule = new Dictionary<string, ArchiveTemplateDto>();

            var filePath00 = Path.Combine(Directory.GetCurrentDirectory(), "档案类型通用模板及字段.xlsx");
            using (var stream = new FileStream(filePath00, FileMode.Open))
            {
                var result = await oExcel.ImportAsync<ArchiveTemplateDto>(stream);
                foreach (var item in result.Data)
                {
                    dicRule.Add(item.Name100, item);
                }
            }


            var dic0 = new Dictionary<string, int>()
            {
                {"文件", 0},
                {"盒", 1},
                {"案卷", 2},
                {"项目", 3}
            };

            var dicType = new Dictionary<string, int>()
            {
                {"整型", 0},
                {"数字型", 1},
                {"字符型", 2},
                {"日期型", 3},
                {"时间型", 4},
                {"枚举型", 5},
                {"树型", 6},
                
            };

            var dicTF = new Dictionary<string, bool>()
            {
                {"是", true},
                {"否", false},
                {"有效", true},
                {"无效", false}
                

            };


            var dic02 = new Dictionary<string, int>();
            StringBuilder sb = new StringBuilder();
            StringBuilder sb2 = new StringBuilder();
            StringBuilder sb3 = new StringBuilder();
            var dic = new Dictionary<string, string>();
            var dic2 = new Dictionary<string, string>();
            


            var filePath0 = Path.Combine(Directory.GetCurrentDirectory(), "档案类型通用模板.xlsx");
            using (var stream = new FileStream(filePath0, FileMode.Open))
            {
                var result = await oExcel.ImportAsync<ArchiveTemplateDto>(stream);
                var index = 0;
                foreach (var item in result.Data)
                {

                    if (index < 2)
                    {

                       
                      
                        
                        //档案  
                        if (index == 0)
                        {
                            sb.Append(@"INSERT INTO ").Append("\"").Append("Document").Append("\" ").Append("(")
                                .Append("\"").Append("DocumentId").Append("\"").Append(", ")
                                .Append("\"").Append("DocumentCollectionId").Append("\"").Append(", ")
                                .Append("\"").Append("DocumenName").Append("\"").Append(", ")
                                .Append("\"").Append("BagName").Append("\"").Append(", ")
                                .Append("\"").Append("Sort").Append("\"").Append(", ")
                                .Append("\"").Append("CreationTime").Append("\"").Append(", ")
                                .Append("\"").Append("Creator").Append("\"").Append(", ")
                                .Append("\"").Append("CreatorId").Append("\"").Append(", ")
                                .Append("\"").Append("LastModificationTime").Append("\"").Append(", ")
                                .Append("\"").Append("LastModifierId").Append("\"").Append(", ")
                                .Append("\"").Append("Modifier").Append("\"").Append(", ")
                                 .Append("\"").Append("IsTemplate").Append("\"").Append(", ")
                                .Append("\"").Append("TemplateId").Append("\"").Append(", ")
                                .Append("\"").Append("IsDeleted").Append("\"").Append(", ")
                                .Append("\"").Append("Version").Append("\"")
                                .Append(") Values ");

                            var name100 = item.Name100;
                            var guid100 = Guid.NewGuid().ToString();
                            sb.AppendFormat(@"( '{0}', '00000000-0000-0000-0000-000000000000', '{1}',  '', 0, '{2}', 'sys', '00000000-0000-0000-0000-000000000000', '{2}', '00000000-0000-0000-0000-000000000000', 'sys',TRUE, '', FALSE,E'\\000\\000\\000' )", guid100, name100,DateTime.Now).Append(",");

                            Console.WriteLine(@"插入档案完毕");
                            dic.Add("100", guid100);

                            var name102 = item.Name102;
                            var guid102 = Guid.NewGuid().ToString();
                            sb.AppendFormat(@"( '{0}', '00000000-0000-0000-0000-000000000000', '{1}',  '', 0, '{2}', 'sys', '00000000-0000-0000-0000-000000000000', '{2}', '00000000-0000-0000-0000-000000000000', 'sys',TRUE, '', FALSE ,E'\\000\\000\\000')", guid102, name102, DateTime.Now).Append(",");
                            Console.WriteLine(@"插入档案完毕");
                            dic.Add("102", guid102);

                            var name104 = item.Name104;
                            var guid104 = Guid.NewGuid().ToString();
                            sb.AppendFormat(@"( '{0}', '00000000-0000-0000-0000-000000000000', '{1}',  '', 0, '{2}', 'sys', '00000000-0000-0000-0000-000000000000', '{2}', '00000000-0000-0000-0000-000000000000', 'sys',TRUE,  '',FALSE ,E'\\000\\000\\000')", guid104, name104, DateTime.Now).Append(",");
                            Console.WriteLine(@"插入档案完毕");
                            dic.Add("104", guid104);

                            var name105 = item.Name105;
                            var guid105 = Guid.NewGuid().ToString();
                            sb.AppendFormat(@"( '{0}', '00000000-0000-0000-0000-000000000000', '{1}',  '', 0, '{2}', 'sys', '00000000-0000-0000-0000-000000000000', '{2}', '00000000-0000-0000-0000-000000000000', 'sys',TRUE, '', FALSE,E'\\000\\000\\000' )", guid105, name105, DateTime.Now).Append(",");
                            Console.WriteLine(@"插入档案完毕");
                            dic.Add("105", guid105);

                            //var name108 = item.Name108;
                            //sb.AppendFormat(@"插入档案完毕");
                            //dic.Add("108", guid);

                            //var name109 = item.Name109;
                            //sb.AppendFormat(@"插入档案完毕");
                            //dic.Add("109", guid);

                            //var name110 = item.Name110;
                            //sb.AppendFormat(@"插入档案完毕");
                            //dic.Add("110", guid);


                            //var name112 = item.Name112;
                            //sb.AppendFormat(@"插入档案完毕");
                            //dic.Add("112", guid);

                            //var name115 = item.Name115;
                            //sb.AppendFormat(@"插入档案完毕");
                            //dic.Add("115", guid);

                            //var name118 = item.Name118;
                            //sb.AppendFormat(@"插入档案完毕");
                            //dic.Add("118", guid);

                            //var name120 = item.Name120;
                            //sb.AppendFormat(@"插入档案完毕");
                            //dic.Add("120", guid);


                            //var name122 = item.Name122;
                            //sb.AppendFormat(@"插入档案完毕");
                            //dic.Add("122", guid);

                            //var name123 = item.Name123;
                            //sb.AppendFormat(@"插入档案完毕");
                            //dic.Add("123", guid);

                            //var name126 = item.Name126;
                            //sb.AppendFormat(@"插入档案完毕");
                            //dic.Add("126", guid);

                            //var name127 = item.Name127;
                            //sb.AppendFormat(@"插入档案完毕");
                            //dic.Add("127", guid);

                            Console.WriteLine(sb.ToString());
                        }
                        else if (index == 1)
                        {
                            //盒 文件

                            sb2.Append(@"INSERT INTO ").Append("\"").Append("DocumentType").Append("\" ").Append("(")
                                .Append("\"").Append("DocumentTypeId").Append("\"").Append(", ")
                                .Append("\"").Append("DocumentId").Append("\"").Append(", ")
                                .Append("\"").Append("DocumentTypeName").Append("\"").Append(", ")
                                .Append("\"").Append("DocumentTypeValue").Append("\"").Append(", ")
                                .Append("\"").Append("Sort").Append("\"").Append(", ")
                                .Append("\"").Append("CreationTime").Append("\"").Append(", ")
                                .Append("\"").Append("Creator").Append("\"").Append(", ")
                                .Append("\"").Append("CreatorId").Append("\"").Append(", ")
                                .Append("\"").Append("LastModificationTime").Append("\"").Append(", ")
                                .Append("\"").Append("LastModifierId").Append("\"").Append(", ")
                                .Append("\"").Append("Modifier").Append("\"").Append(", ")
                                .Append("\"").Append("IsDeleted").Append("\"").Append(", ")
                                .Append("\"").Append("Version").Append("\"")
                                .Append(") Values ");

                            var name100 = item.Name100;
                            var guid100 = Guid.NewGuid().ToString();
                            sb2.AppendFormat(
                                @"('{0}', '{1}', '{2}', '{3}', 0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE,E'\\000\\000\\000')",
                                guid100, dic["100"], name100, dic0[name100.Trim()]).Append(", ");
                            Console.WriteLine(@"插入盒完毕",dic["100"]);
                            dic2.Add("100", guid100);
                            dic02.Add("100", dic0[name100.Trim()]);

                            var name101 = item.Name101;
                            var guid101 = Guid.NewGuid().ToString();
                            sb2.AppendFormat(
                                @"('{0}', '{1}', '{2}', '{3}', 0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE,E'\\000\\000\\000')",
                                guid101, dic["100"], name101, dic0[name101.Trim()]).Append(", ");
                            Console.WriteLine(@"插入盒完毕", dic["100"]);
                            dic2.Add("101", guid101);
                            dic02.Add("101", dic0[name101.Trim()]);

                            var name102 = item.Name102;
                            var guid102 = Guid.NewGuid().ToString();
                            sb2.AppendFormat(
                                @"('{0}', '{1}', '{2}', '{3}', 0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE,E'\\000\\000\\000')",
                                guid102, dic["102"], name102, dic0[name102.Trim()]).Append(", ");
                            Console.WriteLine(@"插入盒完毕", dic["102"]);
                            dic2.Add("102", guid102);
                            dic02.Add("102", dic0[name102.Trim()]);

                            var name103 = item.Name103;
                            var guid103 = Guid.NewGuid().ToString();
                            sb2.AppendFormat(
                                @"('{0}', '{1}', '{2}', '{3}', 0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE,E'\\000\\000\\000')",
                                guid103, dic["102"], name103, dic0[name103.Trim()]).Append(", ");
                            Console.WriteLine(@"插入文件完毕", dic["102"]);
                            dic2.Add("103", guid103);
                            dic02.Add("103", dic0[name103.Trim()]);

                            var name104 = item.Name104;
                            var guid104 = Guid.NewGuid().ToString();
                            sb2.AppendFormat(
                                @"('{0}', '{1}', '{2}', '{3}', 0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE,E'\\000\\000\\000')",
                                guid104, dic["104"], name104, dic0[name104.Trim()]).Append(", ");
                            Console.WriteLine(@"插入盒完毕", dic["104"]);
                            dic2.Add("104", guid104);
                            dic02.Add("104", dic0[name104.Trim()]);

                            var name105 = item.Name105;
                            var guid105 = Guid.NewGuid().ToString();
                            sb2.AppendFormat(
                                @"('{0}', '{1}', '{2}', '{3}', 0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE,E'\\000\\000\\000')",
                                guid105, dic["105"], name105, dic0[name105.Trim()]).Append(", ");
                            Console.WriteLine(@"插入文件完毕", dic["105"]);
                            dic2.Add("105", guid105);
                            dic02.Add("105", dic0[name105.Trim()]);

                            var name106 = item.Name106;
                            var guid106 = Guid.NewGuid().ToString();
                            sb2.AppendFormat(
                                @"('{0}', '{1}', '{2}', '{3}', 0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE,E'\\000\\000\\000')",
                                guid106, dic["105"], name106, dic0[name106.Trim()]).Append(", ");
                            Console.WriteLine(@"插入文件完毕", dic["105"]);
                            dic2.Add("106", guid106);
                            dic02.Add("106", dic0[name106.Trim()]);

                            var name107 = item.Name107;
                            var guid107 = Guid.NewGuid().ToString();
                            sb2.AppendFormat(
                                @"('{0}', '{1}', '{2}', '{3}', 0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE,E'\\000\\000\\000')",
                                guid107, dic["105"], name107, dic0[name107.Trim()]).Append(", ");
                            Console.WriteLine(@"插入文件完毕", dic["105"]);
                            dic2.Add("107", guid107);
                            dic02.Add("107", dic0[name107.Trim()]);

                            //var name108 = item.Name108;
                            //sb.AppendFormat(@"插入文件完毕", dic["108"]);
                            //dic2.Add("108", guid);


                            //var name109 = item.Name109;
                            //sb.AppendFormat(@"插入文件完毕", dic["109"]);
                            //dic2.Add("109", guid);


                            //var name110 = item.Name110;
                            //sb.AppendFormat(@"插入文件完毕", dic["110"]);
                            //dic2.Add("110", guid);

                            //var name111 = item.Name111;
                            //sb.AppendFormat(@"插入盒完毕", dic["110"]);
                            //dic2.Add("111", guid);

                            //var name112 = item.Name112;
                            //sb.AppendFormat(@"插入文件完毕", dic["112"]);
                            //dic2.Add("112", guid);

                            //var name113 = item.Name113;
                            //sb.AppendFormat(@"插入盒完毕", dic["112"]);
                            //dic2.Add("113", guid);


                            //var name114 = item.Name114;
                            //sb.AppendFormat(@"插入盒完毕", dic["112"]);
                            //dic2.Add("114", guid);


                            //var name115 = item.Name115;
                            //sb.AppendFormat(@"插入文件完毕", dic["115"]);
                            //dic2.Add("115", guid);

                            //var name116 = item.Name116;
                            //sb.AppendFormat(@"插入盒完毕", dic["115"]);
                            //dic2.Add("116", guid);


                            //var name117 = item.Name117;
                            //sb.AppendFormat(@"插入盒完毕", dic["115"]);
                            //dic2.Add("117", guid);

                            //var name118 = item.Name118;
                            //sb.AppendFormat(@"插入盒完毕", dic["118"]);
                            //dic2.Add("118", guid);


                            //var name119 = item.Name119;
                            //sb.AppendFormat(@"插入盒完毕", dic["118"]);
                            //dic2.Add("119", guid);

                            //var name120 = item.Name120;
                            //sb.AppendFormat(@"插入盒完毕", dic["120"]);
                            //dic2.Add("120", guid);


                            //var name121 = item.Name121;
                            //sb.AppendFormat(@"插入盒完毕", dic["120"]);
                            //dic2.Add("121", guid);


                            //var name122 = item.Name122;
                            //sb.AppendFormat(@"插入盒完毕", dic["122"]);
                            //dic2.Add("122", guid);

                            //var name123 = item.Name123;
                            //sb.AppendFormat(@"插入盒完毕", dic["123"]);
                            //dic2.Add("123", guid);


                            //var name124 = item.Name124;
                            //sb.AppendFormat(@"插入盒完毕", dic["123"]);
                            //dic2.Add("124", guid);

                            //var name125 = item.Name125;
                            //sb.AppendFormat(@"插入盒完毕", dic["123"]);
                            //dic2.Add("125", guid);

                            //var name126 = item.Name126;
                            //sb.AppendFormat(@"插入盒完毕", dic["126"]);
                            //dic2.Add("126", guid);

                            //var name127 = item.Name127;
                            //sb.AppendFormat(@"插入盒完毕", dic["127"]);
                            //dic2.Add("127", guid);

                        }
                    }
                    else
                    {
                        if (index == 2)
                        {

                        

                        sb3.Append(@"INSERT INTO ").Append("\"").Append("DocumentTypeDefinition").Append("\" ").Append("(")
                            .Append("\"").Append("DocumentTypeDefinitionId").Append("\"").Append(", ")
                            .Append("\"").Append("DocumentId").Append("\"").Append(", ")
                            .Append("\"").Append("DocumentTypeValue").Append("\"").Append(", ")
                            .Append("\"").Append("Name").Append("\"").Append(", ")
                            .Append("\"").Append("Code").Append("\"").Append(", ")
                            .Append("\"").Append("ColumnType").Append("\"").Append(", ")
                            .Append("\"").Append("Length").Append("\"").Append(", ")
                            .Append("\"").Append("UsingVisible").Append("\"").Append(", ")
                            .Append("\"").Append("FullTextSearch").Append("\"").Append(", ")
                            .Append("\"").Append("Required").Append("\"").Append(", ")
                            .Append("\"").Append("AllowEditing").Append("\"").Append(", ")
                            .Append("\"").Append("AllowMultiEditing").Append("\"").Append(", ")
                            .Append("\"").Append("IsSerialNumber").Append("\"").Append(", ")
                            .Append("\"").Append("LinkCode").Append("\"").Append(", ")
                            .Append("\"").Append("Regular").Append("\"").Append(", ")
                            .Append("\"").Append("Enable").Append("\"").Append(", ")
                            .Append("\"").Append("IsFixed").Append("\"").Append(", ")
                            .Append("\"").Append("Sort").Append("\"").Append(", ")
                            .Append("\"").Append("CreationTime").Append("\"").Append(", ")
                            .Append("\"").Append("Creator").Append("\"").Append(", ")
                            .Append("\"").Append("CreatorId").Append("\"").Append(", ")
                            .Append("\"").Append("LastModificationTime").Append("\"").Append(", ")
                            .Append("\"").Append("LastModifierId").Append("\"").Append(", ")
                            .Append("\"").Append("Modifier").Append("\"").Append(", ")
                            .Append("\"").Append("IsInherit").Append("\"").Append(", ")
                            .Append("\"").Append("IsDeleted").Append("\"").Append(", ")
                            .Append("\"").Append("Version").Append("\"")
                            .Append(") Values ");
                        }
                       
                        var name100 = item.Name100;
                        if (!string.IsNullOrEmpty(name100))
                        {
                            sb3.AppendFormat(
                            @"	('{0}', '{1}', '{2}','{3}', '{4}',{5},{6},{7},{8},{9},{10},{11},{12},'{15}','',{13},{14},{16}, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE,FALSE,E'\\000\\000\\000')",
                            Guid.NewGuid().ToString(), dic["100"], dic02["100"], name100, dicRule[name100].Name101,dicType[dicRule[name100].Name102], dicRule[name100].Name103, dicTF[dicRule[name100].Name108], dicTF[dicRule[name100].Name104], dicTF[dicRule[name100].Name105], dicTF[dicRule[name100].Name106], dicTF[dicRule[name100].Name107],false, dicTF[dicRule[name100].Name109],true, dicRule[name100].Name111, dicRule[name100].Name112).Append(", ");

                        }




                        var name101 = item.Name101;
                        if (!string.IsNullOrEmpty(name101))
                        {
                            sb3.AppendFormat(
                                @"	('{0}', '{1}', '{2}','{3}', '{4}',{5},{6},{7},{8},{9},{10},{11},{12},'{15}','',{13},{14},{16}, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys',FALSE, FALSE,E'\\000\\000\\000')",
                                Guid.NewGuid().ToString(), dic["100"], dic02["101"], name101, dicRule[name101].Name101,
                                dicType[dicRule[name101].Name102], dicRule[name101].Name103,
                                dicTF[dicRule[name101].Name108], dicTF[dicRule[name101].Name104],
                                dicTF[dicRule[name101].Name105], dicTF[dicRule[name101].Name106],
                                dicTF[dicRule[name101].Name107], false, dicTF[dicRule[name101].Name109], true,
                                dicRule[name101].Name111, dicRule[name101].Name112).Append(", ");


                            //sb3.AppendFormat(
                            //@"	('{0}', '{1}', '{2}','{3}', '{4}',  0,0,true,true,true,true,true,true,'','',true,true,0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE)",
                            //Guid.NewGuid().ToString(), dic["100"], dic02["101"], name101, name101).Append(", ");
                        }

                        var name102 = item.Name102;  
                        if (!string.IsNullOrEmpty(name102))
                        {

                            sb3.AppendFormat(
                                @"	('{0}', '{1}', '{2}','{3}', '{4}',{5},{6},{7},{8},{9},{10},{11},{12},'{15}','',{13},{14},{16}, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys',FALSE, FALSE,E'\\000\\000\\000')",
                                Guid.NewGuid().ToString(), dic["102"], dic02["102"], name102, dicRule[name102].Name101,
                                dicType[dicRule[name102].Name102], dicRule[name102].Name103,
                                dicTF[dicRule[name102].Name108], dicTF[dicRule[name102].Name104],
                                dicTF[dicRule[name102].Name105], dicTF[dicRule[name102].Name106],
                                dicTF[dicRule[name102].Name107], false, dicTF[dicRule[name102].Name109], true,
                                dicRule[name102].Name111, dicRule[name102].Name112).Append(", ");

                            //sb3.AppendFormat(
                            //    @"	('{0}', '{1}', '{2}','{3}', '{4}',  0,0,true,true,true,true,true,true,'','',true,true,0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE)",
                            //    Guid.NewGuid().ToString(), dic["102"], dic02["102"], name102, name102).Append(", ");
                        }

                        var name103 = item.Name103;
                        if (!string.IsNullOrEmpty(name103))
                        {
                            sb3.AppendFormat(
                                @"	('{0}', '{1}', '{2}','{3}', '{4}',{5},{6},{7},{8},{9},{10},{11},{12},'{15}','',{13},{14},{16}, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys',FALSE, FALSE,E'\\000\\000\\000')",
                                Guid.NewGuid().ToString(), dic["102"], dic02["103"], name103, dicRule[name103].Name101,
                                dicType[dicRule[name103].Name102], dicRule[name103].Name103,
                                dicTF[dicRule[name103].Name108], dicTF[dicRule[name103].Name104],
                                dicTF[dicRule[name103].Name105], dicTF[dicRule[name103].Name106],
                                dicTF[dicRule[name103].Name107], false, dicTF[dicRule[name103].Name109], true,
                                dicRule[name103].Name111, dicRule[name103].Name112).Append(", ");


                            //sb3.AppendFormat(
                            //@"	('{0}', '{1}', '{2}','{3}', '{4}',  0,0,true,true,true,true,true,true,'','',true,true,0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE)",
                            //Guid.NewGuid().ToString(), dic["102"], dic02["103"], name103, name103).Append(", ");
                        }

                        var name104 = item.Name104;  
                        if (!string.IsNullOrEmpty(name104))
                        {
                            sb3.AppendFormat(
                                @"	('{0}', '{1}', '{2}','{3}', '{4}',{5},{6},{7},{8},{9},{10},{11},{12},'{15}','',{13},{14},{16}, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys',FALSE, FALSE,E'\\000\\000\\000')",
                                Guid.NewGuid().ToString(), dic["104"], dic02["104"], name104, dicRule[name104].Name101,
                                dicType[dicRule[name104].Name102], dicRule[name104].Name103,
                                dicTF[dicRule[name104].Name108], dicTF[dicRule[name104].Name104],
                                dicTF[dicRule[name104].Name105], dicTF[dicRule[name104].Name106],
                                dicTF[dicRule[name104].Name107], false, dicTF[dicRule[name104].Name109], true,
                                dicRule[name104].Name111, dicRule[name104].Name112).Append(", ");

                            //sb3.AppendFormat(
                            //    @"	('{0}', '{1}', '{2}','{3}', '{4}',  0,0,true,true,true,true,true,true,'','',true,true,0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE)",
                            //    Guid.NewGuid().ToString(), dic["104"], dic02["104"], name104, name104).Append(", ");
                        }

                        var name105 = item.Name105;  
                        if (!string.IsNullOrEmpty(name105))
                        {
                            sb3.AppendFormat(
                                @"	('{0}', '{1}', '{2}','{3}', '{4}',{5},{6},{7},{8},{9},{10},{11},{12},'{15}','',{13},{14},{16}, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys',FALSE, FALSE,E'\\000\\000\\000')",
                                Guid.NewGuid().ToString(), dic["105"], dic02["105"], name105, dicRule[name105].Name101,
                                dicType[dicRule[name105].Name102], dicRule[name105].Name103,
                                dicTF[dicRule[name105].Name108], dicTF[dicRule[name105].Name104],
                                dicTF[dicRule[name105].Name105], dicTF[dicRule[name105].Name106],
                                dicTF[dicRule[name105].Name107], false, dicTF[dicRule[name105].Name109], true,
                                dicRule[name105].Name111, dicRule[name105].Name112).Append(", ");

                            //sb3.AppendFormat(
                            //    @"	('{0}', '{1}', '{2}','{3}', '{4}',  0,0,true,true,true,true,true,true,'','',true,true,0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE)",
                            //    Guid.NewGuid().ToString(), dic["105"], dic02["105"], name105, name105).Append(", ");
                        }

                        var name106 = item.Name106;  
                        if (!string.IsNullOrEmpty(name106))
                        {

                            sb3.AppendFormat(
                                @"	('{0}', '{1}', '{2}','{3}', '{4}',{5},{6},{7},{8},{9},{10},{11},{12},'{15}','',{13},{14},{16}, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys',FALSE, FALSE,E'\\000\\000\\000')",
                                Guid.NewGuid().ToString(), dic["105"], dic02["106"], name106, dicRule[name106].Name101,
                                dicType[dicRule[name106].Name102], dicRule[name106].Name103,
                                dicTF[dicRule[name106].Name108], dicTF[dicRule[name106].Name104],
                                dicTF[dicRule[name106].Name105], dicTF[dicRule[name106].Name106],
                                dicTF[dicRule[name106].Name107], false, dicTF[dicRule[name106].Name109], true,
                                dicRule[name106].Name111, dicRule[name106].Name112).Append(", ");

                            //sb3.AppendFormat(
                            //@"	('{0}', '{1}', '{2}','{3}', '{4}',  0,0,true,true,true,true,true,true,'','',true,true,0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE)",
                            //Guid.NewGuid().ToString(), dic["105"], dic02["106"], name106, name106).Append(", ");
                        }
                        var name107 = item.Name107;  
                        if (!string.IsNullOrEmpty(name107))
                        {
                            sb3.AppendFormat(
                                @"	('{0}', '{1}', '{2}','{3}', '{4}',{5},{6},{7},{8},{9},{10},{11},{12},'{15}','',{13},{14},{16}, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys',FALSE, FALSE,E'\\000\\000\\000')",
                                Guid.NewGuid().ToString(), dic["105"], dic02["107"], name107, dicRule[name107].Name101,
                                dicType[dicRule[name107].Name102], dicRule[name107].Name103,
                                dicTF[dicRule[name107].Name108], dicTF[dicRule[name107].Name104],
                                dicTF[dicRule[name107].Name105], dicTF[dicRule[name107].Name106],
                                dicTF[dicRule[name107].Name107], false, dicTF[dicRule[name107].Name109], true,
                                dicRule[name107].Name111, dicRule[name107].Name112).Append(", ");



                            //sb3.AppendFormat(
                            //@"	('{0}', '{1}', '{2}','{3}', '{4}',  0,0,true,true,true,true,true,true,'','',true,true,0, '2020/9/22 11:00:14', 'sys', '00000000-0000-0000-0000-000000000000', '2020/9/22 11:00:14', '00000000-0000-0000-0000-000000000000', 'sys', FALSE)",
                            //Guid.NewGuid().ToString(), dic["105"], dic02["107"], name107, name107).Append(", ");
                        }

                        //var name108 = item.Name108; Console.WriteLine(name108);
                        //var name109 = item.Name109; Console.WriteLine(name109);
                        //var name110 = item.Name110; Console.WriteLine(name110);
                        //var name111 = item.Name111; Console.WriteLine(name111);
                        //var name112 = item.Name112; Console.WriteLine(name112);
                        //var name113 = item.Name113; Console.WriteLine(name113);
                        //var name114 = item.Name114; Console.WriteLine(name114);
                        //var name115 = item.Name115; Console.WriteLine(name115);
                        //var name116 = item.Name116; Console.WriteLine(name116);
                        //var name117 = item.Name117; Console.WriteLine(name117);
                        //var name118 = item.Name118; Console.WriteLine(name118);
                        //var name119 = item.Name119; Console.WriteLine(name119);
                        //var name120 = item.Name120; Console.WriteLine(name120);
                        //var name121 = item.Name121; Console.WriteLine(name121);
                        //var name122 = item.Name122; Console.WriteLine(name122);
                        //var name123 = item.Name123; Console.WriteLine(name123);
                        //var name124 = item.Name124; Console.WriteLine(name124);
                        //var name125 = item.Name125; Console.WriteLine(name125);
                        //var name126 = item.Name126; Console.WriteLine(name126);
                        //var name127 = item.Name127; Console.WriteLine(name127);
                        //var name128 = item.Name128; Console.WriteLine(name128);

                    }
                    index++;
                }

             





            }



            //var filePath4 = Path.Combine(Directory.GetCurrentDirectory(), "档案模板数据字典.xlsx");
            //using (var stream = new FileStream(filePath4, FileMode.Open))
            //{
            //    var result = await oExcel.ImportAsync<ArchiveTemplateDto>(stream, 1);
            //}


            // //====================================导出
            // var filePath2 = Path.Combine(Directory.GetCurrentDirectory(), "产品单sheel导出模板.xlsx");
            // var list = new List<ExporterProductDto>()
            //     {
            //         new ExporterProductDto { Name="2", Code="0000000002", Date=null},
            //       //  new ExporterProductDto { Name="3", Code="0000000003", Date=DateTime.Now}
            //     };
            // // 设置需要的字段
            //   var xx = "[{\"code\":\"323234\",\"name\":\"地方\",\"startSection\":\"\",\"endSection\":\"\",\"distence\":0.0,\"area\":0.0,\"functionalSequences\":\"\",\"groundWaterHorzon\":\"\",\"mineralizationDegree\":0.0,\"annualTotalSupply\":0.0,\"annualPredictProduction\":0.0,\"annualActualProduction\":0.0,\"beginExam\":null,\"remark\":\"\",\"divisionDir\":\"泸县\",\"unitDir\":\"泸县县水务局\",\"waterQulityGoalsDir\":\"\",\"nutritionStatusTargetDir\":\"\",\"leadingFunctionDir\":\"\",\"levelDir\":\"\",\"geomorphicCategoryDir\":\"\",\"typeDir\":\"地表水功能区\",\"monitorLevelDir\":\"\",\"isExamDir\":\"\",\"division\":\"泸县\",\"unit\":\"泸县县水务局\",\"waterQulityGoals\":\"\",\"nutritionStatusTarget\":\"\",\"leadingFunction\":\"\",\"level\":\"\",\"geomorphicCategory\":\"\",\"type\":\"地表水功能区\",\"monitorLevel\":\"\",\"isExam\":\"\"}]";


            // //var url = "http://localhost:6010/WaterFunctionArea/Export/Excel ";
            // //var data = RestSharpHelper.HttpGet(url);
            // var obj = JsonConvert.DeserializeObject<List<WaterFunctionAreaExportDto>>(xx);
            // var url2 = "http://localhost:6010/WaterFunctionArea/Export/test";
            //var xxx= RestSharpHelper.HttpPost(url2,"{ \"projectIds\":\"229a7163-8449-4cd8-979c-ff1c39836970\",    \"customHeaders\":[        {            \"code\":\"Name\",            \"displayName\":\"名称\"        },        {            \"code\":\"Code\",            \"displayName\":\"编码\"}                  ]}");

            // var heads = new List<ExcelHeader>();
            // heads.Add(new ExcelHeader { Code = "Name", DisplayName = "自定义名称" });
            // heads.Add(new ExcelHeader { Code = "Code", DisplayName = "自定义编码" });
            // heads.Add(new ExcelHeader { Code = "BeginExam", DisplayName = "自定义编码3" });
            // //var reult = await oExcel.ExportAsync(list, heads);
            // var reult2 = await oExcel.ExportAsync(obj, heads);

            // FileStream fs = new FileStream(filePath2, FileMode.Create);
            // fs.Write(reult2, 0, reult2.Length);
            // fs.Dispose();

            string filePath = @"D:\sql1.txt";
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(System.IO.File.Create(filePath)))
            {
                sb.Length--;
                file.WriteLine(sb.ToString());
            }
            string filePath2 = @"D:\sql2.txt";
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(System.IO.File.Create(filePath2)))
            {
                sb2.Length--;
                file.WriteLine(sb2.ToString());
            }
            string filePath3 = @"D:\sql3.txt";
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(System.IO.File.Create(filePath3)))
            {
                sb3.Length--;
                file.WriteLine(sb3.ToString());
            }

            Console.ReadLine();
        }
    }

    public class RestSharpHelper
    {
        public static string HttpGet(string url, int timeOut = 30 * 1000)
        {
            IRestClient client = new RestClient(url);
            IRestRequest request = new RestRequest(Method.GET);
            request.Timeout = timeOut;
            IRestResponse restResponse = client.Execute(request);
            return restResponse.Content;
        }

        public static string HttpPost(string url,  string content, int timeOut = 30 * 1000)
        {

       
        var client = new RestClient(url);
        var request = new RestRequest(Method.POST);
        request.Timeout = 10000;
        //request.AddHeader("Cache-Control", "no-cache");          
        request.AddParameter("application/json", content, ParameterType.RequestBody);
 
        IRestResponse response = client.Execute(request);
            return response.Content; //返回的结果
        }

    }
}
