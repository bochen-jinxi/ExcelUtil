using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;

namespace ExcelUtil
{
    public partial class Excel : Base, IExcel
    {
        public IExcelImporter oImporter { get; }
        public IExcelExporter oExporter { get; }

        public Excel(IExcelImporter importer, IExcelExporter exporter)
        {
            oImporter = importer;
            oExporter = exporter;
        }

        public async Task<ExportFileInfo> ExportSingByTemplateAsync<T>(string filePath, T data, string template)
            where T : class
        {
            //创建Excel导出对象
            if (File.Exists(filePath)) File.Delete(filePath);
            //根据模板导出
            var result = await oExporter.ExportByTemplate(filePath, data, template);
            return result;
        }

        public async Task<byte[]> ExportSingByTemplateAsByteAsync<T>(T data, string template) where T : class
        {
            //根据模板导出
            var result = await oExporter.ExportBytesByTemplate(data, template);
            return result;
        }

        public async Task<ExportFileInfo> GenerateTemplateAsync<T>(string filePath) where T : class, new()
        {
            DeleteFile(filePath);
            var result = await oImporter.GenerateTemplate<T>(filePath);
            return result;
        }

        public async Task<byte[]> GenerateTemplateBytesAsync<T>() where T : class, new()
        {
            var result = await oImporter.GenerateTemplateBytes<T>();
            return result;
        }

        public async Task<Stream> GenerateTemplateStreamAsync<T>() where T : class, new()
        {
            var bytes = await GenerateTemplateBytesAsync<T>();
            // 把 byte[] 转换成 Stream
            using (var ms = new MemoryStream(100))
            {
                var count = 0;
                while (count < bytes.Length) ms.WriteByte(bytes[count++]);
                return ms;
            }
        }

        public async Task<byte[]> ExportSingSheelAsByteAsync<T>(ICollection<T> datas) where T : class
        {
            var result = await oExporter.ExportAsByteArray(datas);
            return result;
        }

        public async Task<ExportFileInfo> ExportSingSheelAsync<T>(string filePath, ICollection<T> datas) where T : class
        {
            DeleteFile(filePath);
            var result = await oExporter.Export(filePath, datas);
            return result;
        }

        public async Task<byte[]> ExportMultSheelAsByteAsync<T1, T2>(ICollection<T1> li1, ICollection<T2> li2)
            where T1 : class where T2 : class
        {
            oExporter.Append(li1);
            oExporter.Append(li2);
            var result = await oExporter.ExportAppendDataAsByteArray();
            return result;
        }

        public async Task<byte[]> ExportMultSheelAsByteAsync<T1, T2, T3>(ICollection<T1> li1, ICollection<T2> li2,
            ICollection<T3> li3) where T1 : class where T2 : class where T3 : class
        {
            oExporter.Append(li1);
            oExporter.Append(li2);
            oExporter.Append(li3);
            var result = await oExporter.ExportAppendDataAsByteArray();
            return result;
        }

        public async Task<byte[]> ExportMultSheelAsByteAsync<T1, T2, T3, T4>(ICollection<T1> li1, ICollection<T2> li2,
            ICollection<T3> li3, ICollection<T4> li4)
            where T1 : class where T2 : class where T3 : class where T4 : class
        {
            oExporter.Append(li1);
            oExporter.Append(li2);
            oExporter.Append(li3);
            oExporter.Append(li4);
            var result = await oExporter.ExportAppendDataAsByteArray();
            return result;
        }

        public async Task<byte[]> ExportMultSheelAsByteAsync<T1, T2, T3, T4, T5>(ICollection<T1> li1,
            ICollection<T2> li2, ICollection<T3> li3, ICollection<T4> li4, ICollection<T5> li5) where T1 : class
            where T2 : class
            where T3 : class
            where T4 : class
            where T5 : class
        {
            oExporter.Append(li1);
            oExporter.Append(li2);
            oExporter.Append(li3);
            oExporter.Append(li4);
            oExporter.Append(li5);
            var result = await oExporter.ExportAppendDataAsByteArray();
            return result;
        }

        public async Task<ExportFileInfo> ExportMultSheelAsync<T1, T2>(string filePath, ICollection<T1> li1,
            ICollection<T2> li2) where T1 : class where T2 : class
        {
            DeleteFile(filePath);
            oExporter.Append(li1);
            oExporter.Append(li2);
            var result = await oExporter.ExportAppendData(filePath);
            return result;
        }

        public async Task<ExportFileInfo> ExportMultSheelAsync<T1, T2, T3>(string filePath, ICollection<T1> li1,
            ICollection<T2> li2, ICollection<T3> li3) where T1 : class where T2 : class where T3 : class
        {
            DeleteFile(filePath);
            oExporter.Append(li1);
            oExporter.Append(li2);
            oExporter.Append(li3);
            var result = await oExporter.ExportAppendData(filePath);
            return result;
        }

        public async Task<ExportFileInfo> ExportMultSheelAsync<T1, T2, T3, T4>(string filePath, ICollection<T1> li1,
            ICollection<T2> li2, ICollection<T3> li3,
            ICollection<T4> li4) where T1 : class where T2 : class where T3 : class where T4 : class
        {
            DeleteFile(filePath);
            oExporter.Append(li1);
            oExporter.Append(li2);
            oExporter.Append(li3);
            oExporter.Append(li4);
            var result = await oExporter.ExportAppendData(filePath);
            return result;
        }

        public async Task<ExportFileInfo> ExportMultSheelAsync<T1, T2, T3, T4, T5>(string filePath, ICollection<T1> li1,
            ICollection<T2> li2, ICollection<T3> li3, ICollection<T4> li4, ICollection<T5> li5)
            where T1 : class where T2 : class where T3 : class where T4 : class where T5 : class
        {
            DeleteFile(filePath);
            oExporter.Append(li1);
            oExporter.Append(li2);
            oExporter.Append(li3);
            oExporter.Append(li4);
            oExporter.Append(li5);
            var result = await oExporter.ExportAppendData(filePath);
            return result;
        }
    }
}