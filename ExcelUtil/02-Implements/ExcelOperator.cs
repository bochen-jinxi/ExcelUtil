using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;
using ExcelUtil._01_Abstractions;
using ExcelUtil._05_Utils;
using ExcelUtil._06_Model;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;

namespace ExcelUtil
{
    public class ExcelOperator : IExcelOperator
    {
        /// <summary>
        ///     Excel 导入
        /// </summary>
        public IExcelImporter ExcelImporter { get; }

        /// <summary>
        ///     Excel 导出
        /// </summary>
        public IExcelExporter ExcelExporter { get; }

        public ExcelOperator(IExcelImporter excelImporter, IExcelExporter excelExporter)
        {
            ExcelImporter = excelImporter;
            ExcelExporter = excelExporter;
        }

        /// <summary>
        ///     生成模板
        /// </summary>
        /// <returns></returns>
        public async Task<ExportFileInfo> GenerateTemplateFileAsync<T>(string filePath) where T : class, new()
        {
            FileUtil.DeleteFile(filePath);
            var result = await ExcelImporter.GenerateTemplate<T>(filePath);
            return result;
        }

        /// <summary>
        ///     生成模板
        /// </summary>
        /// <returns></returns>
        public async Task<byte[]> GenerateTemplateBytesAsync<T>() where T : class, new()
        {
            var result = await ExcelImporter.GenerateTemplateBytes<T>();
            return result;
        }

        /// <summary>
        ///     生成模板
        /// </summary>
        /// <returns></returns>
        public async Task<Stream> GenerateTemplateStreamAsync<T>() where T : class, new()
        {
            var bytes = await GenerateTemplateBytesAsync<T>();
            using (var ms = new MemoryStream(100))
            {
                var count = 0;
                while (count < bytes.Length) ms.WriteByte(bytes[count++]);
                return ms;
            }
        }

        /// <summary>
        ///     导出数据
        /// </summary>
        /// <typeparam name="T">数据类型</typeparam>
        /// <param name="source">数据源</param>
        /// <param name="templatePath">模板路径</param>
        /// <returns></returns>
        public async Task<byte[]> ExportAsync<T>(List<T> source, string templatePath) where T : class
        {
            var result = await ExcelExporter.ExportBytesByTemplate(source, templatePath);
            return result;
        }

        /// <summary>
        ///     导出数据
        /// </summary>
        /// <param name="source">数据源</param>
        /// <param name="filePath">保存路径</param>
        /// <param name="templatePath">模板路径</param>
        /// <returns></returns>
        public async Task<ExportFileInfo> ExportFileAsync<T>(T source, string templatePath, string filePath)
            where T : class
        {
            FileUtil.DeleteFile(filePath);
            var result = await ExcelExporter.ExportByTemplate(filePath, source, templatePath);
            return result;
        }

        /// <summary>
        ///     导出数据
        /// </summary>
        /// <typeparam name="T">数据类型</typeparam>
        /// <param name="source">数据源</param>
        /// <param name="excelHeaders">自定义 Excel 表头</param>
        /// <returns></returns>
        public async Task<byte[]> ExportAsync<T>(List<T> source, List<ExcelHeader> excelHeaders = null) where T : class
        {
            return excelHeaders == null || !excelHeaders.Any()
                ? await ExcelExporter.ExportAsByteArray(source)
                : await ExcelExporter.ExportAsByteArray(ExcelUtil.ToDataTable(source),
                    new ExcelHeaderFilter(excelHeaders));
        }

        /// <summary>
        ///     导出数据
        /// </summary>
        /// <param name="source">数据源</param>
        /// <param name="filePath">保存路径</param>
        /// <param name="excelHeaders">自定义 Excel 表头</param>
        /// <returns></returns>
        public async Task<ExportFileInfo> ExportFileAsync<T>(List<T> source, string filePath,
            List<ExcelHeader> excelHeaders = null) where T : class
        {
            FileUtil.DeleteFile(filePath);
            return excelHeaders == null || !excelHeaders.Any()
                ? await ExcelExporter.Export(filePath, source)
                : await ExcelExporter.Export(filePath, ExcelUtil.ToDataTable(source),
                    new ExcelHeaderFilter(excelHeaders));
        }

        /// <summary>
        ///     根据模板Dto生成导入Excel模板文件
        /// </summary>
        /// <param name="filePath">保存路径</param>
        /// <returns></returns>
        public async Task GenerateImportTemplateAsync<T>(string filePath = "") where T : class, new()
        {
            FileUtil.DeleteFile(filePath);
            await ExcelImporter.GenerateTemplate<T>(filePath);
        }

        /// <summary>
        ///     表级别读取
        ///     返回结果默认第一个Sheet
        /// </summary>
        /// <param name="filePath">保存路径</param>
        /// <returns></returns>
        public async Task<ImportResult<T>> ImportAsync<T>(string filePath = "") where T : class, new()
        {
            if (!File.Exists(filePath))
            {
                return null;
            }

            var import = await ExcelImporter.Import<T>(filePath);
            return import;
        }

        /// <summary>
        ///  导入 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">文件流</param>
        /// <returns></returns>
        public async Task<ImportResult<T>> ImportAsync<T>(Stream stream) where T : class, new()
        {
            var import = await ExcelImporter.Import<T>(stream);
            return import;
        }
        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">文件流</param>
        /// <param name="rowIndex">起始行</param>
        /// <returns></returns>
        public Task<ImportResult<T>> ImportAsync<T>(Stream stream, int rowIndex) where T : class, new()
        {
            using (var importer = new ImportRowRecords<T>(stream, rowIndex))
            {
                return importer.Import();
            }
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">文件流</param>
        /// <param name="rowIndex">起始行</param>
        /// <param name="mapping">Excel映射列  </param>
        /// <returns></returns>
        public async Task<ImportResult<T>> ImportAsync<T>(Stream stream, int rowIndex, params string [] mapping ) where T : class, new()
        {
            var importer = new ImportRowRecords<T>(stream, rowIndex, mapping);
            using (importer)
            {
                ImportResult<T> obj = null;
                try
                {
                    obj = await importer.Import();
                }
                finally
                {
                    if (importer.Data != null&&obj != null) obj.Data = importer.Data;
                }
                return obj;
            }
        }
     
        /// <summary>
        /// 读取水文关系曲线
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">文件流</param>
        /// <param name="rowIndex">行号</param>
        /// <returns></returns>
        public async Task<Dictionary<string, string>> ReadHydrologicalRelationCurveAsync<T>(Stream stream, int rowIndex)
            where T : ImportWatermark, new()
        {
            var result = await ImportAsync<T>(stream, rowIndex);
            var constVal = 65;
            var dic = new Dictionary<string, string>();
            foreach (var item in result.Data.Where(x => !(x.Watermark.Contains("量") || x.Watermark.Contains("水位(米)") || x.Watermark.Contains("(米3/秒)"))))
            {
                 

                var watermarkA = item.Watermark + (nameof(item.A).ASCII() - constVal);
                var watermarkB = item.Watermark + (nameof(item.B).ASCII() - constVal);
                var watermarkC = item.Watermark + (nameof(item.C).ASCII() - constVal);
                var watermarkD = item.Watermark + (nameof(item.D).ASCII() - constVal);
                var watermarkE = item.Watermark + (nameof(item.E).ASCII() - constVal);
                var watermarkF = item.Watermark + (nameof(item.F).ASCII() - constVal);
                var watermarkG = item.Watermark + (nameof(item.G).ASCII() - constVal);
                var watermarkH = item.Watermark + (nameof(item.H).ASCII() - constVal);
                var watermarkI = item.Watermark + (nameof(item.I).ASCII() - constVal);

                dic.Add(watermarkA, item.A);
                dic.Add(watermarkB, item.B);
                dic.Add(watermarkC, item.C);
                dic.Add(watermarkD, item.D);
                dic.Add(watermarkE, item.E);
                dic.Add(watermarkF, item.F);
                dic.Add(watermarkG, item.G);
                dic.Add(watermarkH, item.H);
                dic.Add(watermarkI, item.I);
            }
            var data = dic.Where(x => x.Value != null).ToArray().ToDictionary(i => i.Key, j => j.Value);
            return data;
        }
    }
}