using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq.Expressions;
using System.Threading.Tasks;
using ExcelUtil._06_Model;
using Magicodes.ExporterAndImporter.Core.Models;

namespace ExcelUtil
{
    public interface IExcelOperator
    {
        #region 生成模板

        /// <summary>
        /// 生成模板
        /// </summary>
        /// <returns></returns>
        Task<ExportFileInfo> GenerateTemplateFileAsync<T>(string filePath) where T : class, new();

        /// <summary>
        /// 生成模板
        /// </summary>
        /// <returns></returns>
        Task<byte[]> GenerateTemplateBytesAsync<T>() where T : class, new();

        /// <summary>
        ///  生成模板
        /// </summary>
        /// <returns></returns>
        Task<Stream> GenerateTemplateStreamAsync<T>() where T : class, new();

        #endregion

        #region 导出数据

        #region 导出数据到模板

        /// <summary>
        ///  导出数据
        /// </summary>
        /// <typeparam name="T">数据类型</typeparam>
        /// <param name="source">数据源</param>
        /// <param name="templatePath">模板路径</param>
        /// <returns></returns>
        Task<byte[]> ExportAsync<T>(List<T> source, string templatePath) where T : class;

        /// <summary>
        ///  导出数据
        /// </summary>
        /// <param name="source">数据源</param>
        /// <param name="filePath">保存路径</param>
        /// <param name="templatePath">模板路径</param>
        /// <returns></returns>
        Task<ExportFileInfo> ExportFileAsync<T>(T source, string templatePath, string filePath) where T : class;

        #endregion

        #region 导出数据自定义表头

        /// <summary>
        ///  导出数据
        /// </summary>
        /// <typeparam name="T">数据类型</typeparam>
        /// <param name="source">数据源</param>
        /// <param name="excelHeaders">自定义 Excel 表头</param>
        /// <returns></returns>
        Task<byte[]> ExportAsync<T>(List<T> source, List<ExcelHeader> excelHeaders = null) where T : class;

        /// <summary>
        ///  导出数据
        /// </summary>
        /// <param name="source">数据源</param>
        /// <param name="filePath">保存路径</param>
        /// <param name="excelHeaders">自定义 Excel 表头</param>
        /// <returns></returns>
        Task<ExportFileInfo> ExportFileAsync<T>(List<T> source, string filePath, List<ExcelHeader> excelHeaders = null) where T : class;

        #endregion

        #endregion

        #region 导入

        /// <summary>
        /// 根据模板Dto生成导入Excel模板文件
        /// </summary>
        /// <returns></returns>
        Task GenerateImportTemplateAsync<T>(string filePath = "") where T : class, new();

        /// <summary>
        /// 导入
        /// 返回结果默认第一个Sheet
        /// </summary>
        /// <returns></returns>
        Task<ImportResult<T>> ImportAsync<T>(string filePath = "") where T : class, new();


        /// <summary>
        /// 导入
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <returns></returns>
        Task<ImportResult<T>> ImportAsync<T>(Stream stream) where T : class, new();

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">文件流</param>
        /// <param name="rowIndex">起始行</param>
        /// <param name="mapping">Excel映射列</param>
        /// <returns></returns>
        Task<ImportResult<T>> ImportAsync<T>(Stream stream, int rowIndex, params string[] mapping) where T : class, new();

   
        /// <summary>
        /// 读取水文关系曲线
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <param name="rowIndex">起始行</param>
        /// <returns></returns>
        Task<Dictionary<string, string>> ReadHydrologicalRelationCurveAsync<T>(Stream stream, int rowIndex) where T : ImportWatermark, new();

        #endregion Import
    }
}