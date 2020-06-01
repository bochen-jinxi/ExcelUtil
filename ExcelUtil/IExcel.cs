using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core.Models;

namespace ExcelUtil
{
    public interface IExcel
    {
        /// <summary>
        ///     生成模板文件
        /// </summary>
        /// <returns></returns>
        Task<ExportFileInfo> GenerateTemplateAsync<T>(string filePath) where T : class, new();

        /// <summary>
        ///     生成模板字节
        /// </summary>
        /// <returns></returns>
        Task<byte[]> GenerateTemplateBytesAsync<T>() where T : class, new();

        /// <summary>
        ///     生成模板流
        /// </summary>
        /// <returns></returns>
        Task<Stream> GenerateTemplateStreamAsync<T>() where T : class, new();

        /// <summary>
        ///     导出数据单个sheel
        /// </summary>
        /// <param name="filePath">保存路径</param>
        /// <param name="datas">数据源</param>
        /// <returns></returns>
        Task<ExportFileInfo> ExportSingSheelAsync<T>(string filePath, ICollection<T> datas) where T : class;

        /// <summary>
        ///     导出多sheel且数据源各不相同
        /// </summary>
        /// <typeparam name="T1"></typeparam>
        /// <typeparam name="T2"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="li1"></param>
        /// <param name="li2"></param>
        /// <returns></returns>
        Task<ExportFileInfo> ExportMultSheelAsync<T1, T2>(string filePath, ICollection<T1> li1, ICollection<T2> li2)
            where T1 : class where T2 : class;

        /// <summary>
        ///     导出多sheel且格式各不相同
        /// </summary>
        /// <typeparam name="T1"></typeparam>
        /// <typeparam name="T2"></typeparam>
        /// <typeparam name="T3"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="li1"></param>
        /// <param name="li2"></param>
        /// <param name="li3"></param>
        /// <returns></returns>
        Task<ExportFileInfo> ExportMultSheelAsync<T1, T2, T3>(string filePath, ICollection<T1> li1, ICollection<T2> li2, ICollection<T3> li3)
            where T1 : class where T2 : class where T3 : class;

        /// <summary>
        ///     导出多sheel且数据源各不相同
        /// </summary>
        /// <typeparam name="T1"></typeparam>
        /// <typeparam name="T2"></typeparam>
        /// <typeparam name="T3"></typeparam>
        /// <typeparam name="T4"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="li1"></param>
        /// <param name="li2"></param>
        /// <param name="li3"></param>
        /// <param name="li4"></param>
        /// <returns></returns>
        Task<ExportFileInfo> ExportMultSheelAsync<T1, T2, T3, T4>(string filePath, ICollection<T1> li1, ICollection<T2> li2, ICollection<T3> li3, ICollection<T4> li4)
            where T1 : class where T2 : class where T3 : class where T4 : class;

        /// <summary>
        ///     导出多sheel且数据源各不相同
        /// </summary>
        /// <typeparam name="T1"></typeparam>
        /// <typeparam name="T2"></typeparam>
        /// <typeparam name="T3"></typeparam>
        /// <typeparam name="T4"></typeparam>
        /// <typeparam name="T5"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="li1"></param>
        /// <param name="li2"></param>
        /// <param name="li3"></param>
        /// <param name="li4"></param>
        /// <param name="li5"></param>
        /// <returns></returns>
        Task<ExportFileInfo> ExportMultSheelAsync<T1, T2, T3, T4, T5>(string filePath, ICollection<T1> li1, ICollection<T2> li2, ICollection<T3> li3, ICollection<T4> li4, ICollection<T5> li5)
            where T1 : class where T2 : class where T3 : class where T4 : class where T5 : class;

        /// <summary>
        ///     根据模板导出
        /// </summary>
        /// <param name="filePath">保存路径</param>
        /// <param name="data"></param>
        /// <param name="template">模板路径</param>
        /// <returns></returns>
        Task<ExportFileInfo> ExportSingByTemplateAsync<T>(string filePath, T data,
            string template) where T : class;

        /// <summary>
        ///     根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="template"></param>
        /// <returns></returns>
        Task<byte[]> ExportSingByTemplateAsByteAsync<T>(T data, string template) where T : class;

        /// <summary>
        ///     导出二进制多sheel且数据源各不相同
        /// </summary>
        /// <typeparam name="T1"></typeparam>
        /// <typeparam name="T2"></typeparam>
        /// <param name="li1"></param>
        /// <param name="li2"></param>
        /// <returns></returns>
        Task<byte[]> ExportMultSheelAsByteAsync<T1, T2>(ICollection<T1> li1, ICollection<T2> li2)
            where T1 : class where T2 : class;

        /// <summary>
        ///     导出二进制多sheel且数据源各不相同
        /// </summary>
        /// <typeparam name="T1"></typeparam>
        /// <typeparam name="T2"></typeparam>
        /// <typeparam name="T3"></typeparam>
        /// <param name="li1"></param>
        /// <param name="li2"></param>
        /// <param name="li3"></param>
        /// <returns></returns>
        Task<byte[]> ExportMultSheelAsByteAsync<T1, T2, T3>(ICollection<T1> li1, ICollection<T2> li2, ICollection<T3> li3)
            where T1 : class where T2 : class where T3 : class;

        /// <summary>
        ///     导出二进制多sheel且数据源各不相同
        /// </summary>
        /// <typeparam name="T1"></typeparam>
        /// <typeparam name="T2"></typeparam>
        /// <typeparam name="T3"></typeparam>
        /// <typeparam name="T4"></typeparam>
        /// <param name="li1"></param>
        /// <param name="li2"></param>
        /// <param name="li3"></param>
        /// <param name="li4"></param>
        /// <returns></returns>
        Task<byte[]> ExportMultSheelAsByteAsync<T1, T2, T3, T4>(ICollection<T1> li1, ICollection<T2> li2, ICollection<T3> li3, ICollection<T4> li4)
            where T1 : class where T2 : class where T3 : class where T4 : class;

        /// <summary>
        ///     导出二进制多sheel且数据源各不相同
        /// </summary>
        /// <typeparam name="T1"></typeparam>
        /// <typeparam name="T2"></typeparam>
        /// <typeparam name="T3"></typeparam>
        /// <typeparam name="T4"></typeparam>
        /// <typeparam name="T5"></typeparam>
        /// <param name="li1"></param>
        /// <param name="li2"></param>
        /// <param name="li3"></param>
        /// <param name="li4"></param>
        /// <param name="li5"></param>
        /// <returns></returns>
        Task<byte[]> ExportMultSheelAsByteAsync<T1, T2, T3, T4, T5>(ICollection<T1> li1, ICollection<T2> li2, ICollection<T3> li3, ICollection<T4> li4, ICollection<T5> li5)
            where T1 : class where T2 : class where T3 : class where T4 : class where T5 : class;

        /// <summary>
        ///     导出数据单个sheel
        /// </summary>
        /// <param name="datas">数据源</param>
        /// <returns></returns>
        Task<byte[]> ExportSingSheelAsByteAsync<T>(ICollection<T> datas) where T : class;

        #region Import

        /// <summary>
        /// 根据模板Dto生成导入Excel模板文件
        /// </summary>
        /// <returns></returns>
        Task GenerateImportTemplateAsync<T>(string filePath = "") where T : class, new();

        /// <summary>
        /// 表级别读取
        /// 返回结果默认第一个Sheet
        /// </summary>
        /// <returns></returns>
        Task<ImportResult<T>> ImportAsync<T>(string filePath = "") where T : class, new();

        /// <summary>
        /// Sheet级别读取
        /// Sheet类型不一致(EG:学生、学生成绩)
        /// 返回结果按照Sheet分组
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        Task<Dictionary<string, ImportResult<object>>> ImportMultipleSheetAsync<T>(string filePath = "") where T : class, new();

        /// <summary>
        /// Sheet级别读取
        /// 返回结果为以Sheet名称为Key，List值为Value的字典，清理掉了系统返回的错误信息
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<Dictionary<string, List<object>>> ImportMultipleSheetLightAsync<T>(string filePath = "") where T : class, new();

        /// <summary>
        /// Sheet级别读取
        /// Sheet类型一致(EG:一班成绩、二班成绩)
        /// 返回结果按照T模型分组
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="TSheet"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        Task<Dictionary<string, ImportResult<TSheet>>> ImportSameSheetsAsync<T, TSheet>(string filePath = "") where T : class, new()
            where TSheet : class, new();

        #endregion Import

        #region Export

        //

        #endregion Export
    }
}