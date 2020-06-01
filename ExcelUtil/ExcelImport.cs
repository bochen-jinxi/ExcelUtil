using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelUtil
{
    public partial class Excel : Base, IExcel
    {
        /// <summary>
        /// 根据模板Dto生成导入Excel模板文件
        /// </summary>
        /// <returns></returns>
        public async Task GenerateImportTemplateAsync<T>(string filePath) where T : class, new()
        {
            if (File.Exists(filePath))
            {
                Base.TryDelete(filePath);
            }
            var result = await oImporter.GenerateTemplate<T>(filePath);
        }

        /// <summary>
        /// 表级别读取
        /// 返回结果默认第一个Sheet
        /// </summary>
        /// <returns></returns>
        public async Task<ImportResult<T>> ImportAsync<T>(string filePath) where T : class, new()
        {
            if (!File.Exists(filePath))
            {
                return null;
            }
            else
            {
                var import = await oImporter.Import<T>(filePath);
                return import;
            }
        }

        /// <summary>
        /// Sheet级别读取
        /// Sheet类型不一致(EG:学生、学生成绩)
        /// 返回结果按照Sheet分组
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public async Task<Dictionary<string, ImportResult<object>>> ImportMultipleSheetAsync<T>(string filePath) where T : class, new()
        {
            if (!File.Exists(filePath))
            {
                return null;
            }
            else
            {
                var import = await oImporter.ImportMultipleSheet<T>(filePath);
                return import;
            }
        }

        /// <summary>
        /// Sheet级别读取
        /// 返回结果为以Sheet名称为Key，List值为Value的字典，清理掉了系统返回的错误信息
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public async Task<Dictionary<string, List<object>>> ImportMultipleSheetLightAsync<T>(string filePath) where T : class, new()
        {
            if (!File.Exists(filePath))
            {
                return null;
            }
            else
            {
                var result = await oImporter.ImportMultipleSheet<T>(filePath);

                Dictionary<string, List<object>> result_Dic = new Dictionary<string, List<object>>();
                foreach (var data in result)
                {
                    if (data.Value.Data != null && data.Value.Data.Any())
                    {
                        result_Dic.Add(data.Key, data.Value.Data.ToList());
                    }
                }

                return result_Dic;
            }
        }

        /// <summary>
        /// Sheet级别读取
        /// Sheet类型一致(EG:一班成绩、二班成绩)
        /// 返回结果按照T模型分组
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="TSheet"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public async Task<Dictionary<string, ImportResult<TSheet>>> ImportSameSheetsAsync<T, TSheet>(string filePath)
            where T : class, new()
            where TSheet : class, new()
        {
            if (!File.Exists(filePath))
            {
                return null;
            }
            else
            {
                var import = await oImporter.ImportSameSheets<T, TSheet>(filePath);
                return import;
            }
        }
    }
}