using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using ExcelUtil._06_Model;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel.Utility;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace ExcelUtil._01_Abstractions
{
    /// <summary>
    /// 导入设置动态行
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ImportRowRecords<T> : ImportHelper<T> where T : class, new()
    {
        /// <summary>
        /// 数据
        /// </summary>
        public List<T> Data { get;set; }
        /// <summary>
        /// 起始行
        /// </summary>
        protected int RowIndex { get; }

         

        public string[] Mapping { get; }


        public ImportRowRecords(Stream stream, int rowIndex) : base(stream) => this.RowIndex = rowIndex;

        /// <inheritdoc />
        public ImportRowRecords(Stream stream, int rowIndex, params string[] mapping) : base(stream)
        {
            this.RowIndex = rowIndex;
            Mapping = mapping;
        }
    
       
        protected override void ParseTemplate(ExcelPackage excelPackage) => ParseHeader();

        protected override bool ParseHeader()
        {
            base.ParseHeader();
            ExcelImporterSettings.HeaderRowIndex = (RowIndex-1>0)?RowIndex - 1:1;
            var importerHeaderInfos = this.ImporterHeaderInfos;
            if (null == Mapping) return true;
            for (int i = 0; i < Mapping.Length; i++)
            {
                var index = i + 1;
                var excelObj = Mapping[i];
                if (string.IsNullOrEmpty(excelObj)) continue;
                foreach (var item in importerHeaderInfos)
                {
                    if (item.PropertyName.Equals(excelObj))
                    {
                        item.Header.ColumnIndex = index;
                    }
                }
            }
            return true;
        }
        
        protected override void ParseData(ExcelPackage excelPackage)
        {
            if (null == Mapping || !Mapping.Any())
            {
                base.ParseData(excelPackage);
                return;
            }
            var worksheet = GetImportSheet(excelPackage);
            //行数
            var rows = worksheet.Dimension.End.Row;
            //头行
            var headRowIndex = ExcelImporterSettings.HeaderRowIndex;
            if (rows > ExcelImporterSettings.MaxCount + headRowIndex)
                throw new ArgumentException($"最大允许导入条数不能超过{ExcelImporterSettings.MaxCount}条！");
            var ImportResultData = new List<T>();
            var propertyInfos = new List<PropertyInfo>(typeof(T).GetProperties());
            for (var rowIndex = headRowIndex + 1;rowIndex <= rows; rowIndex++)
            {
                var emptyRow = worksheet.Cells[rowIndex, 1, rowIndex, worksheet.Dimension.End.Column].All(p => p.Text == string.Empty);
                //跳过空行
                if (emptyRow){continue;}
                var dataItem = new T();
                IEnumerable<PropertyInfo> propertys;
                //是否有设置映射
                if (null == Mapping || !Mapping.Any())
                {
                    propertys = propertyInfos.Where(p =>ImporterHeaderInfos.Any(p1 => p1.PropertyName == p.Name )).ToList();
                }
                else
                {
                    propertys = propertyInfos.Where(x => Mapping.Contains(x.Name)).ToList();
                }
                ReadExcelCell(propertys, worksheet, rowIndex, dataItem);
                ImportResultData.Add(dataItem);
            }
            if (ImportResultData.Any())
            {
                Data = ImportResultData;
            }
        }
        /// <summary>
        /// 读取Excel单元格信息 
        /// </summary>
        /// <param name="propertys">属性</param>
        /// <param name="worksheet">页</param>
        /// <param name="rowIndex">行</param>
        /// <param name="dataItem">数据</param>
        private void ReadExcelCell(IEnumerable<PropertyInfo> propertys, ExcelWorksheet worksheet, int rowIndex, T dataItem)
        {
            foreach (var propertyInfo in propertys)
            {
                //列信息
                var col = ImporterHeaderInfos.First(a => a.PropertyName == propertyInfo.Name);
                //单元格
                var cell = worksheet.Cells[rowIndex, col.Header.ColumnIndex];
                try
                {
                    var cellValue = cell.Value?.ToString();
                    if (cellValue.IsNullOrWhiteSpace()) continue;
                    if (!(col.MappingValues.Count > 0 && col.MappingValues.ContainsKey(cellValue)))
                    {
                        CheckType(propertyInfo, dataItem, rowIndex, col, cellValue, cell);
                    }
                }
                catch (Exception ex)
                {
                    AddRowDataError(rowIndex, col, ex.Message);
                }
            }
        }

        /// <summary>
       /// 类型检查
       /// </summary>
       /// <param name="propertyInfo">属性</param>
       /// <param name="dataItem">数据</param>
       /// <param name="rowIndex">行</param>
       /// <param name="col">列</param>
       /// <param name="cellValue">单元格值</param>
       /// <param name="cell">单元格</param>
        private void CheckType(PropertyInfo propertyInfo, T dataItem, int rowIndex, ImporterHeaderInfo col, string cellValue,
            ExcelRange cell)
        {
            switch (propertyInfo.PropertyType.GetCSharpTypeName())
            {
                case "Boolean":
                    propertyInfo.SetValue(dataItem, false);
                    AddRowDataError(rowIndex, col, $"值 {cellValue} 不存在模板下拉选项中");
                    break;
                case "Nullable<Boolean>":
                    if (string.IsNullOrWhiteSpace(cellValue))
                        propertyInfo.SetValue(dataItem, null);
                    else
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 不合法！");
                    break;
                case "String":
                    //移除所有的空格，包括中间的空格
                    if (col.Header.FixAllSpace)
                        propertyInfo.SetValue(dataItem, cellValue?.Replace(" ", string.Empty));
                    else if (col.Header.AutoTrim)
                        propertyInfo.SetValue(dataItem, cellValue?.Trim());
                    else
                        propertyInfo.SetValue(dataItem, cellValue);

                    break;
                //long
                case "Int64":
                {
                    if (!long.TryParse(cellValue, out var number))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, number);
                }
                    break;
                case "Nullable<Int64>":
                {
                    if (string.IsNullOrWhiteSpace(cellValue))
                    {
                        propertyInfo.SetValue(dataItem, null);
                        break;
                    }

                    if (!long.TryParse(cellValue, out var number))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, number);
                }
                    break;
                case "Int32":
                {
                    if (!int.TryParse(cellValue, out var number))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, number);
                }
                    break;
                case "Nullable<Int32>":
                {
                    if (string.IsNullOrWhiteSpace(cellValue))
                    {
                        propertyInfo.SetValue(dataItem, null);
                        break;
                    }

                    if (!int.TryParse(cellValue, out var number))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, number);
                }
                    break;
                case "Int16":
                {
                    if (!short.TryParse(cellValue, out var number))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, number);
                }
                    break;
                case "Nullable<Int16>":
                {
                    if (string.IsNullOrWhiteSpace(cellValue))
                    {
                        propertyInfo.SetValue(dataItem, null);
                        break;
                    }

                    if (!short.TryParse(cellValue, out var number))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, number);
                }
                    break;
                case "Decimal":
                {
                    if (!decimal.TryParse(cellValue, out var number))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, number);
                }
                    break;
                case "Nullable<Decimal>":
                {
                    if (string.IsNullOrWhiteSpace(cellValue))
                    {
                        propertyInfo.SetValue(dataItem, null);
                        break;
                    }

                    if (!decimal.TryParse(cellValue, out var number))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, number);
                }
                    break;
                case "Double":
                {
                    if (!double.TryParse(cellValue, out var number))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, number);
                }
                    break;
                case "Nullable<Double>":
                {
                    if (string.IsNullOrWhiteSpace(cellValue))
                    {
                        propertyInfo.SetValue(dataItem, null);
                        break;
                    }

                    if (!double.TryParse(cellValue, out var number))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, number);
                }
                    break;
                //case "float":
                case "Single":
                {
                    if (!float.TryParse(cellValue, out var number))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, number);
                }
                    break;
                case "Nullable<Single>":
                {
                    if (string.IsNullOrWhiteSpace(cellValue))
                    {
                        propertyInfo.SetValue(dataItem, null);
                        break;
                    }

                    if (!float.TryParse(cellValue, out var number))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, number);
                }
                    break;
                case "DateTime":
                {
                    if (cell.Value == null || cell.Text.IsNullOrWhiteSpace())
                    {
                        AddRowDataError(rowIndex, col, $"值 {cell.Value} 无效，请填写正确的日期时间格式！");
                        break;
                    }

                    try
                    {
                        var date = cell.GetValue<DateTime>();
                        propertyInfo.SetValue(dataItem, date);
                    }
                    catch (Exception)
                    {
                        AddRowDataError(rowIndex, col, $"值 {cell.Value} 无效，请填写正确的日期时间格式！");
                        break;
                    }
                }
                    break;
                case "DateTimeOffset":
                {
                    if (!DateTimeOffset.TryParse(cell.Text, out var date))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cell.Text} 无效，请填写正确的日期时间格式！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, date);
                }
                    break;
                case "Nullable<DateTime>":
                {
                    if (string.IsNullOrWhiteSpace(cell.Text))
                    {
                        propertyInfo.SetValue(dataItem, null);
                        break;
                    }

                    if (!DateTime.TryParse(cell.Text, out var date))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cell.Text} 无效，请填写正确的日期时间格式！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, date);
                }
                    break;
                case "Nullable<DateTimeOffset>":
                {
                    if (string.IsNullOrWhiteSpace(cell.Text))
                    {
                        propertyInfo.SetValue(dataItem, null);
                        break;
                    }

                    if (!DateTimeOffset.TryParse(cell.Text, out var date))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cell.Text} 无效，请填写正确的日期时间格式！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, date);
                }
                    break;
                case "Guid":
                {
                    if (!Guid.TryParse(cellValue, out var guid))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的Guid格式！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, guid);
                }
                    break;
                case "Nullable<Guid>":
                {
                    if (string.IsNullOrWhiteSpace(cellValue))
                    {
                        propertyInfo.SetValue(dataItem, null);
                        break;
                    }

                    if (!Guid.TryParse(cellValue, out var guid))
                    {
                        AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的Guid格式！");
                        break;
                    }

                    propertyInfo.SetValue(dataItem, guid);
                }
                    break;
                default:
                    propertyInfo.SetValue(dataItem, cell.Value);
                    break;
            }
        }

 
    }
}
