using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;

namespace ExcelUtil.Model
{
    /// <summary>
    ///
    /// </summary>

    [Importer(HeaderRowIndex = 1)]
    public class ImportProductDto
    {
        /// <summary>
        ///     产品名称
        /// </summary>
        [ImporterHeader(Name = "产品名称", Description = "必填")]
        [Required(ErrorMessage = "产品名称是必填的")]
        [Display(Name = "产品名称")]
        public string Name { get; set; }

        /// <summary>
        ///     产品代码
        ///     长度验证
        /// </summary>
        [ImporterHeader(Name = "产品代码", Description = "最大长度为20", AutoTrim = false)]
        [Required(ErrorMessage = "产品代码是必填的")]
        [MaxLength(20, ErrorMessage = "产品代码最大长度为20（中文算两个字符）")]
        [Display(Name = "产品代码")]
        public string Code { get; set; }

        /// <summary>
        ///     产品条码
        /// </summary>
        [ImporterHeader(Name = "产品条码", FixAllSpace = true)]
        [MaxLength(10, ErrorMessage = "产品条码最大长度为10")]
        [RegularExpression(@"^\d*$", ErrorMessage = "产品条码只能是数字哦")]
        [Display(Name = "产品条码")]
        public string BarCode { get; set; }

        /// <summary>
        ///     申报价值
        /// </summary>
        [ImporterHeader(Name = "申报价值")]
        [Display(Name = "申报价值")]
        public double DeclareValue { get; set; }

        /// <summary>
        ///     尺寸
        /// </summary>
        [ImporterHeader(Name = "尺寸(长x宽x高)")]
        [Display(Name = "尺寸(长x宽x高)")]
        public string Size { get; set; }

        /// <summary>
        ///     重量（支持不设置ImporterHeader）
        /// </summary>
        //[ImporterHeader(Name = "重量(KG)")]
        [Display(Name = "重量(KG)")]
        public double? Weight { get; set; }

        /// <summary>
        ///     类型
        /// </summary>
        [ImporterHeader(Name = "类型")]
        [Display(Name = "类型")]
        public ImporterProductType Type { get; set; }
    }

    public class ExporterHeaderFilter : IExporterHeaderFilter
    {
        /// <summary>
        /// 表头筛选器（修改名称）
        /// </summary>
        /// <param name="exporterHeaderInfo"></param>
        /// <returns></returns>
        public ExporterHeaderInfo Filter(ExporterHeaderInfo exporterHeaderInfo)
        {
            if (exporterHeaderInfo.DisplayName.ToLower().Equals("name"))
            {
                exporterHeaderInfo.DisplayName = "名称";
            }
            return exporterHeaderInfo;
        }
    }

    /// <summary>
    /// 生成模板/导出DTO
    /// </summary>
    [Exporter(Name = "通用导出", MaxRowNumberOnASheet = 100, Author = "sailing", TableStyle = "Light10", ExporterHeaderFilter = typeof(ExporterHeaderFilter))]
    public class ExporterProductDto
    {
        /// <summary>
        ///     产品名称
        /// </summary>
        [ExporterHeader(DisplayName = "产品名称", IsBold = true)] // 导出数据
        [Required(ErrorMessage = "产品名称是必填的")]
        [Display(Name = "产品名称")] // 生成模板
        public string Name { get; set; }

        /// <summary>
        ///     产品代码
        ///     长度验证
        /// </summary>
        [ExporterHeader(DisplayName = "产品代码", IsBold = true)]
        [Required(ErrorMessage = "产品代码是必填的")]
        [MaxLength(20, ErrorMessage = "产品代码最大长度为20（中文算两个字符）")]
        [Display(Name = "产品代码")]
        public string Code { get; set; }

        /// <summary>
        ///     产品条码
        /// </summary>
        [ExporterHeader(DisplayName = "产品条码", IsBold = true)]
        [MaxLength(10, ErrorMessage = "产品条码最大长度为10")]
        [RegularExpression(@"^\d*$", ErrorMessage = "产品条码只能是数字")]
        [Display(Name = "产品条码")]
        public string BarCode { get; set; }

        /// <summary>
        ///     类型
        /// </summary>
        [ExporterHeader(DisplayName = "类型", IsBold = true)]
        [Display(Name = "类型")]
        public ImporterProductType Type { get; set; }
    }

    /// <summary>
    /// 生成模板/导出DTO
    /// </summary>
    [Exporter(Name = "通用导出2", MaxRowNumberOnASheet = 100, Author = "sailing", TableStyle = "Light10", ExporterHeaderFilter = typeof(ExporterHeaderFilter))]
    public class ExporterProductDto2
    {
        /// <summary>
        ///     产品名称
        /// </summary>
        [ExporterHeader(DisplayName = "产品名称", IsBold = true)] // 导出数据
        [Required(ErrorMessage = "产品名称是必填的")]
        [Display(Name = "产品名称2")] // 生成模板
        public string Name { get; set; }

        /// <summary>
        ///     产品代码
        ///     长度验证
        /// </summary>
        [ExporterHeader(DisplayName = "产品代码", IsBold = true)]
        [Required(ErrorMessage = "产品代码是必填的")]
        [MaxLength(20, ErrorMessage = "产品代码最大长度为20（中文算两个字符）")]
        [Display(Name = "产品代码")]
        public string Code { get; set; }

        /// <summary>
        ///     产品条码
        /// </summary>
        [ExporterHeader(DisplayName = "产品条码", IsBold = true)]
        [MaxLength(10, ErrorMessage = "产品条码最大长度为10")]
        [RegularExpression(@"^\d*$", ErrorMessage = "产品条码只能是数字")]
        [Display(Name = "产品条码")]
        public string BarCode { get; set; }

        /// <summary>
        ///     类型
        /// </summary>
        [ExporterHeader(DisplayName = "类型", IsBold = true)]
        [Display(Name = "类型")]
        public ImporterProductType Type { get; set; }
    }

    /// <summary>
    /// 生成模板/导出DTO
    /// </summary>
    [Exporter(Name = "通用导出3", MaxRowNumberOnASheet = 100, Author = "sailing", TableStyle = "Light10", ExporterHeaderFilter = typeof(ExporterHeaderFilter))]
    public class ExporterProductDto3
    {
        /// <summary>
        ///     产品名称
        /// </summary>
        [ExporterHeader(DisplayName = "产品名称", IsBold = true)] // 导出数据
        [Required(ErrorMessage = "产品名称是必填的")]
        [Display(Name = "产品名称2")] // 生成模板
        public string Name { get; set; }

        /// <summary>
        ///     产品代码
        ///     长度验证
        /// </summary>
        [ExporterHeader(DisplayName = "产品代码", IsBold = true)]
        [Required(ErrorMessage = "产品代码是必填的")]
        [MaxLength(20, ErrorMessage = "产品代码最大长度为20（中文算两个字符）")]
        [Display(Name = "产品代码")]
        public string Code { get; set; }

        /// <summary>
        ///     产品条码
        /// </summary>
        [ExporterHeader(DisplayName = "产品条码", IsBold = true)]
        [MaxLength(10, ErrorMessage = "产品条码最大长度为10")]
        [RegularExpression(@"^\d*$", ErrorMessage = "产品条码只能是数字")]
        [Display(Name = "产品条码")]
        public string BarCode { get; set; }

        /// <summary>
        ///     类型
        /// </summary>
        [ExporterHeader(DisplayName = "类型", IsBold = true)]
        [Display(Name = "类型")]
        public ImporterProductType Type { get; set; }
    }


    public enum ImporterProductType
    {
        [Display(Name = "第一")] One = 0,
        [Display(Name = "第二")] Two = 1
    }

    /// <summary>
    /// 教材订购信息
    /// </summary>
    public class TextbookOrderInfo
    {
        /// <summary>
        /// 公司名称
        /// </summary>
        public string Company { get; }

        /// <summary>
        /// 地址
        /// </summary>
        public string Address { get; }

        /// <summary>
        /// 联系人
        /// </summary>
        public string Contact { get; }

        /// <summary>
        /// 电话
        /// </summary>
        public string Tel { get; }

        /// <summary>
        /// 制表人
        /// </summary>
        public string Watchmaker { get; }

        /// <summary>
        /// 时间
        /// </summary>
        public string Time { get; }

        public string ImageUrl { get; set; }

        /// <summary>
        /// 教材信息列表
        /// </summary>
        public List<BookInfo> BookInfos { get; }

        public TextbookOrderInfo(string company, string address, string contact, string tel, string watchmaker, string time, string imageUrl, List<BookInfo> bookInfo)
        {
            Company = company;
            Address = address;
            Contact = contact;
            Tel = tel;
            Watchmaker = watchmaker;
            Time = time;
            ImageUrl = imageUrl;
            BookInfos = bookInfo;
        }
    }

    /// <summary>
    /// 教材信息
    /// </summary>
    public class BookInfo
    {
        /// <summary>
        /// 行号
        /// </summary>
        public int RowNo { get; }

        /// <summary>
        /// 书号
        /// </summary>
        public string No { get; }

        /// <summary>
        /// 书名
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// 主编
        /// </summary>
        public string EditorInChief { get; }

        /// <summary>
        /// 出版社
        /// </summary>
        public string PublishingHouse { get; }

        /// <summary>
        /// 定价
        /// </summary>
        public string Price { get; }

        /// <summary>
        /// 采购数量
        /// </summary>
        public int PurchaseQuantity { get; }

        /// <summary>
        /// 备注
        /// </summary>
        public string Remark { get; }

        /// <summary>
        /// 封面
        /// </summary>
        public string Cover { get; set; }

        public BookInfo(int rowNo, string no, string name, string editorInChief, string publishingHouse, string price, int purchaseQuantity, string remark)
        {
            RowNo = rowNo;
            No = no;
            Name = name;
            EditorInChief = editorInChief;
            PublishingHouse = publishingHouse;
            Price = price;
            PurchaseQuantity = purchaseQuantity;
            Remark = remark;
        }
    }
}