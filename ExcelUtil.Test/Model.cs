using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;

namespace ExcelUtil.Model
{





    /// <summary>
    /// 档案模板
    /// </summary>
 
    public class ArchiveTemplateDto
    {



        [ImporterHeader(Name = "100",  ColumnIndex = 1)] public string Name100 { get; set; }
        [ImporterHeader(Name = "101",  ColumnIndex = 2)] public string Name101 { get; set; }
        [ImporterHeader(Name = "102",  ColumnIndex = 3)] public string Name102 { get; set; }
        [ImporterHeader(Name = "103",  ColumnIndex = 4)] public string Name103 { get; set; }
        [ImporterHeader(Name = "104",  ColumnIndex = 5)] public string Name104 { get; set; }
        [ImporterHeader(Name = "105",   ColumnIndex = 6)] public string Name105 { get; set; }
        [ImporterHeader(Name = "106", ColumnIndex = 7)] public string Name106 { get; set; }
        [ImporterHeader(Name = "107", ColumnIndex = 8)] public string Name107 { get; set; }
        [ImporterHeader(Name = "108", ColumnIndex = 9)] public string Name108 { get; set; }
        [ImporterHeader(Name = "109", ColumnIndex = 10)] public string Name109 { get; set; }
        [ImporterHeader(Name = "110", ColumnIndex = 11)] public string Name110 { get; set; }
        [ImporterHeader(Name = "111", ColumnIndex = 12)] public string Name111 { get; set; }
        [ImporterHeader(Name = "112", ColumnIndex = 13)] public string Name112 { get; set; }
        [ImporterHeader(Name = "113", ColumnIndex = 14)] public string Name113 { get; set; }
        [ImporterHeader(Name = "114", ColumnIndex = 15)] public string Name114 { get; set; }
        [ImporterHeader(Name = "115", ColumnIndex = 16)] public string Name115 { get; set; }
        [ImporterHeader(Name = "116", ColumnIndex = 17)] public string Name116 { get; set; }
        [ImporterHeader(Name = "117", ColumnIndex = 18)] public string Name117 { get; set; }
        [ImporterHeader(Name = "118", ColumnIndex = 19)] public string Name118 { get; set; }
        [ImporterHeader(Name = "119", ColumnIndex = 20)] public string Name119 { get; set; }
        [ImporterHeader(Name = "120", ColumnIndex = 21)] public string Name120 { get; set; }
        [ImporterHeader(Name = "121", ColumnIndex = 22)] public string Name121 { get; set; }
        [ImporterHeader(Name = "122", ColumnIndex = 23)] public string Name122 { get; set; }
        [ImporterHeader(Name = "123", ColumnIndex = 24)] public string Name123 { get; set; }
        [ImporterHeader(Name = "124", ColumnIndex = 25)] public string Name124 { get; set; }
        [ImporterHeader(Name = "125", ColumnIndex = 26)] public string Name125 { get; set; }
        [ImporterHeader(Name = "126", ColumnIndex = 27)] public string Name126 { get; set; }
        [ImporterHeader(Name = "127", ColumnIndex = 28)] public string Name127 { get; set; }
        [ImporterHeader(Name = "128", ColumnIndex = 29)] public string Name128 { get; set; }



    }



    /// <summary>
    /// 测试表头位置
    /// </summary>

    public class ImportProductDto
    {
     
        /// <summary>
        ///     2
        ///      
        /// </summary>
        [ImporterHeader(Name = "面积",ColumnIndex =1)]
        
        public string Name { get; set; }


        /// <summary>
        ///     3
        /// </summary>
        [ImporterHeader(Name = "库容", ColumnIndex = 2)]
        
        public string Code3 { get; set; }
        /// <summary>
        ///     4
        /// </summary>
        [ImporterHeader(Name = "列", ColumnIndex = 3)]
        
        public string Code4 { get; set; }
    }



    ///// <summary>
    ///// 测试表头位置
    ///// </summary>
    //[Importer(HeaderRowIndex = 3)]
    //public class ImportProductDto
    //{
    //    /// <summary>
    //    ///     产品名称
    //    /// </summary>
    //    [ImporterHeader(Name = "产品名称", Description = "必填")]
    //    [Required(ErrorMessage = "产品名称是必填的")]
    //    public string Name { get; set; }

    //    /// <summary>
    //    ///     产品代码
    //    ///     长度验证
    //    /// </summary>
    //    [ImporterHeader(Name = "产品代码", Description = "最大长度为20", AutoTrim = false)]
    //    [MaxLength(20, ErrorMessage = "产品代码最大长度为20（中文算两个字符）")]
    //    public string Code { get; set; }

    //    /// <summary>
    //    ///  测试GUID
    //    /// </summary>
    //    public Guid ProductIdTest1 { get; set; }

    //    public Guid? ProductIdTest2 { get; set; }

    //    /// <summary>
    //    ///     产品条码
    //    /// </summary>
    //    [ImporterHeader(Name = "产品条码", FixAllSpace = true)]
    //    [MaxLength(10, ErrorMessage = "产品条码最大长度为10")]
    //    [RegularExpression(@"^\d*$", ErrorMessage = "产品条码只能是数字")]
    //    public string BarCode { get; set; }

    //    /// <summary>
    //    ///     客户Id
    //    /// </summary>
    //    [ImporterHeader(Name = "客户代码")]
    //    public long ClientId { get; set; }

    //    /// <summary>
    //    ///     产品型号
    //    /// </summary>
    //    [ImporterHeader(Name = "产品型号")]
    //    public string Model { get; set; }

    //    /// <summary>
    //    ///     申报价值
    //    /// </summary>
    //    [ImporterHeader(Name = "申报价值")]
    //    public double DeclareValue { get; set; }

    //    /// <summary>
    //    ///     货币单位
    //    /// </summary>
    //    [ImporterHeader(Name = "货币单位")]
    //    public string CurrencyUnit { get; set; }

    //    /// <summary>
    //    ///     品牌名称
    //    /// </summary>
    //    [ImporterHeader(Name = "品牌名称")]
    //    public string BrandName { get; set; }

    //    /// <summary>
    //    ///     尺寸
    //    /// </summary>
    //    [ImporterHeader(Name = "尺寸(长x宽x高)")]
    //    public string Size { get; set; }

    //    /// <summary>
    //    ///     重量（支持不设置ImporterHeader）
    //    /// </summary>
    //    //[ImporterHeader(Name = "重量(KG)")]
    //    [Display(Name = "重量(KG)")]
    //    public double? Weight { get; set; }

    //    /// <summary>
    //    ///     类型
    //    /// </summary>
    //    [ImporterHeader(Name = "类型")]
    //    public ImporterProductType Type { get; set; }

    //    /// <summary>
    //    ///     是否行
    //    /// </summary>
    //    [ImporterHeader(Name = "是否行")]
    //    public bool IsOk { get; set; }

    //    [ImporterHeader(Name = "公式测试")] public DateTime FormulaTest { get; set; }

    //    /// <summary>
    //    ///     身份证
    //    ///     多个错误测试
    //    /// </summary>
    //    [ImporterHeader(Name = "身份证")]
    //    [RegularExpression(@"(^\d{15}$)|(^\d{18}$)|(^\d{17}(\d|X|x)$)", ErrorMessage = "身份证号码无效！")]
    //    [StringLength(18, ErrorMessage = "身份证长度不得大于18！")]
    //    public string IdNo { get; set; }

    //    [Display(Name = "性别")] public string Sex { get; set; }
    //}



    /// <summary>
    /// 生成模板/导出DTO
    /// </summary>
    [Exporter(Name = "通用导出", MaxRowNumberOnASheet = 100, Author = "sailing", TableStyle = "Light10")]
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
        public DateTime? Date { get; set; }

        /// <summary>
        ///     类型
        /// </summary>
        //[ExporterHeader(DisplayName = "类型", IsBold = true)]
        //[Display(Name = "类型")]
       // public ImporterProductType Type { get; set; }
    }

    //public enum ImporterProductType
    //{
    //    [Display(Name = "第一")] One = 0,
    //    [Display(Name = "第二")] Two = 1
    //}

}