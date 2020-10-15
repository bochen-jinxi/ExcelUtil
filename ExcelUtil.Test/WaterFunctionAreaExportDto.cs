using Magicodes.ExporterAndImporter.Core;
using System;
using System.ComponentModel.DataAnnotations;

namespace SaiLing.LuXian.BaseData.Api.Services.ExportDtos.BasicInfo
{
    /// <summary>
    /// 水功能区信息数据导出对象
    /// </summary>
    [Exporter(Name = "水功能区信息数据导出", MaxRowNumberOnASheet = 100000, Author = "赛零信息 · Saing", TableStyle = "Light10")]
    public class WaterFunctionAreaExportDto
    {
        /// <summary>
        /// 水功能区代码
        /// </summary>
        [Required(ErrorMessage = "水功能区代码不能为空")]
        [StringLength(18)]
        [Display(Name = "水功能区代码")]
        public string Code { get; set; }

        /// <summary>
        /// 水功能区名称
        /// </summary>
        [Required(ErrorMessage = "水功能区名称不能为空")]
        [StringLength(50)]
        [Display(Name = "水功能区名称")]
        public string Name { get; set; }

        /// <summary>
        /// 水功能区起始断面
        /// </summary>
        [StringLength(100)]
        [Display(Name = "水功能区起始断面")]
        public string StartSection { get; set; }

        /// <summary>
        /// 水功能区终止断面
        /// </summary>
        [StringLength(100)]
        [Display(Name = "水功能区终止断面")]
        public string EndSection { get; set; }

        /// <summary>
        /// 水功能区长度
        /// </summary>
        [Display(Name = "水功能区长度")]
        public decimal Distence { get; set; }

        /// <summary>
        /// 水功能区面积
        /// </summary>
        [Display(Name = "水功能区面积")]
        public decimal Area { get; set; }

        /// <summary>
        /// 水功能区功能排序
        /// </summary>
        [StringLength(40)]
        [Display(Name = "水功能区功能排序")]
        public string FunctionalSequences { get; set; }

        /// <summary>
        /// 地下水层位
        /// </summary>
        [StringLength(20)]
        [Display(Name = "地下水层位")]
        public string GroundWaterHorzon { get; set; }

        /// <summary>
        /// 矿化度
        /// </summary>
        [Display(Name = "矿化度    ")]
        public decimal MineralizationDegree { get; set; }

        /// <summary>
        /// 年均总补给量
        /// </summary>
        [Display(Name = "年均总补给量  ")]
        public decimal AnnualTotalSupply { get; set; }

        /// <summary>
        /// 年均可开采量
        /// </summary>
        [Display(Name = "年均可开采量  ")]
        public decimal AnnualPredictProduction { get; set; }

        /// <summary>
        /// 年均实开采量
        /// </summary>
        [Display(Name = "年均实开采量")]
        public decimal AnnualActualProduction { get; set; }

        /// <summary>
        /// 开始考核日期
        /// </summary>
        [Display(Name = "开始考核日期")]
        public DateTime? BeginExam { get; set; }//todo:导出可空时间类型异常

        /// <summary>
        /// 备注
        /// </summary>
        [StringLength(256)]
        [Display(Name = "备注")]
        public string Remark { get; set; }

        #region 追加字典数据

        /// <summary>
        /// 所属行政区划字典数据
        /// </summary>
        [Display(Name = "所属行政区划字典数据")]
        [StringLength(1000)]
        public string DivisionDir { get; set; }

        /// <summary>
        /// 管理单位字典数据
        /// </summary>
        [Display(Name = "管理单位字典数据")]
        [StringLength(1000)]
        public string UnitDir { get; set; }

        /// <summary>
        /// 水功能区水质目标字典数据
        /// </summary>
        [Display(Name = "水功能区水质目标字典数据")]
        [StringLength(1000)]
        public string WaterQulityGoalsDir { get; set; }

        /// <summary>
        /// 水功能区营养状态目标字典数据
        /// </summary>
        [Display(Name = "水功能区营养状态目标字典数据")]
        [StringLength(1000)]
        public string NutritionStatusTargetDir { get; set; }

        /// <summary>
        /// 主导功能字典数据
        /// </summary>
        [Display(Name = "主导功能字典数据")]
        [StringLength(1000)]
        public string LeadingFunctionDir { get; set; }

        /// <summary>
        /// 水功能区等级字典数据
        /// </summary>
        [Display(Name = "水功能区等级字典数据")]
        [StringLength(1000)]
        public string LevelDir { get; set; }

        /// <summary>
        /// 地貌类别字典数据
        /// </summary>
        [Display(Name = "地貌类别字典数据")]
        [StringLength(1000)]
        public string GeomorphicCategoryDir { get; set; }

        /// <summary>
        /// 水功能区类型字典数据
        /// </summary>
        [Display(Name = "水功能区类型字典数据")]
        [StringLength(1000)]
        public string TypeDir { get; set; }

        /// <summary>
        /// 监控级别字典数据
        /// </summary>
        [Display(Name = "监控级别字典数据")]
        [StringLength(1000)]
        public string MonitorLevelDir { get; set; }

        /// <summary>
        /// 是否考核字典数据
        /// </summary>
        [Display(Name = "是否考核字典数据")]
        [StringLength(1000)]
        public string IsExamDir { get; set; }

        #endregion 追加字典数据

        #region 字典数据映射

        /// <summary>
        /// 所属行政区划
        /// </summary>
        [Display(Name = "所属行政区划")]
        [StringLength(1000)]
        public string Division
        {
            get { return DivisionDir; }
        }

        /// <summary>
        /// 管理单位
        /// </summary>
        [Display(Name = "管理单位")]
        [StringLength(1000)]
        public string Unit
        {
            get { return UnitDir; }
        }

        /// <summary>
        /// 水功能区水质目标
        /// </summary>
        [Display(Name = "水功能区水质目标")]
        [StringLength(1000)]
        public string WaterQulityGoals
        {
            get { return WaterQulityGoalsDir; }
        }

        /// <summary>
        /// 水功能区营养状态目标
        /// </summary>
        [Display(Name = "水功能区营养状态目标")]
        [StringLength(1000)]
        public string NutritionStatusTarget
        {
            get { return NutritionStatusTargetDir; }
        }

        /// <summary>
        /// 主导功能
        /// </summary>
        [Display(Name = "主导功能")]
        [StringLength(1000)]
        public string LeadingFunction
        {
            get { return LeadingFunctionDir; }
        }

        /// <summary>
        /// 水功能区等级
        /// </summary>
        [Display(Name = "水功能区等级")]
        [StringLength(1000)]
        public string Level
        {
            get { return LevelDir; }
        }

        /// <summary>
        /// 地貌类别
        /// </summary>
        [Display(Name = "地貌类别")]
        [StringLength(1000)]
        public string GeomorphicCategory
        {
            get { return GeomorphicCategoryDir; }
        }

        /// <summary>
        /// 水功能区类型
        /// </summary>
        [Display(Name = "水功能区类型")]
        [StringLength(1000)]
        public string Type
        {
            get { return TypeDir; }
        }

        /// <summary>
        /// 监控级别
        /// </summary>
        [Display(Name = "监控级别")]
        [StringLength(1000)]
        public string MonitorLevel
        {
            get { return MonitorLevelDir; }
        }

        /// <summary>
        /// 是否考核
        /// </summary>
        [Display(Name = "是否考核")]
        [StringLength(1000)]
        public string IsExam
        {
            get { return IsExamDir; }
        }

        #endregion 字典数据映射
    }
}