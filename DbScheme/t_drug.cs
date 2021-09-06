using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ExcelTools
{
    /// <summary>
    /// 药品字典
    /// </summary>
    [Table("T_Base_Drug")]
    public class Drug 
    {
        public string id { get; set; }

        ///<summary>
        /// 药品编码
        ///</summary>
        [Description("药品编码")]
        [Required, MaxLength(100)]
        public string BM { get; set; }

        ///<summary>
        /// 通用名/化学名
        ///</summary>
        [Description("通用名/化学名")]
        [Required, MaxLength(200)]
        public string TYMC { get; set; }

        ///<summary>
        /// 商品名称
        ///</summary>
        [Description("商品名称")]
        [MaxLength(200)]
        public string SPMC { get; set; }

        ///<summary>
        /// 英文描述(英文名)
        ///</summary>
        [Description("英文描述(英文名)")]
        [MaxLength(200)]
        public string YWMC { get; set; }

        ///<summary>
        /// 通用名拼音码
        ///</summary>
        [Description("通用名拼音码")]
        [MaxLength(200)]
        public string PYM { get; set; }

        ///<summary>
        /// 汉语拼音
        ///</summary>
        [Description("汉语拼音")]
        [MaxLength(200)]
        public string HYPY { get; set; }

        ///<summary>
        /// 通用名五笔码
        ///</summary>
        [Description("通用名五笔码")]
        [MaxLength(200)]
        public string WBM { get; set; }

        /// <summary>
        /// 医嘱名称
        /// </summary>
        [NotMapped]
        public string YZMC { get; set; }

        /// <summary>
        /// 医嘱ID
        /// </summary>
        [Description("医嘱ID,T_Base_MedicalOrder/ID")]
        [Required, MaxLength(36)]
        public string YZID { get; set; }

        ///<summary>
        /// 厂家ID 
        ///</summary>
        [Description("厂家ID（修改,跟批次关联）,T_Base_DrugProduct/ID")]
        [MaxLength(36)]
        public string CJID { get; set; }

        ///<summary>
        /// 药品类型ID
        ///</summary>
        [Description("药品类型ID,T_Base_PresSecondType/ID")]
        [Required, MaxLength(36)]
        public string LXID { get; set; }

        /// <summary>
        /// 类型名称
        /// </summary>
        [NotMapped]
        public string LXMC { get; set; }

        ///<summary>
        /// 药品类型编码
        ///</summary>
        [Description("药品类型编码")]
        [MaxLength(36)]
        public string LXBM { get; set; }

        ///<summary>
        /// 药品子类型ID
        ///</summary>
        [Description("药品子类型ID,T_Base_PresThreeType/ID")]
        [Required, MaxLength(36)]
        public string ZLXID { get; set; }

        ///<summary>
        /// 药品子类型编码
        ///</summary>
        [Description("药品子类型编码")]
        [MaxLength(36)]
        public string ZLXBM { get; set; }

        ///<summary>
        /// 药品剂型ID
        ///</summary>
        [Description("药品剂型ID,T_Base_Dosage/ID")]
        [Required, MaxLength(36)]
        public string JXID { get; set; }

        ///<summary>
        /// 药品剂型编码
        ///</summary>
        [Description("药品剂型编码")]
        [MaxLength(36)]
        public string JXBM { get; set; }

        /// <summary>
        /// 剂型名称
        /// </summary>
        [NotMapped]
        public string JXMC { get; set; }

        ///<summary>
        /// 药理分类ID
        ///</summary>
        [Description("药理分类ID,T_Base_Pharmacology/ID")]
        [MaxLength(36)]
        public string YLFLID { get; set; }

        ///<summary>
        /// 费用分类大项目ID
        ///</summary>
        [Description("费用分类大项目ID,T_Base_FeeClass/ID")]
        [Required, MaxLength(36)]
        public string FYFLID { get; set; }

        ///<summary>
        /// 是否精神药 [0:非精神  1:精1  2：精2]
        ///</summary>
        [Required]
        [Description("0:非精神 1:精1 2：精2")]
        public int JSYPBS { get; set; }

        ///<summary>
        /// 国药准字号
        ///</summary>
        [Description("国药准字号")]
        [MaxLength(50)]
        public string GYZZH { get; set; }

        ///<summary>
        /// 进口国产 0-进口、1-国产
        ///</summary>
        [Description("进口国产 0-进口、1-国产")]
        public int? JKGC { get; set; }

        ///<summary>
        /// 药品有效期限（单位：月）
        ///</summary>
        [Description("药品有效期限（单位：月）")]
        public int? YXQ { get; set; }

        ///<summary>
        /// 计价单位名称(用量单位)
        ///</summary>
        [Description("计价单位名称(用量单位)")]
        [Required, MaxLength(50)]
        public string YLDW { get; set; }

        ///<summary>
        /// 0:总量取整、1:单次取整
        ///</summary>
        [Required, Description("0:总量取整、1:单次取整")]
        public int QZFS { get; set; }

        ///<summary>
        /// 门诊是否免发 
        ///</summary>
        [Description("门诊是否免发")]
        public bool MZMF { get; set; }

        ///<summary>
        /// 是否麻醉药 0-否 1-是
        ///</summary>
        [Required]
        [Description("是否麻醉药")]
        public bool MZYBS { get; set; }

        ///<summary>
        /// 是否毒性药 0-否 1-是
        ///</summary>
        [Required]
        [Description("是否毒性药")]
        public bool DXYBS { get; set; }

        ///<summary>
        /// 0:否 1:省基目录 2:国基目录 3:国、省基目录
        ///</summary>
        [Required]
        [Description("0:否 1:省基目录 2:国基目录 3:国、省基目录")]
        public int JBYYBS { get; set; }

        ///<summary>
        /// 是否皮试药 
        ///</summary>
        [Required]
        [Description("是否皮试药")]
        public bool PSYBS { get; set; }

        ///<summary>
        /// 是否处方药
        ///</summary>
        [Required]
        [Description("是否处方药")]
        public bool CFYBS { get; set; }

        ///<summary>
        /// 是否辅助用药
        ///</summary>
        [Required]
        [Description("是否辅助用药")]
        public bool FZYWBS { get; set; }

        /// <summary>
        /// 是否同步HIS
        /// </summary>
        [Description("是否同步HIS")]
        public bool BTBHIS { get; set; }

        ///<summary>
        /// 是否是特殊抗菌药物 
        ///</summary>
        [Description("是否是特殊抗菌药物")]
        public bool KJYWBS { get; set; }

        ///<summary>
        /// 限定日剂量
        ///</summary>
        [Description("DDD值")]
        [Column(TypeName = "decimal(9,2)")]
        [Range(0, double.MaxValue, ErrorMessage = "{0}为非负数")]
        public decimal? DDD { get; set; }

        ///<summary>
        /// 限定日剂量DDD值的药品规格转换系数
        ///</summary>
        [Description("限定日剂量DDD值的药品规格转换系数")]
        [Column(TypeName = "decimal(9,2)")]
        [Range(0, double.MaxValue, ErrorMessage = "{0}为非负数")]
        public decimal? DDDSL { get; set; }

        /// <summary>
        /// 药库规格
        /// </summary>
        [Description("药库规格")]
        public string YKGG { get; set; }

        /// <summary>
        /// 药库单位(入库单位)
        /// </summary>
        [Description("药库单位(入库单位)")]
        [MaxLength(20)]
        public string RKDW { get; set; }

        /// <summary>
        /// 药库系数 (入库单位系数)
        /// </summary>
        [Description("药库系数(入库单位系数)")]
        [Column("RKDWXS", TypeName = "decimal(9,3)")]
        [Range(0, double.MaxValue, ErrorMessage = "{0}为非负数")]
        public decimal? RKDWXS { get; set; }

        ///<summary>
        /// 药房规格
        ///</summary>
        [Description("药房规格")]
        [Required, MaxLength(50)]
        public string GG { get; set; }

        ///<summary>
        /// 药房单位(包装单位)
        ///</summary>
        [Description("药房单位(包装单位)")]
        [Required, MaxLength(20)]
        public string BZDW { get; set; }

        ///<summary>
        /// 基本系数(包装系数)
        ///</summary>
        [Description("基本系数(包装系数)")]
        [Required, Column("BZSL", TypeName = "decimal(9,3)")]
        [Range(0, double.MaxValue, ErrorMessage = "{0}为非负数")]
        public decimal BZSL { get; set; }

        ///<summary>
        /// 基本单位
        ///</summary>
        [Description("基本单位")]
        [Required, MaxLength(20)]
        public string JLDW { get; set; }

        ///<summary>
        /// 剂量系数
        ///</summary>
        [Description("剂量系数")]
        [Required, Column("JLSL", TypeName = "decimal(9,3)")]
        [Range(0, double.MaxValue, ErrorMessage = "{0}为非负数")]
        public decimal JLSL { get; set; }

        ///<summary>
        /// 剂量单位（最小单位）
        ///</summary>
        [Description("剂量单位,最小单位")]
        [Required, MaxLength(20)]
        public string ZXDW { get; set; }

        ///<summary>
        /// 用法ID
        ///</summary>
        [Description("用法ID,T_Base_Usage/ID")]
        [MaxLength(36)]
        public string YF { get; set; }

        ///<summary>
        /// 用法编码
        ///</summary>
        [Description("用法编码")]
        [MaxLength(36)]
        public string YFBM { get; set; }

        ///<summary>
        /// 频次ID
        ///</summary>
        [Description("频次ID,T_Base_Frequency/ID")]
        [MaxLength(36)]
        public string PC { get; set; }

        ///<summary>
        /// 频次编码
        ///</summary>
        [Description("频次编码")]
        [MaxLength(36)]
        public string PCBM { get; set; }

        ///<summary>
        /// 每次剂量(用量)默认
        ///</summary>
        [Description("每次剂量(用量)默认")]
        [Column("MCYL", TypeName = "decimal(6,2)")]
        [Range(0, double.MaxValue, ErrorMessage = "{0}为非负数")]
        public decimal MCYL { get; set; }

        /// <summary>
        /// 每次剂量(用量)单位默认
        /// </summary>
        [Description("每次剂量(用量)单位默认")]
        public string MCYLDW { get; set; }

        ///<summary>
        /// 疗程天数
        ///</summary>
        [Description("疗程天数")]
        [Range(0, Int32.MaxValue, ErrorMessage = "{0}为非负数")]
        [Column("TS")]
        public int? LCTS { get; set; }

        /// <summary>
        /// 用药说明
        /// </summary>
        [Description("用药说明")]
        [MaxLength(100)]
        public string YYSM { get; set; }

        ///<summary>
        /// 自定义编码
        ///</summary>
        [Description("自定义编码")]
        [MaxLength(50)]
        public string DM { get; set; }

        ///<summary>
        /// 是否贵重药 
        ///</summary>
        [Description("是否贵重药")]
        public bool? GZYBS { get; set; }

        ///<summary>
        /// 是否中标  
        ///</summary>
        [Description("是否中标")]
        public bool? ZBBS { get; set; }

        ///<summary>
        /// 抗生素级别 0-否 1-是 可分级
        ///</summary>
        [Description("抗生素级别 0-否 1-是 可分级")]
        public int? KSSBS { get; set; }

        ///<summary>
        /// 是否大输液 0否 1是
        ///</summary>
        [Description("是否大输液 0否1是")]
        public bool? DSYBS { get; set; }

        ///<summary>
        /// 国家药品价格监测系统药品目录YPID
        ///</summary>
        [Description("国家药品价格监测系统药品目录YPID")]
        [MaxLength(50)]
        public string GJBM { get; set; }

        ///<summary>
        /// 医保类型ID
        ///</summary>
        [Description("医保类型ID")]
        [MaxLength(36)]
        public string YBLX { get; set; }

        ///<summary>
        /// 农合比例%
        ///</summary>
        [Description("农合比例%")]
        public int? NHBL { get; set; }

        ///<summary>
        /// 医保项目编码
        ///</summary>
        [Description("医保项目编码")]
        [MaxLength(50)]
        public string YBXMBM { get; set; }

        ///<summary>
        /// 住院免发
        ///</summary>
        [Description("住院免发")]
        public int? ZYMF { get; set; }

        ///<summary>
        /// 药品说明书
        ///</summary>
        [Description("药品说明书")]
        [MaxLength(1000)]
        public string YPSMS { get; set; }

        ///<summary>
        /// 排序号
        ///</summary>
        [Description("排序号")]
        public int? PXH { get; set; }

        ///<summary>
        /// 药房科室ID 暂时没有用到
        ///</summary>
        [Description("药房科室ID")]
        [MaxLength(36)]
        public string YFKSID { get; set; }
    }
}
