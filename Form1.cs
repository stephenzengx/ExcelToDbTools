using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Dapper;
using Excel;
using Excel.OpenXml;
using Excel.Utils;
using MySql.Data.MySqlClient;
using static System.Int32;
using System.Xml.Linq;
using ExcelTools.SqlScheme;

namespace ExcelTools
{
    public partial class Form1 : Form
    {
        #region 变量
        protected static DateTime? LastImportTime=null;
        protected static DateTime? LastValidateTime=null;
        protected static IConfigurationRoot Config = Utils.Config;
        protected static bool HasPassFileValid;
        protected static int Seconds = Utils.Config["SencondLimits"].ToInt32();

        protected static string JGID = string.Empty; //机构id

        protected static int RemainErrorCount;
        protected static int MaxErrorCount;
        protected static List<FileInfo> Files = new List<FileInfo>();
        // [sheetname : dbname.tbname] 
        protected static Dictionary<string, string> SheetTbMapDic = new Dictionary<string, string>();
        // [dbname.tbname : list]  
        protected static Dictionary<string, List<TbDesc>> DbTbDescDic = new Dictionary<string, List<TbDesc>>();
        //dbname : [dbname.tbname : list]
        protected static Dictionary<string, Dictionary<string, List<ScanExcelHeadDesc>>> ScanDescDic = new Dictionary<string, Dictionary<string, List<ScanExcelHeadDesc>>>();
        //唯一数据 [db.tbname.uniqFieldName : list]
        protected static Dictionary<string, List<string>> UniqDataDic = new Dictionary<string, List<string>>();
        //关联基础数据数据 [dbname.tbname.keyFieldname.valueFieldname : [key,value] ]
        protected static Dictionary<string, Dictionary<string, string>> RelatedDataDic = new Dictionary<string, Dictionary<string, string>>();
        //[dbname.tbname, obj]
        protected static Dictionary<string, ReadExcelDataRet> ExecSqlFinalDic = new Dictionary<string, ReadExcelDataRet>();
        //[db,list] 导入前清空表字典
        protected static Dictionary<string, List<string>> ClearTbDic = new Dictionary<string, List<string>>();
        #endregion
        public Form1()
        {
            InitializeComponent();

            TxtDirPath.Text = ReadPathFromXml();

            OrgsInit();
        }

        /// <summary>
        /// 重置状态
        /// </summary>
        private void ResetStatus()
        {
            TxtDirPath.Enabled = true;
            DrpDwnOrgs.Enabled = true;
            TxtMaxErrorCount.Enabled = true;
            HasPassFileValid = false;
        }

        private string ReadPathFromXml()
        {
            /*
             编写程序的时候，经常需要用的项目根目录。自己总结如下
             1、取得控制台应用程序的根目录方法
                 方法1、Environment.CurrentDirectory 取得或设置当前工作目录的完整限定路径
                 方法2、AppDomain.CurrentDomain.BaseDirectory 获取基目录，它由程序集冲突解决程序用来探测程序集
             2、取得Web应用程序的根目录方法
                 方法1、HttpRuntime.AppDomainAppPath.ToString();//获取承载在当前应用程序域中的应用程序的应用程序目录的物理驱动器路径。用于App_Data中获取
                 方法2、Server.MapPath("") 或者 Server.MapPath("~/");//返回与Web服务器上的指定的虚拟路径相对的物理文件路径
                 方法3、Request.ApplicationPath;//获取服务器上ASP.NET应用程序的虚拟应用程序根目录
             3、取得WinForm应用程序的根目录方法  下面5个都是这个 //E:\MyProject\ExcelToDbTools\bin\Debug\netcoreapp3.1
                 1、Environment.CurrentDirectory.ToString();//获取或设置当前工作目录的完全限定路径
                 2、Application.StartupPath.ToString();//获取启动了应用程序的可执行文件的路径，不包括可执行文件的名称
                 3、Directory.GetCurrentDirectory();//获取应用程序的当前工作目录
                 4、AppDomain.CurrentDomain.BaseDirectory;//获取基目录，它由程序集冲突解决程序用来探测程序集
                 5、AppDomain.CurrentDomain.SetupInformation.ApplicationBase;//获取或设置包含该应用程序的目录的名称
             */
            XDocument xdoc = XDocument.Load(Path.Combine(Utils.WorkDirPath, "AppSettings.xml"));
            XElement xeleRoot = xdoc.Root;
            XElement xelePath = xeleRoot.Elements("FilePath").Single();

            return xelePath.Value;
        }

        private void SavePathToXml(string absPath)
        {
            XDocument xdoc = XDocument.Load(Path.Combine(Utils.WorkDirPath, "AppSettings.xml"));
            XElement xeleRoot = xdoc.Root;
            XElement xelePath = xeleRoot.Elements("FilePath").Single();
            
            //xelePath.ReplaceWith(new XElement("FilePath", absPath));
            xelePath.SetValue(absPath);

            xdoc.Save(Path.Combine(Utils.WorkDirPath, "AppSettings.xml"));
        }



        /// <summary>
        /// 点击开始导入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImportBtn_Click(object sender, EventArgs e)
        {
            //其他按钮禁止修改
            TxtDirPath.Enabled = false;
            DrpDwnOrgs.Enabled = false;
            TxtMaxErrorCount.Enabled = false;

            #region 校验
            if (!HasPassFileValid)
            {
                DiagTip("文件校验未通过，请先校验！");
                ResetStatus();
                return;
            }

            if (string.IsNullOrEmpty(JGID))
            {
                DiagTip("请先选择机构！");
                ResetStatus();
                return;
            }

            if (LastImportTime != null && (DateTime.Now - LastImportTime)?.Seconds <= Seconds)
            {
                DiagTip($"操作频率过快,{Seconds}s后再试！");
                ResetStatus();
                return;
            }

            if (RemainErrorCount <= 0)
            {
                Utils.LogInfo("错误记录数已达到阈值，请检查Excel文件！");
                ResetStatus();
                return;
            }
            #endregion

            LastImportTime = DateTime.Now;
            //导入前清空
            if (ClearTbDic.Count > 0)
            {
                //弹窗
                DialogResult dr = MessageBox.Show("导入前会清空有关表数据，是否继续?", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (dr == DialogResult.OK)
                {
                    if (!ClearBeforeImport())
                    {
                        Utils.LogInfo("删除异常,程序结束！....");
                        ResetStatus();
                        return;
                    }
                }
                else
                {
                    ResetStatus();
                    return;
                }
            }

            Utils.LogInfo(GetLineMsg("正在开始导入数据....",true),false);
            ComfirmImportData();
            HasPassFileValid = false;
            ResetStatus();
            ValidateInit();
        }

        /// <summary>
        /// 导入前清空
        /// </summary>
        /// <returns></returns>
        private bool ClearBeforeImport()
        {
            Utils.LogInfo(GetLineMsg("开始清空表数据！....",true), false);
            var ret = true;
            using (IDbConnection con = new MySqlConnection(Config.GetConnectionString(ClearTbDic.FirstOrDefault().Key)))
            {
                foreach (var dbItem in ClearTbDic)
                {
                    con.Open();
                    IDbTransaction transaction = con.BeginTransaction();
                    try
                    {
                        //开始导入数据 dapper 事务
                        foreach (var tbName in dbItem.Value)
                        {
                            con.Execute($"delete from {tbName}", null, transaction);
                            Utils.LogInfo($"{dbItem.Key}.{tbName}表, 清空成功....");
                        }

                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        ret = false;
                        Utils.LogInfo($"批量删除失败,请检查！....");
                        Utils.LogInfo(ex.Message + "\r\n" + ex.StackTrace);

                        transaction.Rollback();
                    }
                }
            }

            ClearTbDic = new Dictionary<string, List<string>>();
            Utils.LogInfo(GetLineMsg("清空表数据成功！....", false), false);

            return ret;
        }

        /// <summary>
        /// 开始导入（确认导入）
        /// </summary>
        private void ComfirmImportData()
        {
            //直接拿某一个连接字符串，因为表名前都加了 db名称
            using (IDbConnection con = new MySqlConnection(Config.GetConnectionString(ExecSqlFinalDic.FirstOrDefault().Value.DbName)))
            {
                con.Open();
                IDbTransaction transaction = con.BeginTransaction();
                try
                {
                    //开始导入数据 dapper 事务
                    foreach (var sqlItem in ExecSqlFinalDic)
                    {
                        if (!sqlItem.Value.ExecParams.Any())
                        {
                            Utils.LogInfo($"{sqlItem.Key} 数据为空,跳过..");
                            continue;
                        }

                        Utils.LogInfo($"开始插入表:{sqlItem.Key}....");
                        con.Execute(sqlItem.Value.ExecSql, sqlItem.Value.ExecParams, transaction);
                    }

                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    Utils.LogInfo("导入失败,请检查！....");
                    Utils.LogInfo(ex.Message + "\r\n" + ex.StackTrace);

                    transaction.Rollback();
                    return;
                }
            }

            Utils.LogInfo(GetLineMsg("导入成功，程序结束！....",false), false);
        }

        /// <summary>
        /// 校验前初始化
        /// </summary>
        public void ValidateInit()
        {
            HasPassFileValid = false;
            MaxErrorCount = 0;
            Files = new List<FileInfo>();
            ClearTbDic = new Dictionary<string, List<string>>();
            SheetTbMapDic = new Dictionary<string, string>();
            DbTbDescDic = new Dictionary<string, List<TbDesc>>();
            ScanDescDic = new Dictionary<string, Dictionary<string, List<ScanExcelHeadDesc>>>();
            ExecSqlFinalDic = new Dictionary<string, ReadExcelDataRet>();
            UniqDataDic = new Dictionary<string, List<string>>();
            RelatedDataDic = new Dictionary<string, Dictionary<string, string>>();

            //如果输入路径和上次路径不同，记录到xml里面
            var inputPath = TxtDirPath.Text;
            var xmlPath = ReadPathFromXml();
            if(!xmlPath.Equals(inputPath))
                SavePathToXml(inputPath);
        }

        /// <summary>
        /// 点击校验按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnValidate_Click(object sender, EventArgs e)
        {
            try
            {
                #region 基础校验
                if (DrpDwnOrgs.Items == null || DrpDwnOrgs.Items.Count <= 0)
                {
                    DiagTip("请先导入机构!");
                    return;
                }

                if (DrpDwnOrgs.SelectedValue == null)
                {
                    DiagTip("请先选择机构!");
                    return;
                }

                JGID = DrpDwnOrgs.SelectedValue.ToString();
                if (LastImportTime != null && (DateTime.Now - LastValidateTime)?.Seconds <= Seconds)
                {
                    DiagTip($"操作频率过快,{Seconds}s后再试！");
                    return;
                }

                ValidateInit();//数据初始化
                LastValidateTime = DateTime.Now;
                if (!ValidateBase(TxtDirPath.Text, TxtMaxErrorCount.Text, ref Files, out MaxErrorCount))//初始化操作
                    return;
                RemainErrorCount = MaxErrorCount;
                #endregion

                #region 文件名回显
                PanelExcelName.Controls.Clear();
                TextBox box1 = new TextBox
                {
                    Multiline = true,
                    Width = PanelExcelName.Width,
                    Height = (int)(PanelExcelName.Height * 0.80),//files.Length * 40,
                    BorderStyle = BorderStyle.None,
                    BackColor = SystemColors.Control,
                };

                for (int i = 0; i < Files.Count; i++)
                {
                    if (Files[i].Name.Contains("$"))
                        continue;
                    box1.AppendText($"{Files[i].Name}" + ((i + 1 == Files.Count) ? string.Empty : Environment.NewLine));
                }
                //添加到窗体
                PanelExcelName.Controls.Add(box1);

                #endregion

                #region 校验文件
                Utils.LogInfo(GetLineMsg("开始校验文件", true),false);
                foreach (var file in Files)
                {
                    //f.fullName 绝对路径 放在hidden Field里面
                    var dbStrArr = file.Name.Split(Config[EnumIdentifier.SpitChar.ToString()]);
                    if (dbStrArr.Length != 2)
                    {
                        Utils.LogInfo("Excel文件名称格式有误,请检查");
                        return;
                    }

                    var dbName = dbStrArr[0].ToLower();
                    var connectionString = Utils.Config.GetConnectionString(dbStrArr[0]);
                    if (string.IsNullOrEmpty(connectionString))
                    {
                        Utils.LogInfo("Excel文件对应数据库不存在,请检查");
                        return;
                    }

                    using (IDbConnection con = new MySqlConnection(connectionString))
                    {
                        con.Open();

                        con.GetTableNamesAndFieldDic(dbName, ref DbTbDescDic);
                        if (DbTbDescDic.Count <= 0)
                            return;
                    }
                }
                Utils.LogInfo(GetLineMsg("校验文件 通过！", false), false);
                #endregion

                #region 校验 Excel表头
                Utils.LogInfo(GetLineMsg("开始校验Excel表头", true), false);
                foreach (var file in Files)
                {
                    var dbStrArr = file.Name.Split(Config[EnumIdentifier.SpitChar.ToString()]);

                    var dbName = dbStrArr[0];
                    var cleartbNameList = new List<string>();
                    var sheetNames = MyMiniExcel.GetSheetNames(file.FullName);
                    Utils.LogInfo($"----开始检查Excel表 '{file.Name}' ----");

                    var tbNameList = new List<string>();//判断表名是否重复

                    Dictionary<string, List<ScanExcelHeadDesc>> dic = new Dictionary<string, List<ScanExcelHeadDesc>>();
                    using (FileStream stream = Helpers.OpenSharedRead(file.FullName))
                    {
                        var xmlSheetReader = new ExcelOpenXmlSheetReader(stream);

                        foreach (var sheetName in sheetNames)
                        {
                            Utils.LogInfo($"开始检查sheet表 '{sheetName}': ");
                            /*
                             1- 检查数据库是否存在
                             2- 检查字段名是否存在                 
                             3- 检查每个字段数据格式是否正确 (数据类型,值范围，长度范围), 不正确记录数加1，判断阈值
                             */
                            var sheetSpitArray = sheetName.Split(Config[EnumIdentifier.SpitChar.ToString()]).ToList();
                            if (sheetSpitArray.Count <= 1 || sheetSpitArray.Count > 3)
                            {
                                Utils.LogInfo($"sheet名'{sheetName}' 格式有误，请检查!");
                                return;
                            }

                            var tbname = sheetSpitArray[0].ToLower();

                            //导入前清空，或跳过扫描
                            if (sheetSpitArray.Count == 3 &&
                                sheetSpitArray[2].Equals(Config[EnumIdentifier.Related.ToString()]))
                            {
                                cleartbNameList.Add(tbname);
                            }
                            else if (sheetSpitArray.Count == 3 &&
                                     sheetSpitArray[2].Equals(Config[EnumIdentifier.SkipScan.ToString()]))
                            {
                                Utils.LogInfo(($"--- '{sheetName}' 表跳过导入----"));
                                continue;
                            }


                            if (tbNameList.Any(m => m.Equals(tbname)))
                            {
                                Utils.LogInfo($"sheet名'{sheetName}'-关联表名'{tbname}'重复出现，请检查!");
                                return;
                            }

                            if (!DbTbDescDic.TryGetValue($"{dbName}.{tbname}", out var tbDescs))
                            {
                                Utils.LogInfo($"{tbname}对应数据库表不存在，请检查");
                                return;
                            }

   


                            SheetTbMapDic.Add($"{sheetName}", $"{dbName}.{tbname}");

                            //处理表头
                            var ret = xmlSheetReader.ResolveSheetHeader(sheetName, DbTbDescDic, tbname, tbDescs);
                            if (!ret.Item1)
                                return;// 表头有问题，直接结束

                            dic.Add($"{dbName}.{tbname}", ret.Item2);
                        }
                    }

                    if (cleartbNameList.Count > 0)
                        ClearTbDic.Add(dbName, cleartbNameList);

                    if (dic.Count >= 0)
                    {
                        ScanDescDic.Add(dbName, dic);
                    }
                }
                Utils.LogInfo(GetLineMsg("校验Excel表头 通过！", false), false);

                //Excel里面表都标记为 跳过
                if (ScanDescDic.Count <= 0)
                {
                    Utils.LogInfo($"Excel表sheet表都标记为跳过，无需导入，程序结束！");
                    return;
                }
                #endregion

                #region 数据库提前取出 关联数据 (唯一，关联)
                foreach (var eachDbScanDesc in ScanDescDic)//dbname:[dbname.tbname : list]
                {
                    var dbname = eachDbScanDesc.Key;
                    var retDic = ClearTbDic.TryGetValue(dbname, out var ClearTbNameList); // [dbname: list<string>]
                    var connectionString = Utils.Config.GetConnectionString(dbname);
                    using (IDbConnection con = new MySqlConnection(connectionString))
                    {
                        con.Open();
                        foreach (var tbScanDesc in eachDbScanDesc.Value)
                        {
                            var fullTbName = tbScanDesc.Key;
                            //当前sheet表 导入前被清空了，无需唯一性
                            if (!retDic || !ClearTbNameList.Contains(fullTbName.Split(Config[EnumIdentifier.Dot.ToString()]).ToList()[1]))
                            {
                                var uniqDescs = tbScanDesc.Value.Where(m => m.Prefix.Contains(Config[EnumIdentifier.Unique.ToString()]));
                                if (uniqDescs.Count() > 0)
                                {
                                    foreach (var uniqDesc in uniqDescs)
                                    {
                                        if (UniqDataDic.TryGetValue($"{fullTbName}.{uniqDesc.FieldName}", out var _ret))//多个表都关联了这个表，避免键值对重复
                                            continue;

                                        List<string> ret = con.GetUniqFieldValues(fullTbName, uniqDesc.FieldName, Config[(eachDbScanDesc.Key + "_JGIDName")], JGID);
                                        if (ret.Count > 0)
                                        {
                                            //唯一数据 key: db.tbname.uniqFieldName
                                            UniqDataDic.Add($"{fullTbName}.{uniqDesc.FieldName}", ret);
                                        }
                                    }
                                }
                            }

                            var rlDescs = tbScanDesc.Value.Where(m => m.Prefix.Contains(Config[EnumIdentifier.Related.ToString()])).ToList();
                            if (rlDescs.Count() > 0)
                            {
                                foreach (var rlDesc in rlDescs)
                                {
                                    if (RelatedDataDic.TryGetValue($"{rlDesc.RelatedFullTbName}.{rlDesc.KeyFieldNamesStr}.{rlDesc.ValueFieldName}", out var _ret))//多个表都关联了这个表，避免键值对重复
                                        continue;

                                    //当前sheet表关联表被清空，无需去取数据库关联数据
                                    if (!retDic || !ClearTbNameList.Contains(rlDesc.RelatedFullTbName.Split(Config[EnumIdentifier.Dot.ToString()]).ToList()[1]))
                                    {
                                        var singleTbName = rlDesc.RelatedFullTbName.Split(Config[EnumIdentifier.Dot.ToString()])[0];
                                        var ret = con.GetDicValues(rlDesc.RelatedFullTbName, rlDesc.KeyFileldNameList, rlDesc.ValueFieldName, Config[(singleTbName + "_JGIDName")], JGID);
                                        //关联基础数据数据 key:dbname.tbname-keyFieldname-valueFieldname
                                        RelatedDataDic.Add($"{rlDesc.RelatedFullTbName}.{rlDesc.KeyFieldNamesStr}.{rlDesc.ValueFieldName}", ret);
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion

                #region 校验行数据
                Utils.LogInfo(GetLineMsg("开始校验Excel行数据", true), false);
                //校验Excel字段数据 
                foreach (var file in Files)
                {
                    var dbStrArr = file.Name.Split('-');
                    var JGIDFieldName = Config[dbStrArr[0] + "_JGIDName"];
                    Utils.LogInfo($"Excel表: {file.Name} --------");
                    var sheetNames = MyMiniExcel.GetSheetNames(file.FullName);

                    using (FileStream stream = Helpers.OpenSharedRead(file.FullName))
                    {
                        var xmlSheetReader = new ExcelOpenXmlSheetReader(stream);

                        foreach (var sheetName in sheetNames)
                        {
                            if(!SheetTbMapDic.TryGetValue(sheetName, out var strDbTb))
                                continue;
                            //var strDbTb = SheetTbMapDic[sheetName];

                            string[] items = strDbTb.Split(Config[EnumIdentifier.Dot.ToString()]);

                            var ret = xmlSheetReader.GetSheetExecSql(sheetName,items[0], items[1], DbTbDescDic[strDbTb], ScanDescDic[items[0]][strDbTb], JGIDFieldName, JGID,
                                ref RemainErrorCount, ref UniqDataDic);
                            if (RemainErrorCount <= 0)
                                return;

                            ExecSqlFinalDic.Add(strDbTb, ret);
                        }
                    }
                }
                Utils.LogInfo(GetLineMsg("校验Excel行数据 通过！", false), false);

                Utils.LogInfo(GetLineMsg("开始校验关联数据", true), false);
                //校验Excel字段 关联数据   ExecSqlFinalDic 
                foreach (var scanDbDesc in ScanDescDic) //dbname : [dbname.tbname : list]
                {
                    foreach (var scanTbDesc in scanDbDesc.Value)//每个sheet表
                    {
                        var fullTbName = scanTbDesc.Key;
                        DbTbDescDic.TryGetValue(fullTbName, out var _curTbDescs);//数据表字段描述
                        var treeTbNames = Utils.Config["treeTbNames"];
                        var singleTbName = fullTbName.Split(Config[EnumIdentifier.Dot.ToString()]).ToList()[1];
                        var isTreeTb = treeTbNames.Contains(singleTbName);
                        var treeTbFields = Utils.Config["treeTbFields"];

                        var scanFieldDescs = scanTbDesc.Value.Where(m => m.Prefix.Contains(Config[EnumIdentifier.Related.ToString()])).ToList();
                        if (scanFieldDescs.Count <= 0)
                            continue;

                        Utils.LogInfo($"fullTbName: {fullTbName} --------");

                        var curScanRet = ExecSqlFinalDic[fullTbName];
                        var curParms = curScanRet.ExecParams;
                        var relatedTbNames = scanFieldDescs.Select(m => m.RelatedFullTbName).Distinct();//获取关联表集合
                        var noMatchInfos = new List<RltNoMatchInfo>();

                        var rScanRetKeyPairs = ExecSqlFinalDic.Where(m => relatedTbNames.Contains(m.Key)).ToList();//关联表 (可能会存在没有的情况)

                        var Comma = Config[EnumIdentifier.Comma.ToString()];
                        //多个字段确定关联关系的 比如调价表通过 找 批次表的ypbm以及批次号 确定yppcid 
                        var multRltScanDescs = scanFieldDescs.Where(m =>m.KeyFileldNameList.Count>1 && m.Prefix.Contains(Config[EnumIdentifier.Related.ToString()])).ToList();
                        var multRltDic = new Dictionary<string, Dictionary<string, string>>();
                        foreach (var curFieldDesc in multRltScanDescs)
                        {
                            //获取关联表 动态参数
                            var dParams = rScanRetKeyPairs.FirstOrDefault(m => m.Key == curFieldDesc.RelatedFullTbName);
                            if (dParams.Equals(default(KeyValuePair<string, ReadExcelDataRet>)))
                                continue;
                            var dicSingle = new Dictionary<string, string>();
                            var rltFields = curFieldDesc.KeyFileldNameList;//多个关联字段
                            foreach (var param in dParams.Value.ExecParams)
                            {
                                var keyStrList = new List<string>();
                                foreach (var fName in rltFields)
                                {
                                    keyStrList.Add(param.Get<string>(fName));
                                }

                                var joinKey = string.Join(Comma, keyStrList);
                                if (dicSingle.TryGetValue(joinKey, out var _value))
                                { 
                                    Utils.LogInfo($"关联key值：'{joinKey}' 重复!");
                                    continue;
                                }

                                dicSingle.Add(string.Join(Comma,keyStrList), param.Get<string>(curFieldDesc.ValueFieldName));
                            }

                            multRltDic.Add($"{curFieldDesc.RelatedFullTbName}.{curFieldDesc.KeyFieldNamesStr}",dicSingle);//to do
                        }

                        /*
                         判断关联数据思路
                         1- 数据库未找到关联数据
                         2- Excel 没有相关sheet表
                         3- Excel 有相关sheet表，但是没有相关数据(Excel没有该字段，有该字段但是没有找到数据)
                         4- 校验通过
                        */
                        for (int i = curParms.Count - 1; i >= 0; i--)    //倒序....
                        {
                            int rowPassItems = 0;
                            var rowIndex = curScanRet.RowIndexs[i];

                            foreach (var curScanFieldDesc in scanFieldDescs) //遍历 关联的字段
                            {
                                //关联值，
                                var relateValues = string.Empty;
                                var KeyFieldValue = curParms[i].Get<string>(curScanFieldDesc.FieldName);

                                //父级Pid 为空,直接跳过
                                if (isTreeTb && string.IsNullOrEmpty(KeyFieldValue) && treeTbFields.Contains(curScanFieldDesc.FieldName))
                                {
                                    rowPassItems++;
                                    continue;
                                }

                                //从数据库找关联数据
                                if (RelatedDataDic.TryGetValue($"{curScanFieldDesc.RelatedFullTbName}.{curScanFieldDesc.KeyFieldNamesStr}.{curScanFieldDesc.ValueFieldName}", out var _v1))
                                {
                                    if(!string.IsNullOrEmpty(KeyFieldValue) && _v1.TryGetValue(KeyFieldValue, out var _v2))
                                    {
                                        relateValues = _v2;
                                        if (curParms[i] != null)
                                        {
                                            curParms[i].Add(curScanFieldDesc.FieldName, relateValues); //覆盖为关联值
                                            rowPassItems++;
                                            continue;
                                        }
                                    }
                                }

                                //从导入的Excel找关联数据
                                if (string.IsNullOrEmpty(relateValues))
                                {
                                    var isNeeded = _curTbDescs
                                        .FirstOrDefault(m => m.FieldName == curScanFieldDesc.FieldName).IsNeeded;

                                    //是否有相关表  
                                    var rScanRetDesc = rScanRetKeyPairs.FirstOrDefault(m => m.Key == curScanFieldDesc.RelatedFullTbName);
                                    if (rScanRetDesc.Equals(default(KeyValuePair<string, ReadExcelDataRet>)))
                                    {
                                        noMatchInfos.Add(new RltNoMatchInfo(rowIndex, $"第'{rowIndex}'行数据:{KeyFieldValue?? "'空值'"} 通过关联'{curScanFieldDesc.RelatedFullTbName}'表的{curScanFieldDesc.KeyFieldNamesStr}字段,未找到'{curScanFieldDesc.ValueFieldName}'字段数据!"));

                                        if (isNeeded)
                                        {
                                            continue;
                                        }

                                        curParms[i].Add(curScanFieldDesc.FieldName);
                                        rowPassItems++;
                                        continue;
                                    }

                                    if (curScanFieldDesc.KeyFileldNameList.Count <= 1)
                                    {                                  
                                        var rParms = rScanRetDesc.Value.ExecParams;
                                        if (string.IsNullOrEmpty(rParms.FirstOrDefault()?.Get<string>(curScanFieldDesc.KeyFieldNamesStr)))
                                        {
                                            noMatchInfos.Add(new RltNoMatchInfo(rowIndex, $"第'{rowIndex}'行数据:{KeyFieldValue ?? "'空值'"} 通过关联'{curScanFieldDesc.RelatedFullTbName}'表的{curScanFieldDesc.KeyFieldNamesStr}字段,未找到'{curScanFieldDesc.ValueFieldName}'字段数据!"));
                                            if (isNeeded)
                                            {
                                                continue;
                                            }

                                            curParms[i].Add(curScanFieldDesc.FieldName);
                                            rowPassItems++;
                                            continue;
                                        }

                                        var rValueParms = rParms.FirstOrDefault(o => o.Get<string>(curScanFieldDesc.KeyFieldNamesStr).Equals(KeyFieldValue));
                                        if (rValueParms == null)
                                        {
                                            noMatchInfos.Add(new RltNoMatchInfo(rowIndex, $"第'{rowIndex}'行数据:{KeyFieldValue ?? "'空值'"} 通过关联'{curScanFieldDesc.RelatedFullTbName}'表的{curScanFieldDesc.KeyFieldNamesStr}字段,未找到'{curScanFieldDesc.ValueFieldName}'字段数据!"));
                                            if (isNeeded)
                                            {
                                                continue;
                                            }

                                            curParms[i].Add(curScanFieldDesc.FieldName);
                                            rowPassItems++;
                                            continue;
                                        }

                                        relateValues = rValueParms.Get<string>(curScanFieldDesc.ValueFieldName);
                                    }
                                    else
                                    {
                                        if (!multRltDic.TryGetValue(
                                            $"{curScanFieldDesc.RelatedFullTbName}.{curScanFieldDesc.KeyFieldNamesStr}", out var _v3))
                                        {
                                            noMatchInfos.Add(new RltNoMatchInfo(rowIndex, $"第'{rowIndex}'行数据:{KeyFieldValue ?? "'空值'"} 通过关联'{curScanFieldDesc.RelatedFullTbName}'表的{curScanFieldDesc.KeyFieldNamesStr}字段,未找到'{curScanFieldDesc.ValueFieldName}'字段数据!"));

                                            if (isNeeded)
                                            {
                                                continue;
                                            }

                                            curParms[i].Add(curScanFieldDesc.FieldName);
                                            rowPassItems++;
                                            continue;
                                        }

                                        if (!_v3.TryGetValue(KeyFieldValue, out var _v4))
                                        {
                                            noMatchInfos.Add(new RltNoMatchInfo(rowIndex, $"第'{rowIndex}'行数据:{KeyFieldValue ?? "'空值'"} 通过关联'{curScanFieldDesc.RelatedFullTbName}'表的{curScanFieldDesc.KeyFieldNamesStr}字段,未找到'{curScanFieldDesc.ValueFieldName}'字段数据!"));
                                            if (isNeeded)
                                            {
                                                continue;
                                            }

                                            curParms[i].Add(curScanFieldDesc.FieldName);
                                            rowPassItems++;
                                            continue;
                                        }

                                        relateValues = _v4;
                                    }

                                    //覆盖值  
                                    curParms[i].Add(curScanFieldDesc.FieldName, relateValues);
                                    rowPassItems++;
                                }
                            }
                            // 不通过 remain--，LegalRowCount，IllegalRowCount, -》删除该行参数
                            if (rowPassItems != scanFieldDescs.Count)
                            {
                                curScanRet.ExecParams.Remove(curParms[i]);
                                curScanRet.LegalRowCount--;
                                curScanRet.IllegalRowIndexSet.Add(rowIndex);
                            }
                        }

                        if (noMatchInfos.Any())
                        {
                            foreach (var infoItem in noMatchInfos.OrderBy(m=>m.Sort))
                            {
                                Utils.LogInfo(infoItem.Info);
                            }
                        }
                    }

                }
                Utils.LogInfo(GetLineMsg("校验关联数据 通过！", false), false);
                #endregion

                #region 校验结果 输出
                Utils.LogInfo(GetLineMsg("所有校验项已全部通过，正在统计校验结果....", true), false);
                //key: dbname.tbname
                foreach (var execItem in ExecSqlFinalDic)
                {
                    Utils.LogInfo($"{execItem.Key}表：");
                    Utils.LogInfo($"总条数：{execItem.Value.AllRowCount},通过条数：{execItem.Value.LegalRowCount}，不通过条数：{execItem.Value.IllegalRowIndexSet.Count}");
                    //这里调试的时候 可以打印其他的信息 to do
                }

                if (ClearTbDic.Count > 0)
                {
                    Utils.LogInfo("\r\n导入前清空的表:");
                    foreach (var item in ClearTbDic)
                    {
                        Utils.LogInfo($"----数据库名:{item.Key}----");
                        Utils.LogInfo($"表名集合:{string.Join(',', item.Value)}");
                    }
                    Utils.LogInfo("！！！点击开始导入按钮前请先检查确认这些表无误,否则导入后可能无法找回数据!");
                }

                #endregion

                HasPassFileValid = true;
            }
            catch (Exception ex)
            {
                DiagTip("校验异常 请检查");

                Utils.LogInfo(ex.Message+ex.StackTrace);
            }
        }

        /// <summary>
        /// 机构初始化
        /// </summary>
        private void OrgsInit()
        {
            Utils.LogInfo("正在读取机构信息....");
            var connectString = Config.GetConnectionString(Config["OrgDb"]);
            if (string.IsNullOrEmpty(connectString))
            {
                DiagTip("机构表所在库名有误，请检查");
                return;
            }

            List<HospitalInfo> infoList;
            using (IDbConnection con = new MySqlConnection(connectString))
            {
                con.Open();
                try
                {
                    infoList = con.Query<HospitalInfo>($"select {Config["OrgIDFieldName"]} as Id,{Config["OrgMcFieldName"]} as Mc from { Config["OrgTb"]} where isdeleted =0 and jybs=0").ToList();
                }
                catch (Exception ex)
                {
                    Utils.LogInfo("读取机构出错" + ex.Message);
                    DiagTip("读取机构出错，请检查！");
                    return;
                }
            }
            if (infoList.Count <= 0)
                DiagTip("数据库中不存在机构！请预先加入！");

            DrpDwnOrgs.DataSource = infoList;
            DrpDwnOrgs.ValueMember = "Id";
            DrpDwnOrgs.DisplayMember = "Mc";
            Utils.LogInfo("读取机构信息 成功！");
            //DiagTip("导入成功！", MessageBoxIcon.Information);
        }

        /// <summary>
        /// 弹窗
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="iconType"></param>
        /// <param name="caption"></param>
        private void DiagTip(string msg, MessageBoxIcon iconType = MessageBoxIcon.Warning, string caption="提示")
        {
            MessageBox.Show(msg, caption, MessageBoxButtons.OK, iconType);
        }

        /// <summary>
        /// 格式化 每段 输出日志
        /// </summary>
        /// <param name="msg"></param>
        /// <returns></returns>
        private string GetLineMsg(string msg,bool IsStart)
        {
            return IsStart ? "\r\n----------"+ DateTime.Now +"--" + msg + "----------" : "  -----" + DateTime.Now + "--" + msg + "-----";
        }

        /// <summary>
        /// 重载配置文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnReloadJson_Click(object sender, EventArgs e)
        {
            Utils.Reload();

            DrpDwnOrgs.DataSource = null;
            DrpDwnOrgs.Text = "请选择机构";

            Config = Utils.Config;
            Seconds = Utils.Config["SencondLimits"].ToInt32();
            OrgsInit();

            DiagTip("重载成功!");
        }

        /// <summary>
        /// 基础验证
        /// </summary>
        /// <param name="dirPath"></param>
        /// <param name="txtMaxCount"></param>
        /// <param name="files"></param>
        /// <param name="maxCount"></param>
        /// <returns></returns>
        public bool ValidateBase(string dirPath, string txtMaxCount, ref List<FileInfo> files, out int maxCount)
        {
            Utils.LogInfo("\r\n\r\n----------------------------" + DateTime.Now + " ----------------------------");
            Utils.LogInfo("  -----开始基础校验! -----");
            maxCount = 0;
            if (string.IsNullOrEmpty(txtMaxCount))
            {
                Utils.LogInfo("  请输入错误记录阈值!");
                DiagTip("请输入错误记录阈值");
                return false;
            }
            if (!TryParse(txtMaxCount, out maxCount))
            {
                Utils.LogInfo("错误记录阈值栏位应为正整数");
                DiagTip("错误记录阈值栏位应为正整数");
                return false;
            }
            if (maxCount <= 0)
            {
                Utils.LogInfo("错误记录阈值栏位应为正整数");
                DiagTip("错误记录阈值栏位应为正整数");
                return false;
            }

            if (string.IsNullOrWhiteSpace(dirPath))
            {
                Utils.LogInfo("请输入文件夹路径!");
                DiagTip("请输入文件夹路径!");
                return false;
            }
            if (!Directory.Exists(dirPath))
            {
                Utils.LogInfo("文件夹不存在，路径错误!");
                DiagTip("文件夹不存在，路径错误!");

                return false;
            }

            DirectoryInfo dir = new DirectoryInfo(dirPath);
            files = dir.GetFiles().ToList();
            if (!files.Any())
            {
                Utils.LogInfo("文件夹内无文件,请确认!");
                DiagTip("文件夹内无文件,请确认!");
                return false;
            }

            files = files.Where(m => !m.Name.Contains("$")).ToList();
            Utils.LogInfo("  -----基础校验通过! -----");

            return true;
        }
    }
}
