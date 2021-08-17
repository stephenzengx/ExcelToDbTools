using Excel.Utils;
using Excel.Zip;
using ExcelTools;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using Dapper;
using ExcelTools.SqlScheme;

namespace Excel.OpenXml
{
    internal class ExcelOpenXmlSheetReader : IExcelReader, IExcelReaderAsync
    {
        private const string _ns = OpenXmlConfig.SpreadsheetmlXmlns;
        private List<SheetRecord> _sheetRecords;
        private List<string> _sharedStrings;
        private MergeCells _mergeCells;
        private ExcelOpenXmlStyles _style;
        private ExcelOpenXmlZip _archive;
        private static readonly XmlReaderSettings _xmlSettings = new XmlReaderSettings
        {
            IgnoreComments = true,
            IgnoreWhitespace = true,
            XmlResolver = null,
        };

        public ExcelOpenXmlSheetReader(Stream stream)
        {
            _archive = new ExcelOpenXmlZip(stream);
        }

        public IEnumerable<IDictionary<string, object>> Query(bool useHeaderRow, string sheetName, string startCell, IConfiguration configuration)
        {
            var config = (OpenXmlConfiguration)configuration ?? OpenXmlConfiguration.DefaultConfig; //TODO:
            if (!ReferenceHelper.ParseReference(startCell, out var startColumnIndex, out var startRowIndex))
                throw new InvalidDataException($"startCell {startCell} is Invalid");
            startColumnIndex--; startRowIndex--;

            //TODO:need to optimize
            SetSharedStrings();

            // if sheets count > 1 need to read xl/_rels/workbook.xml.rels  
            var sheets = _archive.entries.Where(w => w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                || w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
            );
            ZipArchiveEntry sheetEntry = null;
            if (sheetName != null)
            {
                SetWorkbookRels(_archive.entries);
                var s = _sheetRecords.SingleOrDefault(_ => _.Name == sheetName);
                if (s == null)
                    throw new InvalidOperationException("Please check sheetName/Index is correct");
                sheetEntry = sheets.Single(w => w.FullName == $"xl/{s.Path}" || w.FullName == $"/xl/{s.Path}" || w.FullName == s.Path || s.Path == $"/{w.FullName}");
            }
            else if (sheets.Count() > 1)
            {
                SetWorkbookRels(_archive.entries);
                var s = _sheetRecords[0];
                sheetEntry = sheets.Single(w => w.FullName == $"xl/{s.Path}" || w.FullName == $"/xl/{s.Path}");
            }
            else
                sheetEntry = sheets.Single();

            #region MergeCells
            if (config.FillMergedCells)
            {
                _mergeCells = new MergeCells();
                using (var sheetStream = sheetEntry.Open())
                using (XmlReader reader = XmlReader.Create(sheetStream, _xmlSettings))
                {
                    if (!reader.IsStartElement("worksheet", _ns))
                        yield break;
                    while (reader.Read())
                    {
                        if (reader.IsStartElement("mergeCells", _ns))
                        {
                            if (!XmlReaderHelper.ReadFirstContent(reader))
                                yield break;
                            while (!reader.EOF)
                            {
                                if (reader.IsStartElement("mergeCell", _ns))
                                {
                                    var @ref = reader.GetAttribute("ref");
                                    var refs = @ref.Split(':');
                                    if (refs.Length == 1)
                                        continue;

                                    ReferenceHelper.ParseReference(refs[0], out var x1, out var y1);
                                    ReferenceHelper.ParseReference(refs[1], out var x2, out var y2);

                                    _mergeCells.MergesValues.Add(refs[0], null);

                                    // foreach range
                                    var isFirst = true;
                                    for (int x = x1; x <= x2; x++)
                                    {
                                        for (int y = y1; y <= y2; y++)
                                        {
                                            if (!isFirst)
                                                _mergeCells.MergesMap.Add(ReferenceHelper.ConvertXyToCell(x, y), refs[0]);
                                            isFirst = false;
                                        }
                                    }

                                    XmlReaderHelper.SkipContent(reader);
                                }
                                else if (!XmlReaderHelper.SkipContent(reader))
                                {
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            #endregion

            // TODO: need to optimize performance
            var withoutCR = false;
            var maxRowIndex = -1;
            var maxColumnIndex = -1;

            //Q. why need 3 times openstream merge one open read? A. no, zipstream can't use position = 0
            using (var sheetStream = sheetEntry.Open())
            using (XmlReader reader = XmlReader.Create(sheetStream, _xmlSettings))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement("c", _ns))
                    {
                        var r = reader.GetAttribute("r");
                        if (r != null)
                        {
                            if (ReferenceHelper.ParseReference(r, out var column, out var row))
                            {
                                column = column - 1;
                                row = row - 1;
                                maxRowIndex = Math.Max(maxRowIndex, row);
                                maxColumnIndex = Math.Max(maxColumnIndex, column);
                            }
                        }
                        else
                        {
                            withoutCR = true;
                            break;
                        }
                    }
                    //this method logic depends on dimension to get maxcolumnIndex, if without dimension then it need to foreach all rows first time to get maxColumn and maxRowColumn
                    else if (reader.IsStartElement("dimension", _ns))
                    {
                        var @ref = reader.GetAttribute("ref");
                        if (string.IsNullOrEmpty(@ref))
                            throw new InvalidOperationException("Without sheet dimension data");
                        var rs = @ref.Split(':');
                        // issue : https://github.com/shps951023/MiniExcel/issues/102
                        if (ReferenceHelper.ParseReference(rs.Length == 2 ? rs[1] : rs[0], out int cIndex, out int rIndex))
                        {
                            maxColumnIndex = cIndex - 1;
                            maxRowIndex = rIndex - 1;
                            break;
                        }
                        else
                            throw new InvalidOperationException("Invaild sheet dimension start data");
                    }
                }
            }

            if (withoutCR)
            {
                using (var sheetStream = sheetEntry.Open())
                using (XmlReader reader = XmlReader.Create(sheetStream, _xmlSettings))
                {
                    if (!reader.IsStartElement("worksheet", _ns))
                        yield break;
                    if (!XmlReaderHelper.ReadFirstContent(reader))
                        yield break;
                    while (!reader.EOF)
                    {
                        if (reader.IsStartElement("sheetData", _ns))
                        {
                            if (!XmlReaderHelper.ReadFirstContent(reader))
                                continue;

                            while (!reader.EOF)
                            {
                                if (reader.IsStartElement("row", _ns))
                                {
                                    maxRowIndex++;

                                    if (!XmlReaderHelper.ReadFirstContent(reader))
                                        continue;

                                    //Cells
                                    {
                                        var cellIndex = -1;
                                        while (!reader.EOF)
                                        {
                                            if (reader.IsStartElement("c", _ns))
                                            {
                                                cellIndex++;
                                                maxColumnIndex = Math.Max(maxColumnIndex, cellIndex);
                                            }


                                            if (!XmlReaderHelper.SkipContent(reader))
                                                break;
                                        }
                                    }
                                }
                                else if (!XmlReaderHelper.SkipContent(reader))
                                {
                                    break;
                                }
                            }
                        }
                        else if (!XmlReaderHelper.SkipContent(reader))
                        {
                            break;
                        }
                    }

                }
            }

            using (var sheetStream = sheetEntry.Open())
            using (XmlReader reader = XmlReader.Create(sheetStream, _xmlSettings))
            {
                if (!reader.IsStartElement("worksheet", _ns))
                    yield break;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    yield break;

                while (!reader.EOF)
                {
                    if (reader.IsStartElement("sheetData", _ns))
                    {
                        if (!XmlReaderHelper.ReadFirstContent(reader))
                            continue;

                        Dictionary<int, string> headRows = new Dictionary<int, string>();
                        int rowIndex = -1;
                        int nextRowIndex = 0;
                        bool isFirstRow = true;
                        while (!reader.EOF)
                        {
                            if (reader.IsStartElement("row", _ns))
                            {
                                nextRowIndex = rowIndex + 1;
                                if (int.TryParse(reader.GetAttribute("r"), out int arValue))
                                    rowIndex = arValue - 1; // The row attribute is 1-based
                                else
                                    rowIndex++;

                                // row -> c
                                if (!XmlReaderHelper.ReadFirstContent(reader))
                                    continue;

                                // startcell pass rows
                                if (rowIndex < startRowIndex)
                                {
                                    XmlReaderHelper.SkipToNextSameLevelDom(reader);
                                    continue;
                                }


                                // fill empty rows
                                if (!(nextRowIndex < startRowIndex))
                                {
                                    if (nextRowIndex < rowIndex)
                                    {
                                        for (int i = nextRowIndex; i < rowIndex; i++)
                                        {
                                            yield return GetCell(useHeaderRow, maxColumnIndex, headRows, startColumnIndex);
                                        }
                                    }
                                }

                                // Set Cells
                                {
                                    var cell = GetCell(useHeaderRow, maxColumnIndex, headRows, startColumnIndex);
                                    var columnIndex = withoutCR ? -1 : 0;
                                    while (!reader.EOF)
                                    {
                                        if (reader.IsStartElement("c", _ns))
                                        {
                                            var aS = reader.GetAttribute("s");
                                            var aR = reader.GetAttribute("r");
                                            var aT = reader.GetAttribute("t");
                                            var cellValue = ReadCellAndSetColumnIndex(reader, ref columnIndex, withoutCR, startColumnIndex, aR, aT);

                                            if (config.FillMergedCells)
                                            {
                                                if (_mergeCells.MergesValues.ContainsKey(aR))
                                                {
                                                    _mergeCells.MergesValues[aR] = cellValue;
                                                }
                                                else if (_mergeCells.MergesMap.ContainsKey(aR))
                                                {
                                                    var mergeKey = _mergeCells.MergesMap[aR];
                                                    object mergeValue = null;
                                                    if(_mergeCells.MergesValues.ContainsKey(mergeKey))
                                                        mergeValue = _mergeCells.MergesValues[mergeKey];
                                                    cellValue = mergeValue;
                                                }
                                            }

                                            if (columnIndex < startColumnIndex)
                                                continue;

                                            if (!string.IsNullOrEmpty(aS)) // if c with s meaning is custom style need to check type by xl/style.xml
                                            {
                                                int xfIndex = -1;
                                                if (int.TryParse(aS, NumberStyles.Any, CultureInfo.InvariantCulture, out var styleIndex))
                                                    xfIndex = styleIndex;

                                                // only when have s attribute then load styles xml data
                                                if (_style == null)
                                                    _style = new ExcelOpenXmlStyles(_archive);
      
                                                cellValue = _style.ConvertValueByStyleFormat(xfIndex, cellValue);
                                                SetCellsValueAndHeaders(cellValue, useHeaderRow, ref headRows, ref isFirstRow, ref cell, columnIndex);
                                            }
                                            else
                                            {
                                                SetCellsValueAndHeaders(cellValue, useHeaderRow, ref headRows, ref isFirstRow, ref cell, columnIndex);
                                            }
                                        }
                                        else if (!XmlReaderHelper.SkipContent(reader))
                                            break;
                                    }

                                    if (isFirstRow)
                                    {
                                        isFirstRow = false; // for startcell logic
                                        if (useHeaderRow)
                                            continue;
                                    }


                                    yield return cell;
                                }
                            }
                            else if (!XmlReaderHelper.SkipContent(reader))
                            {
                                break;
                            }
                        }

                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        break;
                    }
                }
            }
        }

        private static IDictionary<string, object> GetCell(bool useHeaderRow, int maxColumnIndex, Dictionary<int, string> headRows, int startColumnIndex)
        {
            return useHeaderRow ? CustomPropertyHelper.GetEmptyExpandoObject(headRows) : CustomPropertyHelper.GetEmptyExpandoObject(maxColumnIndex, startColumnIndex);
        }

        private void SetCellsValueAndHeaders(object cellValue, bool useHeaderRow, ref Dictionary<int, string> headRows, ref bool isFirstRow, ref IDictionary<string, object> cell, int columnIndex)
        {
            if (useHeaderRow)
            {
                if (isFirstRow) // for startcell logic
                {
                    var cellValueString = cellValue?.ToString();
                    if (!string.IsNullOrWhiteSpace(cellValueString))
                        headRows.Add(columnIndex, cellValueString);
                }
                else
                {
                    if (headRows.ContainsKey(columnIndex))
                    {
                        var key = headRows[columnIndex];
                        cell[key] = cellValue;
                    }
                }
            }
            else
            {
                //if not using First Head then using A,B,C as index
                cell[ColumnHelper.GetAlphabetColumnName(columnIndex)] = cellValue;
            }
        }

        #region 结合 Dapper，扩展了一些方法 by -zx
        /// <summary>
        /// 处理sheet表头
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="tbDbDescdic"></param>
        /// <param name="tbName"></param>
        /// <param name="curTbDesc"></param>
        /// <param name="startCell"></param>
        /// <param name="configuration"></param>
        /// <returns></returns>
        public Tuple<bool, List<ScanExcelHeadDesc>> ResolveSheetHeader(
            string sheetName, 
            Dictionary<string, List<TbDesc>> tbDbDescdic,
            string tbName, List<TbDesc> curTbDesc,
            string startCell = "A1", IConfiguration configuration=null)
        {
            string PrefixAll = ExcelTools.Utils.Config["PrefixAll"];
            var config = ExcelTools.Utils.Config;
            List<ScanExcelHeadDesc> list = new List<ScanExcelHeadDesc>();
            var ret = new Tuple<bool, List<ScanExcelHeadDesc>>(false,list);

            //获取表头
            var headers = Query(false, sheetName, startCell, configuration).FirstOrDefault()?.Values
                                    ?.Select(s => s?.ToString())?.ToArray();
            if (headers == null || headers.Length <= 0)
            {
                ExcelTools.Utils.LogInfo("表头为空，请检查");
                return ret;
            }

            if (headers.Any(m => string.IsNullOrEmpty(m)))
            {
                ExcelTools.Utils.LogInfo("空缺某些表头栏位，请检查！");
                return ret;
            }

            //全部转成小写
            for (var i = 0; i < headers.Length; i++)
                headers[i] = headers[i]?.ToLower();

            foreach (var head in headers)
            {
                //无特殊含义字段
                if (!head.Contains(config[EnumIdentifier.SpitChar.ToString()]))
                {
                    if (curTbDesc.FirstOrDefault(m => m.FieldName == head) == null)
                    {
                        ExcelTools.Utils.LogInfo($"数据库表'{tbName}'不存在表字段：{head}");
                        return ret;
                    }

                    if (list.Any(m => m.FieldName == head))
                    {
                        ExcelTools.Utils.LogInfo($"表头关联字段'{head}'重复，请检查");
                        return ret;
                    }

                    list.Add(new ScanExcelHeadDesc
                    {
                        HeaderName = head.ToLower(),
                        FieldName = head.ToLower()
                    });

                    continue;
                }
                //拆分表头
                List<string> spitItems = head.Split(config[EnumIdentifier.SpitChar.ToString()]).ToList();
                if (spitItems.Count<=1)
                {
                    ExcelTools.Utils.LogInfo($"表头{head}格式有误,请检查");
                    return ret;
                }
                if (!PrefixAll.Contains(spitItems[0])) //检查前缀格式
                {
                    ExcelTools.Utils.LogInfo($"表头{head} 前缀格式有误,请检查");
                    return ret;
                }

                /*                               
                BM：无前缀为基础字段，无处理
                !-BM: 代表该字段在本表唯一:           
                $-MC: (PYM-WBM 这两个字段是不是在每个表都名字一样) 拼音码/五笔码表头格式, !$唯一，然后拼音码，五笔码             
                *-Phone 手机号加密
                ^-KC-[0,100] 代表数据范围，除非非常严格，一般来说判断非负数，否则可以不加
                #-tb_relative.MC-RID*RID-RBM*RBM: 关联字段，通过tb_relative的MC，获取RID字段对应本表RID字段,获取RBM字段对应本表RBM字段               
                */

                //唯一，拼音/五笔, 加密
                var Prefix = spitItems[0];
                if (Prefix.Contains(config[EnumIdentifier.Unique.ToString()])
                    || Prefix.Contains(config[EnumIdentifier.PYWB.ToString()])
                    || Prefix.StartsWith(config[EnumIdentifier.Encry.ToString()]))
                {
                    if (spitItems.Count != 2)
                    {
                        ExcelTools.Utils.LogInfo($"表头{head}格式有误,请检查");
                        return ret;
                    }

                    if (curTbDesc.FirstOrDefault(m => m.FieldName == spitItems[1]) == null)
                    {
                        ExcelTools.Utils.LogInfo($"数据库表'{tbName}'不存在表字段：{spitItems[1]}");
                        return ret;
                    }
                    if (list.Any(m => m.FieldName == spitItems[1]))
                    {
                        ExcelTools.Utils.LogInfo($"表头关联字段'{head}'重复，请检查");
                        return ret;
                    }

                    list.Add(new ScanExcelHeadDesc
                    {
                        Prefix = Prefix,
                        FieldName = spitItems[1].ToLower(),
                        HeaderName = head.ToLower(),
                    });

                }
                //范围 ^-KC-[0,100]
                else if (Prefix.StartsWith(config[EnumIdentifier.Range.ToString()]))
                {
                    if (spitItems.Count != 3)
                    {
                        ExcelTools.Utils.LogInfo($"表头{head}格式有误,请检查");
                        return ret;
                    }

                    var tbDesc = curTbDesc.FirstOrDefault(m => m.FieldName == spitItems[1]);
                    if (tbDesc == null)
                    {
                        ExcelTools.Utils.LogInfo($"数据库表'{tbName}'不存在表字段：{spitItems[1]}");
                        return ret;
                    }
                    if (list.Any(m => m.FieldName == spitItems[1]))
                    {
                        ExcelTools.Utils.LogInfo($"表头关联字段'{head}'重复，请检查");
                        return ret;
                    }

                    var scanDesc = new ScanExcelHeadDesc
                    {
                        Prefix=Prefix, HeaderName = head.ToLower(), FieldName = spitItems[1].ToLower()
                    };

                    list.Add(scanDesc);
                    var rangItems = spitItems[2].Split(config[EnumIdentifier.Comma.ToString()]).ToList();
                    if (rangItems.Count != 2)
                    {
                        ExcelTools.Utils.LogInfo($"表头{head}范围格式有误,请检查");
                        return ret;
                    }
                    if (!rangItems[0].Contains(config[EnumIdentifier.LeftZkh.ToString()])
                        || !rangItems[1].Contains(config[EnumIdentifier.RightZkh.ToString()]))
                    {
                        ExcelTools.Utils.LogInfo($"表头'{head}'范围格式有误,请检查");
                        return ret;
                    }
                    //[0,]  | [0,22]
                    if (tbDesc.type == typeof(int))
                    {
                        var leftValue = 0;
                        var rightValue = 0;
                        if (!Int32.TryParse(rangItems[0].Substring(1), out leftValue))
                        {
                            ExcelTools.Utils.LogInfo($"表头'{head}'范围格式有误/范围左值有误,请检查");
                            return ret;
                        }

                        if (rangItems[1].Equals(config[EnumIdentifier.RightZkh.ToString()]))
                        {
                            rightValue = int.MaxValue; 
                        }
                        else
                        {
                            if (!Int32.TryParse(rangItems[1].Substring(0, rangItems[1].Length - 1), out rightValue))
                            {
                                ExcelTools.Utils.LogInfo($"'表头{head}'范围格式有误/范围右值有误,请检查");
                                return ret;
                            }
                        }
                        scanDesc.RangeInt = new Tuple<int, int>(leftValue, rightValue);
                    }
                    else if (tbDesc.type == typeof(decimal))
                    {
                        var leftValue = new decimal(0);
                        var rightValue = new decimal(0);
                        if (!decimal.TryParse(rangItems[0].Substring(1), out leftValue))
                        {
                            ExcelTools.Utils.LogInfo($"表头'{head}'范围格式有误/范围左值有误,请检查");
                            return ret;
                        }

                        if (rangItems[1].Equals(config[EnumIdentifier.RightZkh.ToString()]))
                        {
                            rightValue = decimal.MaxValue; 
                        }
                        else
                        {
                            if (!decimal.TryParse(rangItems[1].Substring(0, rangItems[1].Length - 1), out rightValue))
                            {
                                ExcelTools.Utils.LogInfo($"表头'{head}'范围格式有误/右值有误,请检查");
                                return ret;
                            }
                        }
                        scanDesc.RangeDecimal = new Tuple<decimal, decimal>(leftValue, rightValue);
                    }
                }
                //关联字符 #-db.tb_relative-MC-RID*RID-RBM*RBM
                else if (Prefix.StartsWith(config[EnumIdentifier.Related.ToString()]))
                {
                    if (spitItems.Count<=2)
                    {
                        ExcelTools.Utils.LogInfo($"表头'{head}'范围有误,请检查");
                        return ret;
                    }

                    var rtbName = spitItems[1];
                    if (!tbDbDescdic.TryGetValue(spitItems[1], out var rTbDesc))
                    {
                        ExcelTools.Utils.LogInfo($"表头 '{head}' 关联数据库表 '{rtbName}' 在数据库不存在,请检查");
                        return ret;
                    }

                    var keyFieldName = spitItems[2];
                    if (!rTbDesc.Any(m => m.FieldName.Equals(keyFieldName)))
                    {
                        ExcelTools.Utils.LogInfo($"表头关联表'{rtbName}'在不存在字段{keyFieldName},请检查");
                        return ret;
                    }

                    var rSpitItems = spitItems.GetRange(3,spitItems.Count-3);
                    foreach (var item in rSpitItems)
                    {
                        var rFields = item.Split(config[EnumIdentifier.Encry.ToString()]).ToList();
                        if(rFields.Count!=2)
                        {
                            ExcelTools.Utils.LogInfo($"关联表表头'{rtbName}'关联字段内容'{item}'格式有误,请检查");
                            return ret;
                        }

                        if (!rTbDesc.Any(m => m.FieldName.Equals(rFields[0])))
                        {
                            ExcelTools.Utils.LogInfo($"关联表表头'{rtbName}'在不存在字段{rFields[0]},请检查");
                            return ret;
                        }

                        if (!curTbDesc.Any(m => m.FieldName.Equals(rFields[1])))
                        {
                            ExcelTools.Utils.LogInfo($"主表'{tbName}'在不存在字段{rFields[1]},请检查");
                            return ret;
                        }

                        list.Add(new ScanExcelHeadDesc
                        {
                            HeaderName = head.ToLower(),
                            Prefix = Prefix,
                            FieldName = rFields[1].ToLower(),
                            RelatedTbName = rtbName.ToLower(),
                            KeyFieldName = keyFieldName.ToLower(),
                            ValueFieldName = rFields[0] //这个不能转小写，是值
                        });
                    }
                }
                else
                {
                    ExcelTools.Utils.LogInfo($"表头'{head}'前缀格式有误,请检查");
                    return ret;
                }
            }

            //遍历数据库中的必填字段，再Excel表头是否都出现过 <id,isdelete,cjsj,pym,wbm,,jgid,jguuid 代码生成，跳过验证>
            var strIgnores = ExcelTools.Utils.Config["IgnoreNeedValidFields"];
            foreach (var desc in curTbDesc.Where(m=>m.IsNeeded && !strIgnores.Contains(m.FieldName)))
            {
                //这已经验证过，表头一定在数据库存在,所以不用判空.
                var entity = list.FirstOrDefault(m => m.FieldName == desc.FieldName);//<通过其他表取的字段 跳过验证>
                if (entity == null)
                {
                    continue;
                }

                if (entity.Prefix.StartsWith(config[EnumIdentifier.Related.ToString()]))
                    continue;

                if (list.FirstOrDefault(m => m.FieldName == desc.FieldName) == null)
                {
                    ExcelTools.Utils.LogInfo($"表'{tbName}'字段 '{desc.FieldName}' 在数据库为必填字段，在Excel未填！");
                    return ret;
                }
                //其他待补充... to do
            }

            var retFinal = new Tuple<bool, List<ScanExcelHeadDesc>>(true, ret.Item2);
            return retFinal;
        }

        /// <summary>
        /// 读取表格数据拼接 Sql
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="dbName"></param>
        /// <param name="tbName"></param>
        /// <param name="tbDescs"></param>
        /// <param name="scanDescs"></param>
        /// <param name="JGID"></param>
        /// <param name="remainErrorCount"></param>
        /// <param name="uniqDataDic">唯一数据 key:db.tbname.uniqFieldName</param>
        /// <param name="relatedDataDic">key:dbname.tbname-keyFieldname-valueFieldname</param>
        /// <param name="startCell"></param>
        /// <param name="configuration"></param>
        /// <returns></returns>
        public ReadExcelDataRet GetSheetExecSql(string sheetName,string dbName, string tbName, 
            List<TbDesc> tbDescs, List<ScanExcelHeadDesc> scanDescs, string JGIDFieldName,
            string JGID, ref int remainErrorCount,
            ref Dictionary<string, List<string>> uniqDataDic,
            string startCell = "A1", IConfiguration configuration = null)
        {
            ExcelTools.Utils.LogInfo($"开始校验sheet表-行数据: {sheetName} --------");

            //思路：获取每行数据，将Excel每个字段(验证格式是否正确) 转成指定类型的值，然后拼接sql
            //拼接公共表字段

            var rowIndex = 2;//从第二行开始读
            bool isExistPYWB = scanDescs.Exists(m => m.Prefix.Contains(ExcelTools.Utils.Config[EnumIdentifier.PYWB.ToString()]));
            bool isExistJybs= tbDescs.Exists(m => m.FieldName.Equals("jybs"));
            ReadExcelDataRet ret = new ReadExcelDataRet
            {
                ExecParams = new List<DynamicParameters>(),
                RowIndexs = new List<int>()
            };

            var queryRet = Query(true, sheetName, startCell, configuration);
            //遍历每行数据
            foreach (var row in queryRet)
            {
                var rowItemPassCount = 0;
                Dictionary<string, object> rowRetDic = new Dictionary<string, object>();
                //遍历每个字段
                foreach (var rowItem in row)
                {
                    for (var i = 0; i < scanDescs.Count; i++)
                    {
                        if (!scanDescs[i].HeaderName.Equals(rowItem.Key.ToLower()))
                            continue;

                        object convertValue = null;
                        object itemValue = rowItem.Value;

                        if (!MyTypeMappingImpl(tbDescs.FirstOrDefault(m => m.FieldName == scanDescs[i].FieldName), scanDescs[i], ref convertValue, itemValue, rowIndex))
                            continue;

                        if (string.IsNullOrEmpty(scanDescs[i].Prefix))
                        {
                            rowRetDic.Add(scanDescs[i].FieldName, convertValue);
                        }
                        else if (scanDescs[i].Prefix.Contains(ExcelTools.Utils.Config[EnumIdentifier.Related.ToString()])) //关联字段
                        {
                            //scanDescs[i].KeyFieldValue = convertValue?.ToString(); //关于找关联字段值
                            rowRetDic.Add(scanDescs[i].FieldName, convertValue); //用于定位 paramlist 哪一行然后更新值为关联的值
                        }
                        else if (scanDescs[i].Prefix.Contains(ExcelTools.Utils.Config[EnumIdentifier.PYWB.ToString()])) //拼音,五笔
                        {
                            rowRetDic.Add(scanDescs[i].FieldName, convertValue);
                            rowRetDic.Add("pym", convertValue?.ToString().GetFirstPY());
                            rowRetDic.Add("wbm", convertValue?.ToString().GetFirstWB());
                        }
                        else if (scanDescs[i].Prefix.Contains(ExcelTools.Utils.Config[EnumIdentifier.Unique.ToString()])) //唯一
                        {
                            var existFlag = false;
                            //uniqDataDic key:db.tbname.uniqFieldName
                            if (    uniqDataDic.TryGetValue($"{dbName}.{tbName}.{scanDescs[i].FieldName}", out var list))
                            {
                                existFlag = true;
                                if (convertValue!=null && list.Contains(convertValue.ToString()))
                                {
                                    ExcelTools.Utils.LogInfo($"第{rowIndex}行，'{scanDescs[i].HeaderName}'列值 已存在!");
                                    continue;
                                }
                            }

                            if (!existFlag)
                            {
                                list = new List<string>();
                                uniqDataDic.Add($"{dbName}.{tbName}.{scanDescs[i].FieldName}", list);
                            }

                            //scanDescs[i].KeyFieldValue = convertValue;
                            list.Add(convertValue?.ToString());
                            //重置字典
                            uniqDataDic[$"{dbName}.{tbName}.{scanDescs[i].FieldName}"] = list;

                            rowRetDic.Add(scanDescs[i].FieldName, convertValue);
                        }
                        else if (scanDescs[i].Prefix.Contains(ExcelTools.Utils.Config[EnumIdentifier.Encry.ToString()])) // 手机号加密
                        {
                            rowRetDic.Add(scanDescs[i].FieldName, convertValue.ToString().ToPhoneCardNoEncryption());
                        }

                        rowItemPassCount++;
                    }
                }

                ret.AllRowCount++;
                if (rowItemPassCount== scanDescs.Count)
                {
                    ret.RowIndexs.Add(rowIndex++);
                    DynamicParameters parm = DbConnectionExtensions.GetBaseDynamicParameters(JGIDFieldName,JGID,isExistJybs);
                    foreach (var rowDic in rowRetDic)
                    {
                        parm.Add(rowDic.Key,rowDic.Value);
                    }
                    ret.ExecParams.Add(parm);
                    ret.LegalRowCount++;
                }
                else
                {
                    ret.RowIndexs.Add(rowIndex++);
                    ret.IllegalRowCount++;
                    remainErrorCount--;
                }

                if (remainErrorCount <= 0)
                {
                    ExcelTools.Utils.LogInfo($"已达到最大错误记录数阈值，程序结束！");
                    return null;
                }
            }

            ret.DbName = dbName;
            ret.TbName = tbName;
            ret.ExecSql = DbConnectionExtensions.JoinInsertHeadSql(tbName, scanDescs.Select(m => m.FieldName).ToList(), isExistPYWB, isExistJybs, JGIDFieldName, out var v1 );
            ret.FieldNameList = v1;

            //验证（数据值，以及跟Excel的行对应关系是正确的）
            //if (tbName == "t_base_empdept")
            //{
            //    for (var i=0; i< ret.ExecParams.Count;i++)
            //    {
            //        Console.WriteLine($"{ret.RowIndexs[i]} : ksuuid-{ret.ExecParams[i].Get<string>("ksuuid")}");
            //    }
            //}

            return ret;
        }

        private static bool MyTypeMappingImpl(TbDesc tbdesc, ScanExcelHeadDesc scanDesc, ref object newValue, object itemValue,int rowIndex)
        {
            if (itemValue == null && !tbdesc.IsNeeded)
            {
                return true;
            }

            if ((itemValue == null || string.IsNullOrEmpty(itemValue.ToString())) && 
                (tbdesc.IsNeeded && !scanDesc.Prefix.Contains(ExcelTools.Utils.Config[EnumIdentifier.Related.ToString()])))
            {
                ExcelTools.Utils.LogInfo($"第{rowIndex}行，表头{scanDesc.HeaderName} 为必填项！");
                return false;
            }

            var type = tbdesc.type;
            // longtext, varchar(n), decimal,int,bit,datetime,
            if (type == typeof(string))
            {
                if (tbdesc.MaxLength > 0 && itemValue.ToString().Length > tbdesc.MaxLength)
                {
                    ExcelTools.Utils.LogInfo($"第{rowIndex}行，'{scanDesc.HeaderName}'列 内容长度超过数据库定义长度！");
                    return false;
                }
                newValue = XmlEncoder.DecodeString(itemValue.ToString());
            }
            else if (type == typeof(DateTime))
            {

                if (itemValue is DateTime || itemValue is DateTime?)
                    newValue = itemValue;
                else if (DateTime.TryParse(itemValue.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var _v))
                    newValue = _v;
                else if (DateTime.TryParseExact(itemValue.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var _v2))
                    newValue = _v2;
                else if (double.TryParse(itemValue.ToString(), NumberStyles.None, CultureInfo.InvariantCulture, out var _d))
                    newValue = DateTimeHelper.FromOADate(_d);
                else
                {
                    ExcelTools.Utils.LogInfo($"第{rowIndex}行，'{scanDesc.HeaderName}'列 转<datetime>有误！");
                    return false;
                }
            }
            else if (type == typeof(bool))
            {
                var vs = itemValue.ToString();
                if (vs != "0" || vs != "1")
                {
                    ExcelTools.Utils.LogInfo($"第{rowIndex}行，'{scanDesc.HeaderName}'列  转<bool>值出错！");

                    return false;
                }

                newValue = (vs == "0" ? false : true);
            }
            else if (type == typeof(decimal))
            {
                if (!decimal.TryParse(itemValue.ToString(), out var _v2))
                {
                    ExcelTools.Utils.LogInfo($"第{rowIndex}行，'{scanDesc.HeaderName}'列 转<decimal>出错！");
                    return false;
                }

                if (scanDesc.RangeDecimal!=null && scanDesc.Prefix.Contains(ExcelTools.Utils.Config[EnumIdentifier.Range.ToString()]) 
                    && _v2<scanDesc.RangeDecimal.Item1 
                    &&_v2> scanDesc.RangeDecimal.Item2)  //范围
                {
                    ExcelTools.Utils.LogInfo($"第{rowIndex}行，'{scanDesc.HeaderName}'列 值超出范围！");
                    return false;
                }
                newValue = _v2;
            }
            else if (type == typeof(int))
            {
                if (!int.TryParse(itemValue.ToString(), out var _v2))
                {
                    ExcelTools.Utils.LogInfo($"第{rowIndex}行，'{scanDesc.HeaderName}'列 转int值出错！");
                    return false;
                }
                if (scanDesc.RangeInt != null && scanDesc.Prefix.Contains(ExcelTools.Utils.Config[EnumIdentifier.Range.ToString()])
                                                  && _v2 < scanDesc.RangeInt.Item1
                                                  && _v2 > scanDesc.RangeInt.Item2)  //范围
                {
                    ExcelTools.Utils.LogInfo($"第{rowIndex}行，'{scanDesc.HeaderName}'列 值超出范围！");
                    return false;
                }

                newValue = _v2;
            }
            else
            {
                return false;
                //newValue = Convert.ChangeType(itemValue, type);
            }

            return true;
        }
        #endregion
        /// <summary>
        /// 通过Type转List<Object>数据
        /// </summary>
        /// <param name="type"></param>
        /// <param name="sheetName"></param>
        /// <param name="startCell"></param>
        /// <param name="configuration"></param>
        /// <returns></returns>
        public IEnumerable<object> Query(Type type, string sheetName, string startCell, IConfiguration configuration)
        {

            List<Helpers.ExcelCustomPropertyInfo> props = null;
            var headers = Query(false, sheetName, startCell, configuration).FirstOrDefault()?.Values?.Select(s => s?.ToString())?.ToArray(); //TODO:need to optimize

            var first = true;
            var rowIndex = 0;
            foreach (var item in Query(true, sheetName, startCell, configuration))
            {
                if (first)
                {
                    //TODO: alert don't duplicate column name
                    props = CustomPropertyHelper.GetExcelCustomPropertyInfos(type, headers);
                    first = false;
                }
                //var v = new T();

                var v = Activator.CreateInstance(type);

                foreach (var pInfo in props)
                {
                    //TODO:don't need to check every time?
                    if (item.ContainsKey(pInfo.ExcelColumnName))
                    {
                        object newV = null;
                        object itemValue = item[pInfo.ExcelColumnName];

                        if (itemValue == null)
                            continue;

                        newV = TypeHelper.TypeMapping(v, pInfo, newV, itemValue, rowIndex, startCell);
                    }
                }
                rowIndex++;
                yield return v;
            }
        }

        public IEnumerable<T> Query<T>(string sheetName, string startCell, IConfiguration configuration) where T : class, new()
        {
            var type = typeof(T);

            List<Helpers.ExcelCustomPropertyInfo> props = null;
            var headers = Query(false, sheetName, startCell, configuration).FirstOrDefault()?.Values?.Select(s => s?.ToString())?.ToArray(); //TODO:need to optimize

            var first = true;
            var rowIndex = 0;
            foreach (var item in Query(true, sheetName, startCell, configuration))
            {
                if (first)
                {
                    //TODO: alert don't duplicate column name
                    props = CustomPropertyHelper.GetExcelCustomPropertyInfos(type, headers);
                    first = false;
                }
                var v = new T();
                foreach (var pInfo in props)
                {
                    //TODO:don't need to check every time?
                    if (item.ContainsKey(pInfo.ExcelColumnName))
                    {
                        object newV = null;
                        object itemValue = item[pInfo.ExcelColumnName];

                        if (itemValue == null)
                            continue;

                        newV = TypeHelper.TypeMapping(v, pInfo, newV, itemValue, rowIndex, startCell);
                    }
                }
                rowIndex++;
                yield return v;
            }
        }

        private void SetSharedStrings()
        {
            if (_sharedStrings != null)
                return;
            var sharedStringsEntry = _archive.GetEntry("xl/sharedStrings.xml");
            if (sharedStringsEntry == null)
                return;
            using (var stream = sharedStringsEntry.Open())
            {
                _sharedStrings = GetSharedStrings(stream).ToList();
            }
        }

        internal List<string> GetSharedStrings()
        {
            if (_sharedStrings == null)
                SetSharedStrings();
            return _sharedStrings;
        }

        private IEnumerable<string> GetSharedStrings(Stream stream)
        {
            using (var reader = XmlReader.Create(stream))
            {
                if (!reader.IsStartElement("sst", _ns))
                    yield break;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    yield break;

                while (!reader.EOF)
                {
                    if (reader.IsStartElement("si", _ns))
                    {
                        var value = StringHelper.ReadStringItem(reader);
                        yield return value;
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        break;
                    }
                }
            }
        }

        private void SetWorkbookRels(ReadOnlyCollection<ZipArchiveEntry> entries)
        {
            if (_sheetRecords != null)
                return;
            _sheetRecords = GetWorkbookRels(entries);
        }

        internal static IEnumerable<SheetRecord> ReadWorkbook(ReadOnlyCollection<ZipArchiveEntry> entries)
        {
            using (var stream = entries.Single(w => w.FullName == "xl/workbook.xml").Open())
            using (XmlReader reader = XmlReader.Create(stream, _xmlSettings))
            {
                if (!reader.IsStartElement("workbook", _ns))
                    yield break;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    yield break;

                while (!reader.EOF)
                {
                    if (reader.IsStartElement("sheets", _ns))
                    {
                        if (!XmlReaderHelper.ReadFirstContent(reader))
                            continue;

                        while (!reader.EOF)
                        {
                            if (reader.IsStartElement("sheet", _ns))
                            {
                                yield return new SheetRecord(
                                    reader.GetAttribute("name"),
                                    uint.Parse(reader.GetAttribute("sheetId")),
                                    reader.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                                );
                                reader.Skip();
                            }
                            else if (!XmlReaderHelper.SkipContent(reader))
                            {
                                break;
                            }
                        }
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        yield break;
                    }
                }
            }
        }

        internal static List<SheetRecord> GetWorkbookRels(ReadOnlyCollection<ZipArchiveEntry> entries)
        {
            var sheetRecords = ReadWorkbook(entries).ToList();

            using (var stream = entries.Single(w => w.FullName == "xl/_rels/workbook.xml.rels").Open())
            using (XmlReader reader = XmlReader.Create(stream, _xmlSettings))
            {
                if (!reader.IsStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships"))
                    return null;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    return null;

                while (!reader.EOF)
                {
                    if (reader.IsStartElement("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships"))
                    {
                        string rid = reader.GetAttribute("Id");
                        foreach (var sheet in sheetRecords)
                        {
                            if (sheet.Rid == rid)
                            {
                                sheet.Path = reader.GetAttribute("Target");
                                break;
                            }
                        }

                        reader.Skip();
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        break;
                    }
                }
            }

            return sheetRecords;
        }

        internal static DataTable QueryAsDataTableImpl(Stream stream, bool useHeaderRow, ref string sheetName, ExcelType excelType, string startCell, IConfiguration configuration)
        {
            if (sheetName == null)
                sheetName = stream.GetSheetNames().First();

            var dt = new DataTable(sheetName);
            var first = true;
            var rows = ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType)).Query(useHeaderRow, sheetName, startCell, configuration);
            foreach (IDictionary<string, object> row in rows)
            {
                if (first)
                {

                    foreach (var key in row.Keys)
                    {
                        var column = new DataColumn(key, typeof(object)) { Caption = key };
                        dt.Columns.Add(column);
                    }

                    dt.BeginLoadData();
                    first = false;
                }

                var newRow = dt.NewRow();
                foreach (var key in row.Keys)
                {
                    newRow[key] = row[key]; //TODO: optimize not using string key
                }

                dt.Rows.Add(newRow);
            }

            dt.EndLoadData();
            return dt;
        }

        private object ReadCellAndSetColumnIndex(XmlReader reader, ref int columnIndex, bool withoutCR, int startColumnIndex, string aR, string aT)
        {
            var newColumnIndex = 0;
            int xfIndex = -1;

            if (withoutCR)
                newColumnIndex = columnIndex + 1;
            //TODO:need to check only need nextColumnIndex or columnIndex
            else if (ReferenceHelper.ParseReference(aR, out int referenceColumn, out _))
                newColumnIndex = referenceColumn - 1; // ParseReference is 1-based
            else
                newColumnIndex = columnIndex;

            columnIndex = newColumnIndex;

            if (columnIndex < startColumnIndex)
            {
                if (!XmlReaderHelper.ReadFirstContent(reader))
                    return null;
                while (!reader.EOF)
                    if (!XmlReaderHelper.SkipContent(reader))
                        break;
                return null;
            }

            if (!XmlReaderHelper.ReadFirstContent(reader))
                return null;

            object value = null;
            while (!reader.EOF)
            {
                if (reader.IsStartElement("v", _ns))
                {
                    string rawValue = reader.ReadElementContentAsString();
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, xfIndex, out value);
                }
                else if (reader.IsStartElement("is", _ns))
                {
                    string rawValue = StringHelper.ReadStringItem(reader);
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, xfIndex, out value);
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }



            return value;
        }

        private void ConvertCellValue(string rawValue, string aT, int xfIndex, out object value)
        {
            const NumberStyles style = NumberStyles.Any;
            var invariantCulture = CultureInfo.InvariantCulture;

            switch (aT)
            {
                case "s":
                    if (int.TryParse(rawValue, style, invariantCulture, out var sstIndex))
                    {
                        if (sstIndex >= 0 && sstIndex < _sharedStrings.Count)
                        {
                            //value = Helpers.ConvertEscapeChars(_SharedStrings[sstIndex]);
                            value = XmlEncoder.DecodeString(_sharedStrings[sstIndex]);
                            return;
                        }
                    }
                    value = null;
                    return;
                case "inlineStr":
                case "str":
                    value = XmlEncoder.DecodeString(rawValue);
                    return;
                case "b":
                    value = rawValue == "1";
                    return;
                case "d":
                    if (DateTime.TryParseExact(rawValue, "yyyy-MM-dd", invariantCulture, DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite, out var date))
                    {
                        value = date;
                        return;
                    }

                    value = rawValue;
                    return;
                case "e":
                    value = rawValue;
                    return;
                default:
                    if (double.TryParse(rawValue, style, invariantCulture, out var n))
                    {
                        value = n;
                        return;
                    }

                    value = rawValue;
                    return;
            }
        }

        public Task<IEnumerable<IDictionary<string, object>>> QueryAsync(bool UseHeaderRow, string sheetName, string startCell, IConfiguration configuration)
        {
            return Task.Run(() => Query(UseHeaderRow, sheetName, startCell, configuration));
        }

        public Task<IEnumerable<T>> QueryAsync<T>(string sheetName, string startCell, IConfiguration configuration) where T : class, new()
        {
            return Task.Run(() => Query<T>(sheetName, startCell, configuration));
        }
    }
}
