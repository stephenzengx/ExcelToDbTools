
using Excel.Attributes;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Excel.Utils
{
    internal static class Helpers
    {
        private const int GENERAL_COLUMN_INDEX = 255;
        private const int MAX_COLUMN_INDEX = 16383;
        private static Dictionary<int, string> _IntMappingAlphabet;
        private static Dictionary<string, int> _AlphabetMappingInt;

        public static FileStream OpenSharedRead(string path)
        {
            return File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }

        static Helpers()
        {
            if (Helpers._IntMappingAlphabet != null || Helpers._AlphabetMappingInt != null)
                return;
            Helpers._IntMappingAlphabet = new Dictionary<int, string>();
            Helpers._AlphabetMappingInt = new Dictionary<string, int>();
            for (int key = 0; key <= (int)byte.MaxValue; ++key)
            {
                Helpers._IntMappingAlphabet.Add(key, Helpers.IntToLetters(key));
                Helpers._AlphabetMappingInt.Add(Helpers.IntToLetters(key), key);
            }
        }

        public static string GetAlphabetColumnName(int columnIndex)
        {
            Helpers.CheckAndSetMaxColumnIndex(columnIndex);
            return Helpers._IntMappingAlphabet[columnIndex];
        }

        public static int GetColumnIndex(string columnName)
        {
            int columnIndex = Helpers._AlphabetMappingInt[columnName];
            Helpers.CheckAndSetMaxColumnIndex(columnIndex);
            return columnIndex;
        }

        private static void CheckAndSetMaxColumnIndex(int columnIndex)
        {
            if (columnIndex < Helpers._IntMappingAlphabet.Count)
                return;
            if (columnIndex > 16383)
                throw new InvalidDataException(string.Format("ColumnIndex {0} over excel vaild max index.", (object)columnIndex));
            for (int count = Helpers._IntMappingAlphabet.Count; count <= columnIndex; ++count)
            {
                Helpers._IntMappingAlphabet.Add(count, Helpers.IntToLetters(count));
                Helpers._AlphabetMappingInt.Add(Helpers.IntToLetters(count), count);
            }
        }

        internal static string IntToLetters(int value)
        {
            ++value;
            string str = string.Empty;
            for (; --value >= 0; value /= 26)
                str = ((char)(65 + value % 26)).ToString() + str;
            return str;
        }

        internal static IDictionary<string, object> GetEmptyExpandoObject(
          int maxColumnIndex,
          int startCellIndex)
        {
            IDictionary<string, object> dictionary = (IDictionary<string, object>)new ExpandoObject();
            for (int columnIndex = startCellIndex; columnIndex <= maxColumnIndex; ++columnIndex)
            {
                string alphabetColumnName = Helpers.GetAlphabetColumnName(columnIndex);
                if (!dictionary.ContainsKey(alphabetColumnName))
                    dictionary.Add(alphabetColumnName, (object)null);
            }
            return dictionary;
        }

        internal static IDictionary<string, object> GetEmptyExpandoObject(
          Dictionary<int, string> hearrows)
        {
            IDictionary<string, object> dictionary = (IDictionary<string, object>)new ExpandoObject();
            foreach (KeyValuePair<int, string> hearrow in hearrows)
            {
                if (!dictionary.ContainsKey(hearrow.Value))
                    dictionary.Add(hearrow.Value, (object)null);
            }
            return dictionary;
        }

        internal static List<Helpers.ExcelCustomPropertyInfo> GetSaveAsProperties(
          this Type type)
        {
            List<Helpers.ExcelCustomPropertyInfo> list1 = Helpers.GetExcelPropertyInfo(type, BindingFlags.Instance | BindingFlags.Public).Where<Helpers.ExcelCustomPropertyInfo>((Func<Helpers.ExcelCustomPropertyInfo, bool>)(prop =>
            {
                if (prop.Property.GetGetMethod() != (MethodInfo)null)
                    return !prop.Property.GetAttributeValue<ExcelIgnoreAttribute, bool>((Func<ExcelIgnoreAttribute, bool>)(x => x.ExcelIgnore), true);
                return false;
            })).ToList<Helpers.ExcelCustomPropertyInfo>();
            if (list1.Count == 0)
                throw new InvalidOperationException(type.Name + " un-ignore properties count can't be 0");
            IEnumerable<Helpers.ExcelCustomPropertyInfo> source = list1.Where<Helpers.ExcelCustomPropertyInfo>((Func<Helpers.ExcelCustomPropertyInfo, bool>)(w =>
            {
                if (!w.ExcelColumnIndex.HasValue)
                    return false;
                int? excelColumnIndex = w.ExcelColumnIndex;
                int num = -1;
                return excelColumnIndex.GetValueOrDefault() > num & excelColumnIndex.HasValue;
            }));
            if (source.GroupBy<Helpers.ExcelCustomPropertyInfo, int?>((Func<Helpers.ExcelCustomPropertyInfo, int?>)(g => g.ExcelColumnIndex)).Any<IGrouping<int?, Helpers.ExcelCustomPropertyInfo>>((Func<IGrouping<int?, Helpers.ExcelCustomPropertyInfo>, bool>)(_ => _.Count<Helpers.ExcelCustomPropertyInfo>() > 1)))
                throw new InvalidOperationException("Duplicate column name");
            int val2 = list1.Count - 1;
            if (source.Any<Helpers.ExcelCustomPropertyInfo>())
                val2 = Math.Max(source.Max<Helpers.ExcelCustomPropertyInfo>((Func<Helpers.ExcelCustomPropertyInfo, int?>)(w => w.ExcelColumnIndex)).Value, val2);
            List<Helpers.ExcelCustomPropertyInfo> list2 = list1.Where<Helpers.ExcelCustomPropertyInfo>((Func<Helpers.ExcelCustomPropertyInfo, bool>)(w => !w.ExcelColumnIndex.HasValue)).ToList<Helpers.ExcelCustomPropertyInfo>();
            List<Helpers.ExcelCustomPropertyInfo> customPropertyInfoList = new List<Helpers.ExcelCustomPropertyInfo>();
            int index = 0;
            for (int i = 0; i <= val2; i++)
            {
                Helpers.ExcelCustomPropertyInfo customPropertyInfo1 = source.SingleOrDefault<Helpers.ExcelCustomPropertyInfo>((Func<Helpers.ExcelCustomPropertyInfo, bool>)(s =>
                {
                    int? excelColumnIndex = s.ExcelColumnIndex;
                    int num = i;
                    return excelColumnIndex.GetValueOrDefault() == num & excelColumnIndex.HasValue;
                }));
                if (customPropertyInfo1 != null)
                {
                    customPropertyInfoList.Add(customPropertyInfo1);
                }
                else
                {
                    Helpers.ExcelCustomPropertyInfo customPropertyInfo2 = list2.ElementAtOrDefault<Helpers.ExcelCustomPropertyInfo>(index);
                    if (customPropertyInfo2 == null)
                    {
                        customPropertyInfoList.Add((Helpers.ExcelCustomPropertyInfo)null);
                    }
                    else
                    {
                        customPropertyInfo2.ExcelColumnIndex = new int?(i);
                        customPropertyInfoList.Add(customPropertyInfo2);
                    }
                    ++index;
                }
            }
            return customPropertyInfoList;
        }

        internal static List<Helpers.ExcelCustomPropertyInfo> GetExcelCustomPropertyInfos(
          Type type,
          string[] headers)
        {
            List<Helpers.ExcelCustomPropertyInfo> list = Helpers.GetExcelPropertyInfo(type, BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty).Where<Helpers.ExcelCustomPropertyInfo>((Func<Helpers.ExcelCustomPropertyInfo, bool>)(prop =>
            {
                if (prop.Property.GetSetMethod() != (MethodInfo)null)
                    return !prop.Property.GetAttributeValue<ExcelIgnoreAttribute, bool>((Func<ExcelIgnoreAttribute, bool>)(x => x.ExcelIgnore), true);
                return false;
            })).ToList<Helpers.ExcelCustomPropertyInfo>();
            if (list.Count == 0)
                throw new InvalidOperationException(type.Name + " un-ignore properties count can't be 0");
            if (list.Where<Helpers.ExcelCustomPropertyInfo>((Func<Helpers.ExcelCustomPropertyInfo, bool>)(w =>
            {
                if (!w.ExcelColumnIndex.HasValue)
                    return false;
                int? excelColumnIndex = w.ExcelColumnIndex;
                int num = -1;
                return excelColumnIndex.GetValueOrDefault() > num & excelColumnIndex.HasValue;
            })).GroupBy<Helpers.ExcelCustomPropertyInfo, int?>((Func<Helpers.ExcelCustomPropertyInfo, int?>)(g => g.ExcelColumnIndex)).Any<IGrouping<int?, Helpers.ExcelCustomPropertyInfo>>((Func<IGrouping<int?, Helpers.ExcelCustomPropertyInfo>, bool>)(_ => _.Count<Helpers.ExcelCustomPropertyInfo>() > 1)))
                throw new InvalidOperationException("Duplicate column name");
            foreach (Helpers.ExcelCustomPropertyInfo customPropertyInfo1 in list)
            {
                int? excelColumnIndex = customPropertyInfo1.ExcelColumnIndex;
                if (excelColumnIndex.HasValue)
                {
                    excelColumnIndex = customPropertyInfo1.ExcelColumnIndex;
                    int length = headers.Length;
                    if (excelColumnIndex.GetValueOrDefault() >= length & excelColumnIndex.HasValue)
                        throw new ArgumentException(string.Format("ExcelColumnIndex {0} over haeder max index {1}", (object)customPropertyInfo1.ExcelColumnIndex, (object)headers.Length));
                    Helpers.ExcelCustomPropertyInfo customPropertyInfo2 = customPropertyInfo1;
                    string[] strArray = headers;
                    excelColumnIndex = customPropertyInfo1.ExcelColumnIndex;
                    int index = excelColumnIndex.Value;
                    string str = strArray[index];
                    customPropertyInfo2.ExcelColumnName = str;
                    if (customPropertyInfo1.ExcelColumnName == null)
                        throw new InvalidOperationException(string.Format("{0} {1}'s ExcelColumnIndex {2} can't find excel column name", (object)customPropertyInfo1.Property.DeclaringType.Name, (object)customPropertyInfo1.Property.Name, (object)customPropertyInfo1.ExcelColumnIndex));
                }
            }
            return list;
        }

        private static IEnumerable<Helpers.ExcelCustomPropertyInfo> GetExcelPropertyInfo(
          Type type,
          BindingFlags bindingFlags)
        {
            return ((IEnumerable<PropertyInfo>)type.GetProperties(bindingFlags)).Select<PropertyInfo, Helpers.ExcelCustomPropertyInfo>((Func<PropertyInfo, Helpers.ExcelCustomPropertyInfo>)(p =>
            {
                Type underlyingType = Nullable.GetUnderlyingType(p.PropertyType);
                ExcelColumnNameAttribute attribute1 = p.GetAttribute<ExcelColumnNameAttribute>(true);
                ExcelColumnIndexAttribute attribute2 = p.GetAttribute<ExcelColumnIndexAttribute>(true);
                Helpers.ExcelCustomPropertyInfo customPropertyInfo = new Helpers.ExcelCustomPropertyInfo();
                customPropertyInfo.Property = p;
                Type type1 = underlyingType;
                if ((object)type1 == null)
                    type1 = p.PropertyType;
                customPropertyInfo.ExcludeNullableType = type1;
                customPropertyInfo.Nullable = underlyingType != (Type)null;
                customPropertyInfo.ExcelColumnName = attribute1?.ExcelColumnName ?? p.Name;
                customPropertyInfo.ExcelColumnIndex = attribute2?.ExcelColumnIndex;
                customPropertyInfo.ExcelFormat = p.GetAttribute<ExcelFormatAttribute>(true)?.Format;
                return customPropertyInfo;
            }));
        }

        public static bool IsNumericType(Type type, bool isNullableUnderlyingType = false)
        {
            if (isNullableUnderlyingType)
            {
                Type type1 = Nullable.GetUnderlyingType(type);
                if ((object)type1 == null)
                    type1 = type;
                type = type1;
            }
            return (uint)(Type.GetTypeCode(type) - 7) <= 8U;
        }

        public static object TypeMapping<T>(
          T v,
          Helpers.ExcelCustomPropertyInfo pInfo,
          object newValue,
          object itemValue,
          int rowIndex,
          string startCell)
          where T : class, new()
        {
            try
            {
                return Helpers.TypeMappingImpl<T>(v, pInfo, ref newValue, itemValue);
            }
            catch (Exception ex) when (ex is InvalidCastException || ex is FormatException)
            {
                throw new InvalidCastException(string.Format("ColumnName : {0}, CellRow : {1}, Value : {2}, it can't cast to {3} type.", (object)(pInfo.ExcelColumnName ?? pInfo.Property.Name), (object)(ReferenceHelper.ConvertCellToXY(startCell).Item2 + rowIndex + 1), itemValue, (object)pInfo.Property.PropertyType.Name));
            }
        }

        private static object TypeMappingImpl<T>(
          T v,
          Helpers.ExcelCustomPropertyInfo pInfo,
          ref object newValue,
          object itemValue)
          where T : class, new()
        {
            if (pInfo.ExcludeNullableType == typeof(Guid))
                newValue = (object)Guid.Parse(itemValue.ToString());
            else if (pInfo.ExcludeNullableType == typeof(DateTime))
            {
                if (itemValue is DateTime || itemValue is DateTime?)
                {
                    newValue = itemValue;
                    pInfo.Property.SetValue((object)v, newValue);
                    return newValue;
                }
                string s = itemValue?.ToString();
                if (pInfo.ExcelFormat != null)
                {
                    DateTime result;
                    if (DateTime.TryParseExact(s, pInfo.ExcelFormat, (IFormatProvider)CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
                        newValue = (object)result;
                }
                else
                {
                    DateTime result1;
                    if (DateTime.TryParse(s, (IFormatProvider)CultureInfo.InvariantCulture, DateTimeStyles.None, out result1))
                    {
                        newValue = (object)result1;
                    }
                    else
                    {
                        DateTime result2;
                        if (DateTime.TryParseExact(s, "dd/MM/yyyy", (IFormatProvider)CultureInfo.InvariantCulture, DateTimeStyles.None, out result2))
                        {
                            newValue = (object)result2;
                        }
                        else
                        {
                            double result3;
                            if (!double.TryParse(s, NumberStyles.None, (IFormatProvider)CultureInfo.InvariantCulture, out result3))
                                throw new InvalidCastException(s + " can't cast to datetime");
                            newValue = (object)DateTimeHelper.FromOADate(result3);
                        }
                    }
                }
            }
            else if (pInfo.ExcludeNullableType == typeof(bool))
            {
                string str = itemValue.ToString();
                newValue = !(str == "1") ? (!(str == "0") ? (object)bool.Parse(str) : (object)false) : (object)true;
            }
            else
                newValue = !(pInfo.Property.PropertyType == typeof(string)) ? (!pInfo.Property.PropertyType.IsEnum ? Convert.ChangeType(itemValue, pInfo.ExcludeNullableType) : Enum.Parse(pInfo.Property.PropertyType, itemValue?.ToString(), true)) : (object)XmlEncoder.DecodeString(itemValue?.ToString());
            pInfo.Property.SetValue((object)v, newValue);
            return newValue;
        }

        internal class ExcelCustomPropertyInfo
        {
            public int? ExcelColumnIndex { get; set; }

            public string ExcelColumnName { get; set; }

            public PropertyInfo Property { get; set; }

            public Type ExcludeNullableType { get; set; }

            public bool Nullable { get; internal set; }

            public string ExcelFormat { get; internal set; }
        }
    }
}
