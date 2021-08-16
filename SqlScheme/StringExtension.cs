using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace ExcelTools
{
    /// <summary>
    /// 字符串<see cref="T:System.String" />类型的扩展辅助操作类
    /// </summary>
    public static class StringExtension
    {
        /// <summary>指示所指定的正则表达式在指定的输入字符串中是否找到了匹配项</summary>
        /// <param name="value">要搜索匹配项的字符串</param>
        /// <param name="pattern">要匹配的正则表达式模式</param>
        /// <param name="isContains">是否包含，否则全匹配</param>
        /// <returns>如果正则表达式找到匹配项，则为 true；否则，为 false</returns>
        public static bool IsMatch(this string value, string pattern, bool isContains = true)
        {
            if (value == null)
                return false;
            if (!isContains)
                return Regex.Match(value, pattern).Success;
            return Regex.IsMatch(value, pattern);
        }

        /// <summary>在指定的输入字符串中搜索指定的正则表达式的第一个匹配项</summary>
        /// <param name="value">要搜索匹配项的字符串</param>
        /// <param name="pattern">要匹配的正则表达式模式</param>
        /// <returns>一个对象，包含有关匹配项的信息</returns>
        public static string Match(this string value, string pattern)
        {
            return value ?? Regex.Match(value, pattern).Value;
        }

        /// <summary>在指定的输入字符串中搜索指定的正则表达式的所有匹配项的字符串集合</summary>
        /// <param name="value"> 要搜索匹配项的字符串 </param>
        /// <param name="pattern"> 要匹配的正则表达式模式 </param>
        /// <returns> 一个集合，包含有关匹配项的字符串值 </returns>
        public static IEnumerable<string> Matches(this string value, string pattern)
        {
            if (value == null)
                return (IEnumerable<string>)Array.Empty<string>();
            return Regex.Matches(value, pattern).Cast<Match>().Select<Match, string>((Func<Match, string>)(match => match.Value));
        }

        /// <summary>在指定的输入字符串中匹配第一个数字字符串</summary>
        public static string MatchFirstNumber(this string value)
        {
            MatchCollection matchCollection = Regex.Matches(value, "\\d+");
            if (matchCollection.Count == 0)
                return string.Empty;
            return matchCollection[0].Value;
        }

        /// <summary>在指定字符串中匹配最后一个数字字符串</summary>
        public static string MatchLastNumber(this string value)
        {
            MatchCollection matchCollection = Regex.Matches(value, "\\d+");
            if (matchCollection.Count == 0)
                return string.Empty;
            return matchCollection[matchCollection.Count - 1].Value;
        }

        /// <summary>在指定字符串中匹配所有数字字符串</summary>
        public static IEnumerable<string> MatchNumbers(this string value)
        {
            return value.Matches("\\d+");
        }

        /// <summary>检测指定字符串中是否包含数字</summary>
        public static bool IsMatchNumber(this string value)
        {
            return value.IsMatch("\\d", true);
        }

        /// <summary>检测指定字符串是否全部为数字并且长度等于指定长度</summary>
        public static bool IsMatchNumber(this string value, int length)
        {
            return new Regex("^\\d{" + length.ToString() + "}$").IsMatch(value);
        }

        /// <summary>用正则表达式截取字符串</summary>
        public static string Substring2(this string source, string startString, string endString)
        {
            return source.Substring2(startString, endString, false);
        }

        /// <summary>用正则表达式截取字符串</summary>
        public static string Substring2(
          this string source,
          string startString,
          string endString,
          bool containsEmpty)
        {
            if (source.IsMissing())
                return string.Empty;
            string str1 = containsEmpty ? "\\s\\S" : "\\S";
            string str2 = source.Match(string.Format("(?<={0})([{1}]+?)(?={2})", (object)startString, (object)str1, (object)endString));
            if (!str2.IsMissing())
                return str2;
            return (string)null;
        }

        /// <summary>是否电子邮件</summary>
        public static bool IsEmail(this string value)
        {
            return value.IsMatch("^[\\w-]+(\\.[\\w-]+)*@[\\w-]+(\\.[\\w-]+)+$", true);
        }

        /// <summary>是否是IP地址</summary>
        public static bool IsIpAddress(this string value)
        {
            return value.IsMatch("^((?:(?:25[0-5]|2[0-4]\\d|((1\\d{2})|([1-9]?\\d)))\\.){3}(?:25[0-5]|2[0-4]\\d|((1\\d{2})|([1-9]?\\d))))$", true);
        }

        /// <summary>是否是整数</summary>
        public static bool IsNumeric(this string value)
        {
            return value.IsMatch("^\\-?[0-9]+$", true);
        }

        /// <summary>是否是Unicode字符串</summary>
        public static bool IsUnicode(this string value)
        {
            return value.IsMatch("^[\\u4E00-\\u9FA5\\uE815-\\uFA29]+$", true);
        }

        /// <summary>是否Url字符串</summary>
        public static bool IsUrl(this string value)
        {
            try
            {
                if (value.IsNullOrEmptyXYS() || value.Contains(' '))
                    return false;
                Uri uri = new Uri(value);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 是否身份证号，验证如下3种情况：
        /// 1.身份证号码为15位数字；
        /// 2.身份证号码为18位数字；
        /// 3.身份证号码为17位数字+1个字母
        /// </summary>
        public static bool IsIdentityCardId(this string value)
        {
            if (value.Length != 15 && value.Length != 18)
                return false;
            DateTime result;
            if (value.Length == 15)
            {
                Regex regex = new Regex("^(\\d{6})(\\d{2})(\\d{2})(\\d{2})(\\d{3})_");
                if (!regex.Match(value).Success)
                    return false;
                string[] strArray = regex.Split(value);
                return DateTime.TryParse(string.Format("{0}-{1}-{2}", (object)("19" + strArray[2]), (object)strArray[3], (object)strArray[4]), out result);
            }
            Regex regex1 = new Regex("^(\\d{6})(\\d{4})(\\d{2})(\\d{2})(\\d{3})([0-9Xx])$");
            if (!regex1.Match(value).Success)
                return false;
            string[] strArray1 = regex1.Split(value);
            if (!DateTime.TryParse(string.Format("{0}-{1}-{2}", (object)strArray1[2], (object)strArray1[3], (object)strArray1[4]), out result))
                return false;
            string[] array = ((IEnumerable<char>)value.ToCharArray()).Select<char, string>((Func<char, string>)(m => m.ToString())).ToArray<string>();
            int[] numArray = new int[17]
            {
        7,
        9,
        10,
        5,
        8,
        4,
        2,
        1,
        6,
        3,
        7,
        9,
        10,
        5,
        8,
        4,
        2
            };
            int num1 = 0;
            for (int index = 0; index < 17; ++index)
            {
                int num2 = int.Parse(array[index]);
                num1 += num2 * numArray[index];
            }
            string str = ((IEnumerable<char>)"10X98765432".ToCharArray()).ElementAt<char>(num1 % 11).ToString();
            return ((IEnumerable<string>)array).Last<string>().ToUpper() == str;
        }

        /// <summary>是否手机号码</summary>
        /// <param name="value"></param>
        /// <param name="isRestrict">是否按严格格式验证</param>
        public static bool IsMobileNumber(this string value, bool isRestrict = false)
        {
            string pattern = isRestrict ? "^[1][3-8]\\d{9}$" : "^[1]\\d{10}$";
            return value.IsMatch(pattern, true);
        }

        /// <summary>指示指定的字符串是 null 或者 System.String.Empty 字符串</summary>
        public static bool IsNullOrEmptyXYS(this string value)
        {
            return string.IsNullOrEmpty(value);
        }

        /// <summary>指示指定的字符串是 null、空或者仅由空白字符组成。</summary>
        public static bool IsNullOrWhiteSpace(this string value)
        {
            return string.IsNullOrWhiteSpace(value);
        }

        /// <summary>指示指定的字符串是 null、空或者仅由空白字符组成。</summary>
        public static bool IsMissing(this string value)
        {
            return string.IsNullOrWhiteSpace(value);
        }

        /// <summary>为指定格式的字符串填充相应对象来生成字符串</summary>
        /// <param name="format">字符串格式，占位符以{n}表示</param>
        /// <param name="args">用于填充占位符的参数</param>
        /// <returns>格式化后的字符串</returns>
        public static string FormatWith(this string format, params object[] args)
        {
            return string.Format((IFormatProvider)CultureInfo.CurrentCulture, format, args);
        }

        /// <summary>将字符串反转</summary>
        /// <param name="value">要反转的字符串</param>
        public static string ReverseString(this string value)
        {
            return new string(value.Reverse<char>().ToArray<char>());
        }

        /// <summary>单词变成单数形式</summary>
        /// <param name="word"></param>
        /// <returns></returns>
        public static string ToSingular(this string word)
        {
            Regex regex1 = new Regex("(?<keep>[^aeiou])ies$");
            Regex regex2 = new Regex("(?<keep>[aeiou]y)s$");
            Regex regex3 = new Regex("(?<keep>[sxzh])es$");
            Regex regex4 = new Regex("(?<keep>[^sxzhyu])s$");
            if (regex1.IsMatch(word))
                return regex1.Replace(word, "${keep}y");
            if (regex2.IsMatch(word))
                return regex2.Replace(word, "${keep}");
            if (regex3.IsMatch(word))
                return regex3.Replace(word, "${keep}");
            if (regex4.IsMatch(word))
                return regex4.Replace(word, "${keep}");
            return word;
        }

        /// <summary>单词变成复数形式</summary>
        /// <param name="word"></param>
        /// <returns></returns>
        public static string ToPlural(this string word)
        {
            Regex regex1 = new Regex("(?<keep>[^aeiou])y$");
            Regex regex2 = new Regex("(?<keep>[aeiou]y)$");
            Regex regex3 = new Regex("(?<keep>[sxzh])$");
            Regex regex4 = new Regex("(?<keep>[^sxzhy])$");
            if (regex1.IsMatch(word))
                return regex1.Replace(word, "${keep}ies");
            if (regex2.IsMatch(word))
                return regex2.Replace(word, "${keep}s");
            if (regex3.IsMatch(word))
                return regex3.Replace(word, "${keep}es");
            if (regex4.IsMatch(word))
                return regex4.Replace(word, "${keep}s");
            return word;
        }

        /// <summary>判断指定路径是否图片文件</summary>
        public static bool IsImageFile(this string filename)
        {
            if (!File.Exists(filename))
                return false;
            byte[] numArray = File.ReadAllBytes(filename);
            if (numArray.Length == 0)
                return false;
            switch (BitConverter.ToUInt16(numArray, 0))
            {
                case 18759:
                case 19778:
                case 20617:
                case 55551:
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>以指定字符串作为分隔符将指定字符串分隔成数组</summary>
        /// <param name="value">要分割的字符串</param>
        /// <param name="strSplit">字符串类型的分隔符</param>
        /// <param name="removeEmptyEntries">是否移除数据中元素为空字符串的项</param>
        /// <returns>分割后的数据</returns>
        public static string[] Split(this string value, string strSplit, bool removeEmptyEntries = false)
        {
            return value.Split(new string[1] { strSplit }, (StringSplitOptions)(removeEmptyEntries ? 1 : 0));
        }

        /// <summary>支持汉字的字符串长度，汉字长度计为2</summary>
        /// <param name="value">参数字符串</param>
        /// <returns>当前字符串的长度，汉字长度为2</returns>
        public static int TextLength(this string value)
        {
            ASCIIEncoding asciiEncoding = new ASCIIEncoding();
            int num1 = 0;
            string s = value;
            foreach (byte num2 in asciiEncoding.GetBytes(s))
            {
                if (num2 == (byte)63)
                    num1 += 2;
                else
                    ++num1;
            }
            return num1;
        }

        /// <summary>给URL添加查询参数</summary>
        /// <param name="url">URL字符串</param>
        /// <param name="queries">要添加的参数，形如："id=1,cid=2"</param>
        /// <returns></returns>
        public static string AddUrlQuery(this string url, params string[] queries)
        {
            foreach (string query in queries)
            {
                if (!url.Contains("?"))
                    url += "?";
                else if (!url.EndsWith("&"))
                    url += "&";
                url += query;
            }
            return url;
        }

        /// <summary>获取URL中指定参数的值，不存在返回空字符串</summary>
        public static string GetUrlQuery(this string url, string key)
        {
            string query = new Uri(url).Query;
            if (query.IsNullOrEmptyXYS())
                return string.Empty;
            Dictionary<string, string> dictionary = ((IEnumerable<string>)query.TrimStart('?').Split("&", true)).Select(m => new
            {
                m = m,
                strs = m.Split("=", StringSplitOptions.None)
            }).Select(_param1 => new KeyValuePair<string, string>(_param1.strs[0], _param1.strs[1])).ToDictionary<KeyValuePair<string, string>, string, string>((Func<KeyValuePair<string, string>, string>)(m => m.Key), (Func<KeyValuePair<string, string>, string>)(m => m.Value));
            if (dictionary.ContainsKey(key))
                return dictionary[key];
            return string.Empty;
        }

        /// <summary>给URL添加 # 参数</summary>
        /// <param name="url">URL字符串</param>
        /// <param name="query">要添加的参数</param>
        /// <returns></returns>
        public static string AddHashFragment(this string url, string query)
        {
            if (!url.Contains("#"))
                url += "#";
            return url + query;
        }

        /// <summary>
        /// 将字符串转换为<see cref="T:System.Byte" />[]数组，默认编码为<see cref="P:System.Text.Encoding.UTF8" />
        /// </summary>
        public static byte[] ToBytes(this string value, Encoding encoding = null)
        {
            if (encoding == null)
                encoding = Encoding.UTF8;
            return encoding.GetBytes(value);
        }

        /// <summary>
        /// 将<see cref="T:System.Byte" />[]数组转换为字符串，默认编码为<see cref="P:System.Text.Encoding.UTF8" />
        /// </summary>
        public static string ToString2(this byte[] bytes, Encoding encoding = null)
        {
            if (encoding == null)
                encoding = Encoding.UTF8;
            return encoding.GetString(bytes);
        }

        /// <summary>
        /// 将<see cref="T:System.Byte" />[]数组转换为Base64字符串
        /// </summary>
        public static string ToBase64String(this byte[] bytes)
        {
            return Convert.ToBase64String(bytes);
        }

        /// <summary>
        /// 将字符串转换为Base64字符串，默认编码为<see cref="P:System.Text.Encoding.UTF8" />
        /// </summary>
        /// <param name="source">正常的字符串</param>
        /// <param name="encoding">编码</param>
        /// <returns>Base64字符串</returns>
        public static string ToBase64String(this string source, Encoding encoding = null)
        {
            if (encoding == null)
                encoding = Encoding.UTF8;
            return Convert.ToBase64String(encoding.GetBytes(source));
        }

        /// <summary>
        /// 将Base64字符串转换为正常字符串，默认编码为<see cref="P:System.Text.Encoding.UTF8" />
        /// </summary>
        /// <param name="base64String">Base64字符串</param>
        /// <param name="encoding">编码</param>
        /// <returns>正常字符串</returns>
        public static string FromBase64String(this string base64String, Encoding encoding = null)
        {
            if (encoding == null)
                encoding = Encoding.UTF8;
            byte[] bytes = Convert.FromBase64String(base64String);
            return encoding.GetString(bytes);
        }

        /// <summary>将字符串进行UrlDecode解码</summary>
        /// <param name="source">待UrlDecode解码的字符串</param>
        /// <returns>UrlDecode解码后的字符串</returns>
        public static string ToUrlDecode(this string source)
        {
            return HttpUtility.UrlDecode(source);
        }

        /// <summary>将字符串进行UrlEncode编码</summary>
        /// <param name="source">待UrlEncode编码的字符串</param>
        /// <returns>UrlEncode编码后的字符串</returns>
        public static string ToUrlEncode(this string source)
        {
            return HttpUtility.UrlEncode(source);
        }

        /// <summary>将字符串进行HtmlDecode解码</summary>
        /// <param name="source">待HtmlDecode解码的字符串</param>
        /// <returns>HtmlDecode解码后的字符串</returns>
        public static string ToHtmlDecode(this string source)
        {
            return HttpUtility.HtmlDecode(source);
        }

        /// <summary>将字符串进行HtmlEncode编码</summary>
        /// <param name="source">待HtmlEncode编码的字符串</param>
        /// <returns>HtmlEncode编码后的字符串</returns>
        public static string ToHtmlEncode(this string source)
        {
            return HttpUtility.HtmlEncode(source);
        }

        /// <summary>
        /// 将字符串转换为十六进制字符串，默认编码为<see cref="P:System.Text.Encoding.UTF8" />
        /// </summary>
        public static string ToHexString(this string source, Encoding encoding = null)
        {
            if (encoding == null)
                encoding = Encoding.UTF8;
            return encoding.GetBytes(source).ToHexString();
        }

        /// <summary>
        /// 将十六进制字符串转换为常规字符串，默认编码为<see cref="P:System.Text.Encoding.UTF8" />
        /// </summary>
        public static string FromHexString(this string hexString, Encoding encoding = null)
        {
            if (encoding == null)
                encoding = Encoding.UTF8;
            byte[] hexBytes = hexString.ToHexBytes();
            return encoding.GetString(hexBytes);
        }

        /// <summary>将byte[]编码为十六进制字符串</summary>
        /// <param name="bytes">byte[]数组</param>
        /// <returns>十六进制字符串</returns>
        public static string ToHexString(this byte[] bytes)
        {
            return ((IEnumerable<byte>)bytes).Aggregate<byte, string>(string.Empty, (Func<string, byte, string>)((current, t) => current + t.ToString("X2")));
        }

        /// <summary>将十六进制字符串转换为byte[]</summary>
        /// <param name="hexString">十六进制字符串</param>
        /// <returns>byte[]数组</returns>
        public static byte[] ToHexBytes(this string hexString)
        {
            if (hexString == null)
                hexString = "";
            hexString = hexString.Replace(" ", "");
            byte[] numArray = new byte[hexString.Length / 2];
            for (int index = 0; index < numArray.Length; ++index)
                numArray[index] = Convert.ToByte(hexString.Substring(index * 2, 2), 16);
            return numArray;
        }

        /// <summary>将字符串进行Unicode编码，变成形如“\u7f16\u7801”的形式</summary>
        /// <param name="source">要进行编号的字符串</param>
        public static string ToUnicodeString(this string source)
        {
            return new Regex("[^\\u0000-\\u00ff]").Replace(source, (MatchEvaluator)(m => string.Format("\\u{0:x4}", (object)(short)m.Value[0])));
        }

        /// <summary>将形如“\u7f16\u7801”的Unicode字符串解码</summary>
        public static string FromUnicodeString(this string source)
        {
            return new Regex("\\\\u([0-9a-fA-F]{4})", RegexOptions.Compiled).Replace(source, (MatchEvaluator)(m =>
            {
                short result;
                if (short.TryParse(m.Groups[1].Value, NumberStyles.HexNumber, (IFormatProvider)CultureInfo.InstalledUICulture, out result))
                    return ((char)result).ToString() ?? "";
                return m.Value;
            }));
        }

        /// <summary>将驼峰字符串的第一个字符小写</summary>
        public static string LowerFirstChar(this string str)
        {
            if (string.IsNullOrEmpty(str) || !char.IsUpper(str[0]))
                return str;
            if (str.Length == 1)
                return char.ToLower(str[0]).ToString();
            return char.ToLower(str[0]).ToString() + str.Substring(1, str.Length - 1);
        }

        /// <summary>将小驼峰字符串的第一个字符大写</summary>
        public static string UpperFirstChar(this string str)
        {
            if (string.IsNullOrEmpty(str) || !char.IsLower(str[0]))
                return str;
            if (str.Length == 1)
                return char.ToUpper(str[0]).ToString();
            return char.ToUpper(str[0]).ToString() + str.Substring(1, str.Length - 1);
        }

        /// <summary>计算当前字符串与指定字符串的编辑距离(相似度)</summary>
        /// <param name="source">源字符串</param>
        /// <param name="target">目标字符串</param>
        /// <param name="similarity">输出相似度</param>
        /// <param name="ignoreCase">是否忽略大小写</param>
        /// <returns>编辑距离</returns>
        public static int LevenshteinDistance(
          this string source,
          string target,
          out double similarity,
          bool ignoreCase = false)
        {
            if (string.IsNullOrEmpty(source))
            {
                if (string.IsNullOrEmpty(target))
                {
                    similarity = 1.0;
                    return 0;
                }
                similarity = 0.0;
                return target.Length;
            }
            if (string.IsNullOrEmpty(target))
            {
                similarity = 0.0;
                return source.Length;
            }
            string str1;
            string str2;
            if (ignoreCase)
            {
                str1 = source;
                str2 = target;
            }
            else
            {
                str1 = source.ToLower();
                str2 = source.ToLower();
            }
            int length1 = str1.Length;
            int length2 = str2.Length;
            int[,] numArray = new int[length1 + 1, length2 + 1];
            for (int index = 0; index <= length1; ++index)
                numArray[index, 0] = index;
            for (int index = 1; index <= length2; ++index)
                numArray[0, index] = index;
            for (int index1 = 1; index1 <= length1; ++index1)
            {
                char ch = str1[index1 - 1];
                for (int index2 = 1; index2 <= length2; ++index2)
                    numArray[index1, index2] = (int)ch != (int)str2[index2 - 1] ? Math.Min(numArray[index1 - 1, index2 - 1], Math.Min(numArray[index1 - 1, index2], numArray[index1, index2 - 1])) + 1 : numArray[index1 - 1, index2 - 1];
            }
            int num = Math.Max(length1, length2);
            similarity = (double)(num - numArray[length1, length2]) / (double)num;
            return numArray[length1, length2];
        }

        /// <summary>
        /// 计算两个字符串的相似度，应用公式：相似度=kq*q/(kq*q+kr*r+ks*s)(kq&gt;0,kr&gt;=0,ka&gt;=0)
        /// 其中，q是字符串1和字符串2中都存在的单词的总数，s是字符串1中存在，字符串2中不存在的单词总数，r是字符串2中存在，字符串1中不存在的单词总数. kq,kr和ka分别是q,r,s的权重，根据实际的计算情况，我们设kq=2，kr=ks=1.
        /// </summary>
        /// <param name="source">源字符串</param>
        /// <param name="target">目标字符串</param>
        /// <param name="ignoreCase">是否忽略大小写</param>
        /// <returns>字符串相似度</returns>
        public static double GetSimilarityWith(this string source, string target, bool ignoreCase = false)
        {
            if (string.IsNullOrEmpty(source) && string.IsNullOrEmpty(target))
                return 1.0;
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target))
                return 0.0;
            char[] charArray1 = source.ToCharArray();
            char[] charArray2 = target.ToCharArray();
            int num1 = ((IEnumerable<char>)charArray1).Intersect<char>((IEnumerable<char>)charArray2).Count<char>();
            int num2 = charArray1.Length - num1;
            int num3 = charArray2.Length - num1;
            return 2.0 * (double)num1 / (2.0 * (double)num1 + 1.0 * (double)num3 + 1.0 * (double)num2);
        }

        /// <summary>若输入字符串为null，则返回String.Empty。</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>输入字符串或String.Empty</returns>
        public static string EmptyIfNull(this string s)
        {
            return s ?? string.Empty;
        }

        /// <summary>字符串比较忽略大小写</summary>
        /// <param name="source">源字符串</param>
        /// <param name="target">目标字符串</param>
        /// <returns></returns>
        public static bool EqualsIgnoreCase(this string source, string target)
        {
            return source.Equals(target, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>去除输入字符串两边的空白，若为null，则结果为string.Empty</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>去除两边空白的字符串或string.Empty</returns>
        public static string TrimOrEmpty(this string s)
        {
            return s.EmptyIfNull().Trim();
        }

        /// <summary>去除输入字符串两边的空白，若为null，则结果为null</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>去除两边空白的字符串或null</returns>
        public static string TrimOrNull(this string s)
        {
            return s?.Trim();
        }

        /// <summary>将输入字符串转换成Int32，转换不成功则产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>Int32值</returns>
        public static int ToInt32(this string s)
        {
            return Convert.ToInt32(s);
        }

        /// <summary>将输入字符串转换成Int32，转换不成功则返回提供的默认值，不产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <param name="defaultValue">提供的默认值</param>
        /// <returns>Int32值</returns>
        public static int ToInt32OrDefault(this string s, int defaultValue)
        {
            int result;
            if (!int.TryParse(s, out result))
                return defaultValue;
            return result;
        }

        /// <summary>将输入字符串转换成Int32，转换不成功返回0，不产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>Int32值</returns>
        public static int ToInt32OrDefault(this string s)
        {
            return s.ToInt32OrDefault(0);
        }

        /// <summary>将输入字符串转换成decimal，转换不成功则产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>decimal值</returns>
        public static Decimal ToDecimal(this string s)
        {
            return Convert.ToDecimal(s);
        }

        /// <summary>将输入字符串转换成decimal，转换不成功则返回提供的默认值，不产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <param name="defaultValue">提供的默认值</param>
        /// <returns>decimal值或提供的默认值</returns>
        public static Decimal ToDecimalOrDefault(this string s, Decimal defaultValue)
        {
            Decimal result;
            if (!Decimal.TryParse(s, out result))
                return defaultValue;
            return result;
        }

        /// <summary>将输入字符串转换成decimal，转换不成功则返回0，不产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>decimal值或0</returns>
        public static Decimal ToDecimalOrDefault(this string s)
        {
            return s.ToDecimalOrDefault(Decimal.Zero);
        }

        /// <summary>将输入字符串转换成DateTime，转换不成功则产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>DateTime值</returns>
        public static DateTime ToDateTime(this string s)
        {
            return Convert.ToDateTime(s);
        }

        /// <summary>将输入字符串转换成DateTime，转换不成功则产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <param name="format">字符串格式</param>
        /// <returns>DateTime值</returns>
        public static DateTime ToDateTime(this string s, string format)
        {
            return DateTime.ParseExact(s, format, (IFormatProvider)null);
        }

        /// <summary>将输入字符串转换成Int16，转换不成功则产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>Int16值</returns>
        public static short ToInt16(this string s)
        {
            return Convert.ToInt16(s);
        }

        /// <summary>将输入字符串转换成Int16，转换不成功则返回提供的默认值，不产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <param name="defaultValue">提供的默认值</param>
        /// <returns>Int16值或提供的默认值</returns>
        public static short ToInt16OrDefault(this string s, short defaultValue)
        {
            short result;
            if (!short.TryParse(s, out result))
                return defaultValue;
            return result;
        }

        /// <summary>将输入字符串转换成Int16，转换不成功则返回0，不产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>Int16值或0</returns>
        public static short ToInt16OrDefault(this string s)
        {
            return s.ToInt16OrDefault((short)0);
        }

        /// <summary>将输入字符串转换成Int64，转换不成功则产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>Int64值</returns>
        public static long ToInt64(this string s)
        {
            return Convert.ToInt64(s);
        }

        /// <summary>将输入字符串转换成Int64，转换不成功则返回提供的默认值，不产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <param name="defaultValue">提供的默认值</param>
        /// <returns>Int64值或提供的默认值</returns>
        public static long ToInt64OrDefault(this string s, long defaultValue)
        {
            long result;
            if (!long.TryParse(s, out result))
                return defaultValue;
            return result;
        }

        /// <summary>将输入字符串转换成Int64，转换不成功则返回0，不产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>Int64值或0</returns>
        public static long ToInt64OrDefault(this string s)
        {
            return s.ToInt64OrDefault(0L);
        }

        /// <summary>将输入字符串转换成byte，转换不成功则产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>byte值</returns>
        public static byte ToByte(this string s)
        {
            return Convert.ToByte(s);
        }

        /// <summary>将输入字符串转换成byte，转换不成功则返回提供的默认值，不产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <param name="defaultValue">提供的默认值</param>
        /// <returns>byte值或提供的默认值</returns>
        public static byte ToByteOrDefault(this string s, byte defaultValue)
        {
            byte result;
            if (!byte.TryParse(s, out result))
                return defaultValue;
            return result;
        }

        /// <summary>将输入字符串转换成byte，转换不成功则返回0，不产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>byte值或0</returns>
        public static byte ToByteOrDefault(this string s)
        {
            return s.ToByteOrDefault((byte)0);
        }

        /// <summary>将输入字符串转换成double，转换不成功则产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>double值</returns>
        public static double ToDouble(this string s)
        {
            return Convert.ToDouble(s);
        }

        /// <summary>将输入字符串转换成double，转换不成功则返回提供的默认值，不产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <param name="defaultValue">提供的默认值</param>
        /// <returns>double值或提供的默认值</returns>
        public static double ToDoubleOrDefault(this string s, double defaultValue)
        {
            double result;
            if (!double.TryParse(s, out result))
                return defaultValue;
            return result;
        }

        /// <summary>将输入字符串转换成double，转换不成功则返回0，不产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>double值或0</returns>
        public static double ToDoubleOrDefault(this string s)
        {
            return s.ToDoubleOrDefault(0.0);
        }

        /// <summary>将输入字符串转换成float，转换不成功则产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>float值</returns>
        public static float ToFloat(this string s)
        {
            return Convert.ToSingle(s);
        }

        /// <summary>将输入字符串转换成float，转换不成功则返回提供的默认值，不产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <param name="defaultValue">提供的默认值</param>
        /// <returns>float值或提供的默认值</returns>
        public static float ToFloatOrDefault(this string s, float defaultValue)
        {
            float result;
            if (!float.TryParse(s, out result))
                return defaultValue;
            return result;
        }

        /// <summary>将输入字符串转换成float，转换不成功则返回0，不产生异常</summary>
        /// <param name="s">输入字符串</param>
        /// <returns>float值或0</returns>
        public static float ToFloatOrDefault(this string s)
        {
            return s.ToFloatOrDefault(0.0f);
        }

        public static byte[] ToByteArray(this string hex)
        {
            return Enumerable.Range(0, hex.Length).Where<int>((Func<int, bool>)(x => x % 2 == 0)).Select<int, byte>((Func<int, byte>)(x => Convert.ToByte(hex.Substring(x, 2), 16))).ToArray<byte>();
        }

        public static string ToPhoneCardNoEncryption(this string str)
        {
            return PhoneCardNoEncryption.Encryption(str);
        }

        public static string ToPhoneCardNoDecrypt(this string str)
        {
            return PhoneCardNoEncryption.Decrypt(str);
        }
    }
}
