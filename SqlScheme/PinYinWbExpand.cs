using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;

namespace ExcelTools.SqlScheme
{
    public static class PinYinWbExpand
    {
        public static IDictionary<char, char> mapPy = new Dictionary<char, char>();
        public static IDictionary<char, char> mapWb = new Dictionary<char, char>();
        static PinYinWbExpand()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var path = $"{assembly.GetName().Name}.pinyinwubi.xml";
            var xml = new XmlDocument();
            using (var stream = assembly.GetManifestResourceStream(path))
                xml.Load(stream);

            foreach (XmlNode xn in xml.GetElementsByTagName("PY"))
            {
                foreach (XmlNode xnn in xn.ChildNodes)
                {
                    for (var i = 0; i < xnn.InnerText.Length; i++)
                        mapPy[xnn.InnerText[i]] = xnn.Name[0];
                }
            }

            foreach (XmlNode xn in xml.GetElementsByTagName("WB"))
            {
                foreach (XmlNode xnn in xn.ChildNodes)
                {
                    for (var i = 0; i < xnn.InnerText.Length; i++)
                        mapWb[xnn.InnerText[i]] = xnn.Name[0];
                }
            }
        }

        /// <summary>
        /// 获取编码
        /// </summary>
        /// <param name="map"></param>
        /// <param name="txt"></param>
        /// <returns></returns>
        private static string GetCode(this string txt, IDictionary<char, char> map)
        {
            if (string.IsNullOrEmpty(txt))
                return txt;

            var dest = new StringBuilder();
            foreach (var c in txt)
            {
                if (c >= 'a' && c <= 'z')
                {
                    dest.Append(char.ToUpper(c));
                    continue;
                }
                if (c >= 'A' && c <= 'Z')
                {
                    dest.Append(char.ToUpper(c));
                    continue;
                }
                if (c >= '0' && c <= '9')
                {
                    dest.Append(char.ToUpper(c));
                    continue;
                }
                if (map.ContainsKey(c))
                {
                    dest.Append(map[c]);
                    continue;
                }
                dest.Append(c);
            }
            return dest.ToString();
        }

        /// <summary>
        /// 获取中文首字母拼音
        /// </summary>
        /// <param name="txt"></param>
        /// <returns></returns>
        public static string GetFirstPY(this string txt)
        {
            return txt.GetCode(mapPy);
        }

        /// <summary>
        /// 获取中文五笔编码
        /// </summary>
        /// <param name="txt"></param>
        /// <returns></returns>
        public static string GetFirstWB(this string txt)
        {
            return txt.GetCode(mapWb);
        }
    }
}
