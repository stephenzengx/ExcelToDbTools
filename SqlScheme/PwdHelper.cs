using System;

namespace ExcelTools.SqlScheme
{
    public static class PwdHelper
    {
        public static string Str(int length, bool sleep)
        {
            if (sleep) System.Threading.Thread.Sleep(3);
            char[] Pattern = new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'g', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z' };
            string result = "";
            int n = Pattern.Length;
            System.Random random = new Random(~unchecked((int)DateTime.Now.ToCstTime().Ticks));
            for (int i = 0; i < length; i++)
            {
                int rnd = random.Next(0, n);
                result += Pattern[rnd];
            }
            return result;
        }

        /// <summary>
        /// 装换为本地时间
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public static DateTime ToCstTime(this DateTime dateTime)
        {
            var ntcTime = TimeZoneInfo.ConvertTimeToUtc(dateTime);
            return ntcTime.ToLocalTime();
        }

        public static string ToPassWord(string password, string pwdSalt)
        {
            return MD5Encryption.Encrypt($"{pwdSalt}{password}");
        }

        public static string ToPassWordGetSalt(string password, out string pwdSalt)
        {
            pwdSalt = Str(10, true);
            return ToPassWord(password, pwdSalt);
        }
    }
}
