using System;
using System.IO;
using Microsoft.Extensions.Configuration;

namespace ExcelTools
{
    public static class Utils
    {
        public static string BasePath = AppDomain.CurrentDomain.BaseDirectory;

        public static IConfigurationRoot Config;
        public static bool IsRecordLog;

        static Utils()
        {
            var configurationBuilder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile("AppSettings.Json");
            Config = configurationBuilder.Build();
            IsRecordLog = Config["IsRecordLog"] == "1";
        }

        public static void Reload()
        {
            var configurationBuilder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile("AppSettings.Json");
            Config = configurationBuilder.Build();
            IsRecordLog = Config["IsRecordLog"] == "1";
        }

        /// <summary>
        /// 记录日志
        /// </summary>
        /// <param name="msg"></param>
        public static void LogInfo(string msg,bool timeFlag=true)
        {
            Console.WriteLine(msg);
            if(IsRecordLog)
                WriteLogLine(msg,timeFlag);    
        }

        /// <summary>
        /// 写日志
        /// </summary>
        /// <param name="msg"></param>
        public static void WriteLogLine(string msg,bool timeFlag=true)
        {
            string logName = DateTime.Now.ToString("yyyyMMdd");
            string logFilePath = Path.Combine(BasePath, "log", logName + ".txt");
            if (!File.Exists(logFilePath))
            {
                using (File.Create(logFilePath))
                {

                }
            }

            var content = (timeFlag ? DateTime.Now.ToString() : "") + " " + msg + "\r\n";

            File.AppendAllText(logFilePath, content);
        }
    }
}
