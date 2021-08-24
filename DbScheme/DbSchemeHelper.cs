using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using AutoMapper;
using Excel;
using ForExcelImport;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;

namespace ExcelTools
{
    public class DbSchemeHelper
    {
        public static Dictionary<string, Type> typeDic = new Dictionary<string, Type>();

        public static IMapper mapper = null;

        static DbSchemeHelper()
        {
            InitDic();
            MapperInit();
        }

        public static void Test()
        {
            var db = GetTestDbContext(typeof(tb_relative));

            var path = @"F:\1-HIS基础数据整理.xlsx";


            var sheets = MyMiniExcel.GetSheetNames(path);

            foreach (var sheet in sheets)
            {
                Console.WriteLine($"sheet name : {sheet} ");
                var strArray = sheet.Split('-').ToList();
                if (strArray.Count <= 1)
                {
                    Console.WriteLine($"sheet名 {sheet} 格式有误，请检查!");
                    //continue;
                    return;
                }

                var type = typeDic.GetValueOrDefault(strArray[0]);

                //var rows = MyMiniExcel.Query<tb_test>(path, sheetName: sheet);
                var rows2 = MyMiniExcel.Query(path, type, sheet);

                foreach (var row in rows2)
                {
                    Console.WriteLine(JsonConvert.SerializeObject(row));
                }

                //db.AddRange(rows2);
                //db.SaveChanges();
            }

            //var originProject = await _lighterDbContext.Projects.FirstOrDefaultAsync(p => p.Id == id, cancellationToken);
            //if (originProject == null)
            //{
            //    return NotFound();
            //}
            //var properties = _lighterDbContext.Entry(originProject).Properties.ToList();
            //foreach (var query in HttpContext.Request.Query)
            //{
            //    var property = properties.FirstOrDefault(p => p.Metadata.Name == query.Key);
            //    if (property == null)
            //        continue;

            //    var currentValue = Convert.ChangeType(query.Value.First(), property.Metadata.ClrType);
            //    _lighterDbContext.Entry(originProject).Property(query.Key).CurrentValue = currentValue;
            //    _lighterDbContext.Entry(originProject).Property(query.Key).IsModified = true;
            //}

            //await _lighterDbContext.SaveChangesAsync();

            //db.SaveChanges();
        }

        public static DbContext GetTestDbContext(Type type)
        {
            var configurationBuilder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile("AppSettings.Json");
            var configuration = configurationBuilder.Build();
            var dbAttr = type.GetCustomAttribute<DbAttribute>();
            string dbName = "cloudih";
            if (dbAttr != null)
                dbName = dbAttr.DbName;
            var dbContextOptionsBuilder = new DbContextOptionsBuilder<TestDbContext>();

            dbContextOptionsBuilder.UseMySql(configuration.GetConnectionString(dbName));

            return new TestDbContext(dbContextOptionsBuilder.Options);
        }

        /// <summary>
        /// 基础表 Type 字典初始化
        /// </summary>
        public static void InitDic()
        {
            typeDic.Add("tb_test", typeof(tb_test));//继承baseEntity接口 反射加入字段
            typeDic.Add("tb_test_dto", typeof(tb_relative));
        }

        /// <summary>
        /// automapper 映射，主要用于后续 关联表 和实体类的映射
        /// </summary>
        public static void MapperInit()
        {
            var config = new MapperConfiguration(cfg =>
            {
                //dto mapper (consider reflect)， Dto (继承接口反射) -> Entity
                cfg.CreateMap<tb_relative, tb_test>();
            });

            mapper = config.CreateMapper();

        }

    }
}
