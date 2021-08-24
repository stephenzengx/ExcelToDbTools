using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTools
{
    public class CustomeUnit
    {
        public List<string> Units { get; private set; } = new List<string>();

        public List<double> Rates { get; private set; } = new List<double>();

        /// <summary>自定义单位转换</summary>
        /// <param name="units">cm/dm/m</param>
        /// <param name="rates">1/10/100</param>
        /// <param name="splitChar">分割符</param>
        public CustomeUnit(string units, string rates, char splitChar = '/')
        {

            //var customeUnit = new CustomeUnit($"{drug.ZXDW}/{drug.JLDW}/{drug.BZDW}", $"1/{drug.JLSL}/{drug.BZSL}");
            foreach (string str in units.Split(splitChar, StringSplitOptions.None))
            {
                if (this.Units.IndexOf(str) == -1)
                    this.Units.Add(str);
            }
            foreach (string str in rates.Split(splitChar, StringSplitOptions.None))
                this.Rates.Add(Convert.ToDouble(str));
            if (this.Units.Count != this.Rates.Count)
                throw new Exception("单位与转换关系对应关系错误");
        }

        /// <summary>获取当前单位的下一个单位</summary>
        /// <param name="unit"></param>
        /// <returns></returns>
        public string GetNextUnit(string unit)
        {
            int num = this.Units.IndexOf(unit);
            if (num == -1)
                throw new Exception("[" + unit + "] is not found");
            if (num >= this.Units.Count - 1)
                return unit;
            return this.Units[num + 1];
        }

        /// <summary>单位转换</summary>
        /// <param name="num">数量</param>
        /// <param name="unit">原单位</param>
        /// <param name="toUnit">目标单位</param>
        /// <returns></returns>
        public double UnitConversion(double num, string unit, string toUnit)
        {
            int num1 = this.Units.IndexOf(unit);
            int num2 = this.Units.IndexOf(toUnit);
            if (num1 == -1 || num2 == -1)
                throw new Exception("not convert " + unit + " => " + toUnit + " unitTmp:" + string.Join<string>('/', (IEnumerable<string>)this.Units));
            double num3 = 1.0;
            for (int index = 0; index <= num1; ++index)
                num3 *= this.Rates[index];
            double num4 = 1.0;
            for (int index = 0; index <= num2; ++index)
                num4 *= this.Rates[index];
            if (num4 <= num3)
                return num * (num3 / num4);
            return num / (num4 / num3);
        }
    }
}
