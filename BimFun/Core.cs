using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BimFun
{
    public static class Core
    {
        /// <summary>
        /// 一个类似于循环链的数组
        /// </summary>
        /// <param name="numbers">要处理的数组</param>
        /// <param name="count">个数</param>
        /// <returns></returns>
        public static List<List<double>> ChainNumber(List<double> numbers, int count = 2)
        {
            List<List<double>> numsList = new List<List<double>>();
            for (int i = 0; i < numbers.Count; i++)
            {
                if (i == numbers.Count - count + 1) break;
                if (count - numbers.Count >= 2) throw new ArgumentOutOfRangeException();
                List<double> nums = new List<double>();
                for (int j = 0; j < count; j++)
                {
                    nums.Add(numbers[i + j]);
                }
                numsList.Add(nums);
            }
            return numsList;
        }
        public static string NumberToString(double number) => number.ToString();
    }
}
