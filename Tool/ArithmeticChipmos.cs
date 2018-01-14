using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Tool
{
    /// <summary>
    /// 南茂常用的數學計算公式
    /// </summary>
    public static class ArithmeticChipmos
    {
        private const double MISSINGVALUE = 1.23456E+30;

        /// <summary>
        /// 計算 Yield (pass die/ total die *100)
        /// </summary>
        /// <param name="obj">包含 Pass die 和 Total die 的陣列</param>
        /// <param name="dec">要顯示的小數位數</param>
        /// <returns></returns>
        public static double Yiled(object obj, int? dec = 2)
        {
            IEnumerable<double> passDie = null;
            IEnumerable<double> ttlDie = null;
            if (obj is IEnumerable<IEnumerable<double>>)
            {
                List<IEnumerable<double>> enumerable = (obj as IEnumerable<IEnumerable<double>>).ToList();
                passDie = enumerable.Select(x => (x as double[])[0]);
                ttlDie = enumerable.Select(x => (x as double[])[1]);
            }
            else
            {
                throw new ArgumentException("輸入的變數的型別必須為IEnumerable<IEnumerable<double>>");
            }


            if (passDie == null || ttlDie == null) throw new ArgumentNullException("輸入物件不可為NULL");
            if (passDie.Count() != ttlDie.Count()) throw new Exception("計算Yield的數據長度不同");
            if (passDie.Count() == 0) throw new Exception("計算 Yield 的長度不可為0");
            double result = passDie.Where(x => x < MISSINGVALUE).Sum() / ttlDie.Where(x => x < MISSINGVALUE).Sum() * 100;
            if (dec != null) result = Math.Round(result, (int)dec);
            return result;
        }

        /// <summary>
        /// 計算PPM (defect/total die * 10^6)
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="dec"></param>
        /// <returns></returns>
        public static double PPM(object obj, int? dec = 0)
        {
            IEnumerable<double> defect = null;
            IEnumerable<double> ttlDie = null;
            if (obj is IEnumerable<IEnumerable<double>>)
            {
                List<IEnumerable<double>> enumerable = (obj as IEnumerable<IEnumerable<double>>).ToList();
                defect = enumerable.Select(x => (x as double[])[0]);
                ttlDie = enumerable.Select(x => (x as double[])[1]);
            }
            else
            {
                throw new ArgumentException("輸入的變數的型別必須為IEnumerable<IEnumerable<double>>");
            }


            if (defect == null || ttlDie == null) throw new ArgumentNullException("輸入物件不可為NULL");
            if (defect.Count() != ttlDie.Count()) throw new Exception("計算Yield的數據長度不同");
            if (defect.Count() == 0) throw new Exception("計算 Yield 的長度不可為0");
            double result = defect.Where(x => x < MISSINGVALUE).Sum() / ttlDie.Where(x => x < MISSINGVALUE).Sum() * Math.Pow(10, 6);
            if (dec != null) result = Math.Round(result, (int)dec);
            return result;
        }

        /// <summary>
        /// 處理
        /// </summary>
        /// <param name="colsToAggregate"></param>
        /// <param name="fun"></param>
        /// <param name="colGroups"></param>
        /// <param name="datatable"></param>
        /// <returns></returns>
        public static DataTable Apply(string[] colsToAggregate,
            Func<object, int?, double> fun,
            string[] colGroups,
            DataTable datatable)
        {
            //取得包含計算欄位的 column index
            int[] indexOfColsToAggre = datatable.Columns.Cast<DataColumn>()
                .Select((x, i) => new { Col = x, Index = i })
                .Where(x => colsToAggregate.Contains(x.Col.ColumnName)).Select(x => x.Index).ToArray();

            var groupData = datatable.AsEnumerable().GroupBy(
                    r => new NTuple<object>(from col in colGroups select r[col]));

            var ttt = groupData.Select(g => g.Select(r => r.ItemArray.TakeWhile((value, i) => indexOfColsToAggre.Contains(i)).ToArray()).ToArray()).ToArray();
            foreach (var item in ttt)
            {
                Console.WriteLine(item);
            }
            var applyData = groupData.Select(g =>
                new
                {
                    Group = g.Key.Values,
                    Value = fun(g.Select(r => r.ItemArray.TakeWhile((value, i) => indexOfColsToAggre.Contains(i)).Select(x => Convert.ToDouble(x)).ToArray()), null)
                    //Value = fun(g.Select(r => r[colAggregate].GetType() == typeof(double) ? r.Field<double>(colAggregate) : Convert.ToDouble(r.Field<decimal>(colAggregate))).ToArray())
                }).ToArray();

            //var result = applyData.ToDictionary(x => x.Group, x => x.Value);
            DataTable dt = new DataTable();
            var keyValues = applyData.Select(x => x.Group).FirstOrDefault();
            for (int i = 0; i < keyValues.Count(); i++)
            {
                Type t = keyValues[i].GetType();
                DataColumn dc = new DataColumn("ByVar" + (i + 1), t);
                dt.Columns.Add(dc);
            }
            dt.Columns.Add(new DataColumn("Value", applyData.Select(x => x.Value).FirstOrDefault().GetType()));

            foreach (var item in applyData)
            {
                object[] o = new object[item.Group.Count() + 1];
                for (int i = 0; i < item.Group.Count(); i++)
                {
                    o[i] = item.Group[i];
                }
                o[o.Length - 1] = item.Value;
                dt.Rows.Add(o);
            }
            return dt;
        }

        /// <summary>
        /// 把 IEnumerable<T> 轉成 DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="items"></param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(this IEnumerable<T> items)
        {
            var tb = new DataTable(typeof(T).Name);

            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (var prop in props)
            {
                tb.Columns.Add(prop.Name, prop.PropertyType);
            }

            foreach (var item in items)
            {
                var values = new object[props.Length];
                for (var i = 0; i < props.Length; i++)
                {
                    values[i] = props[i].GetValue(item, null);
                }

                tb.Rows.Add(values);
            }

            return tb;
        }

        ///<summary>Finds the index of the first item matching an expression in an enumerable.</summary>
        ///<param name="items">The enumerable to search.</param>
        ///<param name="predicate">The expression to test the items against.</param>
        ///<returns>The index of the first matching item, or -1 if no items match.</returns>
        public static int FindIndex<T>(this IEnumerable<T> items, Func<T, bool> predicate)
        {
            if (items == null) throw new ArgumentNullException("items");
            if (predicate == null) throw new ArgumentNullException("predicate");

            int retVal = 0;
            foreach (var item in items)
            {
                if (predicate(item)) return retVal;
                retVal++;
            }
            return -1;
        }
        ///<summary>Finds the index of the first occurrence of an item in an enumerable.</summary>
        ///<param name="items">The enumerable to search.</param>
        ///<param name="item">The item to find.</param>
        ///<returns>The index of the first matching item, or -1 if the item was not found.</returns>
        public static int IndexOf<T>(this IEnumerable<T> items, T item) { return items.FindIndex(i => EqualityComparer<T>.Default.Equals(item, i)); }
    }
}
