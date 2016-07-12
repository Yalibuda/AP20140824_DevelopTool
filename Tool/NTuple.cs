using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Tool
{
    public class NTuple<T> : IEquatable<NTuple<T>>
    {
        /// <summary>
        /// <para>
        /// 將多個物件包起來，能做為 linq 或 lambda 動態 grouping 的依據
        /// E.g. 
        /// </para>
        /// <para>IEnumerable&lt;string&gt; columnsToGroupBy = ...;</para>
        /// <para>
        /// var groups = dt.AsEnumerable().GroupBy(r =&gt; new NTuple&lt;object&gt;(from column in columnsToGroupBy select r[column]));
        /// </para>        
        /// <pare>
        /// ref: http://stackoverflow.com/questions/26658978/c-sharp-linq-how-to-build-group-by-clause-dynamically
        /// </pare>
        /// </summary>
        /// <param name="values"></param>
        public NTuple(IEnumerable<T> values)
        {
            Values = values.ToArray();
        }

        public readonly T[] Values;

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj))
                return true;
            if (obj == null)
                return false;
            return Equals(obj as NTuple<T>);
        }

        public bool Equals(NTuple<T> other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other == null)
                return false;
            var length = Values.Length;
            if (length != other.Values.Length)
                return false;
            for (var i = 0; i < length; ++i)
                if (!Equals(Values[i], other.Values[i]))
                    return false;
            return true;
        }

        public override int GetHashCode()
        {
            var hc = 17;
            foreach (var value in Values)
                hc = hc * 37 + (!ReferenceEquals(value, null) ? value.GetHashCode() : 0);
            return hc;
        }
    }
}
