using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.HLBarLinePlot
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class DatalabOption : IDatalabOption
    {
        //設定是否使用系統顯示
        public bool AutoDecimal
        {
            get { return _auto; }
            set { _auto = value; }
        }
        bool _auto = true;
        /// <summary>
        /// 指定顯示的小數位數，預設為3。
        /// </summary>
        /// <exception cref="ArgumentException"></exception>
        public int DecimalPlace
        {
            get
            {
                return _decPlace;
            }
            set
            {
                if (value < 0) throw new ArgumentException("指定的小數位數必須大於等於0");
                _decPlace = value;
            }
        }
        int _decPlace = 3;
    }
}
