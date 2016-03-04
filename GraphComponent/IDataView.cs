using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    public interface IDataView
    {
        void SetType(dynamic type);
        int[] GetTypes();
        /// <summary>
        /// 輸入顏色的設定值，輸入int[]
        /// </summary>
        /// <param name="color"></param>
        void SetColor(dynamic color);
        int[] GetColor();
        void SetSize(dynamic size);
        int[] GetSize();
        void SetDefault();
        String GetCommand();
    }
}
