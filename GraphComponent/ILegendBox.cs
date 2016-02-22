using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    public interface ILegendBox
    {
        ///<summary> 方法說明
        /// 2015/05/12
        /// VerticalBase 用於調整Bar-line plot legend box 的高度..其他legendbox 暫時沒作用 
        /// </summary>

        Size GetSize();
        Single FontSize { set; get; }
        double VerticalBase { set; get; }
        bool Show { set; get; }
        LegendboxLook LegendBoxLook { set; get; }
        LegendNotationType NotationType { set; get; }
        Location Location { set; get; }
        String SectTitle { set; get; }
        bool HideHead { set; get; }
        void SetVariables(ref String[] variables);
        Size CalculateSize(String[] variables);
        String GetCommand();
        void SetDefault();


    }
}
