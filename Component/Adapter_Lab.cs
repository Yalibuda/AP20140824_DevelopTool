using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component
{

    internal class Adapter_Lab : ILabel
    {
        Mtblib.Graph.Component.Label _lab;
        /// <summary>
        /// 將元件轉為使用者可使用的介面
        /// </summary>
        /// <param name="lab"></param>
        public Adapter_Lab(Mtblib.Graph.Component.Label lab)
        {
            _lab = lab;
        }

        public bool Visible
        {
            get { return _lab.Visible; }
            set { _lab.Visible = value; }
        }
        /// <summary>
        /// 設定或取得 Label 的內容
        /// </summary>
        public string Text
        {
            get
            {
                return _lab.Text;
            }
            set
            {
                _lab.Text = value;
            }
        }

        /// <summary>
        /// 設定或取得 Label 的字體大小
        /// </summary>
        public float FontSize
        {
            get
            {
                return _lab.FontSize;
            }
            set
            {
                _lab.FontSize = value;
            }
        }

        /// <summary>
        /// 設定或取得 Label 的字體大小
        /// </summary>
        public int FontColor
        {
            get
            {
                return _lab.FontColor;
            }
            set
            {
                _lab.FontColor = value;
            }
        }

        /// <summary>
        /// 設定 Label 的位移的距離，合法的輸入是一個數值陣列(double[2])。
        /// 第一個參數是水平位移、第二的是垂直位移(Figure unit)
        /// </summary>
        public void OffSet(double hOffSet, double vOffSet)
        {
            if (hOffSet >= Mtblib.Tools.MtbTools.MISSINGVALUE ||
                vOffSet >= Mtblib.Tools.MtbTools.MISSINGVALUE)
            {
                _lab.Offset = null;
                return;
            }
            else if (hOffSet > 1 || hOffSet < -1 || vOffSet > 1 || vOffSet < -1)
            {
                throw new ArgumentException("OffSet 的參數必須為介於-1~1之間的數值");
            }
            _lab.Offset = new double[] { hOffSet, vOffSet };
        }
    }
}
