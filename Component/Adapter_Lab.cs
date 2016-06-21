using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component
{

    public class Adapter_Lab: ILabel
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
    }
}
