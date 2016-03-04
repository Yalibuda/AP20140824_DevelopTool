using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MtbGraph.GraphComponent
{
    public class Title:INotation
    {
        public Title()
        {
            Text = null;
            Size = 13;
            Italic = false;
            Bold = false;
            Color = -1;
        }

        /// <summary>
        /// 設定 Title 的內容，預設為 null。如果不要顯示 Title 就使用""
        /// </summary>
        public string Text { set; get; }

        /// <summary>
        /// 設定 Title 的尺寸
        /// </summary>
        public float Size { set; get; }

        /// <summary>
        /// 設定 Title 是否要斜體
        /// </summary>
        public bool Italic { set; get; }

        /// <summary>
        /// 設定 Title 是否粗體
        /// </summary>
        public bool Bold { set; get; }

        /// <summary>
        /// 設定 Title 顏色
        /// </summary>
        public int Color { set; get; }

        /// <summary>
        /// 複製 Title 內所有屬性
        /// </summary>
        /// <returns></returns>
        public INotation Clone()
        {
            return this.MemberwiseClone() as INotation;
        }
    }
}
