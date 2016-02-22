using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MtbGraph.GraphComponent
{
    [ClassInterface(ClassInterfaceType.None)]
    public class BarLineLegendBox : Graphcomp, ICOMInterop_BarLineLegend, ILegendBox
    {
        private Font fontLgnd;
        private Size sizeOfLgnd;
        private int incrPercent;
        public BarLineLegendBox()
        {
            Form f = new Form();
            float dpiX, dpiY;
            Graphics g = f.CreateGraphics();
            dpiX = g.DpiX;
            dpiY = g.DpiY;
            this.incrPercent = (dpiX == 96 ? 100 : (dpiX == 120 ? 125 : 150));
            this.fontLgnd = new Font(System.Drawing.SystemFonts.DialogFont.Name, (float)(8 * 100 / incrPercent), FontStyle.Regular);
            this.sizeOfLgnd = new Size(0, 0);
            Show = true;
            NotationType = LegendNotationType.Trend;
            this.GraphSize = new Size(576, 384);
            this.FontSize = 0f;
            this.VerticalBase = 0;
        }
        public Size GraphSize { set; get; }
        public Single FontSize { set; get; }
        public double VerticalBase { set; get; }
        public BarLineLegendBox Clone()
        {
            /*
             * 如同 Reference line, 做淺層複製就好，因為已設定變
             * 數(影響legend size) 不太會變，如果要變，複製後再輸入字串陣列就好。
             */
            return (BarLineLegendBox)this.MemberwiseClone();
        }

        public Size GetSize()
        {
            return sizeOfLgnd;
        }
        public bool Show { get; set; }
        public LegendboxLook LegendBoxLook { set; get; }
        public Location Location { set; get; }
        public bool HideHead { set; get; }
        /*
         * 變數是由 Clomun 的 Label 構成
         * 
         */
        public void SetVariables(ref String[] variables)
        {
            this.sizeOfLgnd = CalculateSize(variables);
        }

        //private String sectTitle = null;
        public String SectTitle { set; get; }

        /*
         * 計算legend box 需要的大小，不自動考慮 head，回傳 pixel。如果需
         * 要計算包含 head 的大小，直接把 head 納入其中
         * 
         * 使用時(通常是搭配修改位置)要和圖形大小(default 576*384) 搭配
         * 算出 Figure unit...要如何處理 Graph size 的變動，需要想想
         * 
         */
        public Size CalculateSize(String[] variables)
        {
            Size size = new Size(0, 0);
            Size tmp;
            if (this.FontSize != 0f)
            {
                this.fontLgnd = new Font(System.Drawing.SystemFonts.DialogFont.Name, (float)(this.FontSize *100 / incrPercent), FontStyle.Regular);
            }
            foreach (String str in variables)
            {
                tmp = TextRenderer.MeasureText(str, this.fontLgnd);
                if (tmp.Width > size.Width)
                {
                    size.Width = tmp.Width;
                }
                size.Height = size.Height + tmp.Height + 1;
            }

            switch (this.NotationType)
            {
                case LegendNotationType.Trend:
                    size.Width = size.Width + 33;//(30 標記 + 3 間距 )
                    size.Height = size.Height + 4;
                    break;
                case LegendNotationType.Bar:
                    size.Width = size.Width + 17;//(13 標記 + 4 間距 )
                    size.Height = size.Height;
                    break;
            }

            return size;
        }

        public String GetCommand()
        {
            return null;
        }


        public void SetDefault()
        {
            this.SectTitle = null;
            this.HideHead = false;
            this.Location = Location.Auto;
            this.LegendBoxLook = LegendboxLook.Normal;
            this.Show = true;
            this.sizeOfLgnd = new Size(0, 0);
        }

        public LegendNotationType NotationType { set; get; }
    }
}
