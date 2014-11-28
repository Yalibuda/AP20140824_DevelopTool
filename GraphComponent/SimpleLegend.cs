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
    public class SimpleLegend : Graphcomp, ICOMInterop_Legend, ILegendBox
    {
        private Font fontLgnd;
        private Size sizeOfLgnd;
        public SimpleLegend()
        {
            Form f = new Form();
            float dpiX, dpiY;
            Graphics g = f.CreateGraphics();
            dpiX = g.DpiX;
            dpiY = g.DpiY;
            int incrPercent = (dpiX == 96 ? 100 : (dpiX == 120 ? 125 : 150));
            this.fontLgnd = new Font(System.Drawing.SystemFonts.DialogFont.Name, (float)(8 * 100 / incrPercent), FontStyle.Regular);
            this.sizeOfLgnd = new Size(0, 0);
            Show = true;
            NotationType = LegendNotationType.Trend;
        }

        public SimpleLegend Clone()
        {
            /*
             * 如同 Reference line, 做淺層複製就好，因為已設定變
             * 數(影響legend size) 不太會變，如果要變，複製後再輸入字串陣列就好。
             */
            return (SimpleLegend)this.MemberwiseClone();
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

            foreach (String str in variables)
            {
                tmp = TextRenderer.MeasureText(str, this.fontLgnd);
                if (tmp.Width > size.Width)
                {
                    size.Width = tmp.Width;
                }
                size.Height = size.Height + tmp.Height;
            }

            switch (this.NotationType)
            {
                case LegendNotationType.Trend:
                    size.Width = size.Width + 33;//(30 標記 + 3 間距 )
                    size.Height = size.Height + 4;
                    break;
                case LegendNotationType.Bar:
                    size.Width = size.Width + 17;//(13 標記 + 4 間距 )
                    size.Height = size.Height + 8;
                    break;
            }

            return size;
        }

        public String GetCommand()
        {
            StringBuilder cmnd = new StringBuilder();
            if (Show)
            {
                String coordinate = "";
                double xmin, ymin, xmax, ymax;

                /* 
                 * 如果要顯示 Section title，先計算 title 的 size()，此類別以處理
                 * Multi-column input 狀態的簡單 legend box，所以通常其名稱為 Variable
                 */
                Size tmp = new Size(0, 0);
                if (!HideHead)
                {
                    tmp = TextRenderer.MeasureText((this.SectTitle == null ? "Variable" : this.SectTitle), this.fontLgnd);
                }
                /*
                 * 判斷是否要自訂 Legend 位置
                 */
                if (sizeOfLgnd.IsEmpty)
                {
                    cmnd.AppendLine("  #Legend box 尺寸無法計算(未執行 Legend 的 SetVariable方法)");
                    coordinate = String.Empty;
                }
                else
                {
                    switch (this.Location)
                    {
                        case Location.RightTop:
                            xmax = 0.998;
                            ymax = 0.998;
                            dynamic aa = (double)tmp.Width / 576;
                            xmin = Math.Max(xmax - Math.Max((double)sizeOfLgnd.Width / 576, (double)tmp.Width / 576), 0);
                            ymin = Math.Max(ymax - ((double)(sizeOfLgnd.Height + tmp.Height) / 384), 0);
                            coordinate = " " + xmin + " " + xmax + " " + ymin + " " + ymax;
                            break;
                        case Location.LeftDown:
                            xmin = 0.002;
                            ymin = 0.001;
                            xmax = Math.Min(xmin + Math.Max((double)sizeOfLgnd.Width / 576, (double)tmp.Width / 576), 1);
                            ymax = Math.Min(ymin + (double)(sizeOfLgnd.Height + tmp.Height / 384), 1);
                            coordinate = " " + xmin + " " + xmax + " " + ymin + " " + ymax;
                            break;
                    }
                }

                cmnd.AppendLine(" LEGE" + coordinate + ";");
                if (this.LegendBoxLook == LegendboxLook.Transparent) cmnd.AppendLine("  TYPE 0;" + Environment.NewLine + "ETYPE 0;");
                cmnd.AppendLine("  SECT 1;");
                if (HideHead)
                {
                    cmnd.AppendLine("   CHHIDE;");
                }
                else if (!String.IsNullOrEmpty(this.SectTitle))
                {
                    cmnd.AppendLine("   CHEAD 2 \"" + this.SectTitle + "\";");
                }
            }
            else
            {
                cmnd.AppendLine(" NOLEGEND;");
            }
            return cmnd.ToString();
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
