using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.IO;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Forms;

namespace MtbGraph
{
    public enum ScaleTickAttribute
    {
        None,
        NumberOfMajorTick,
        IntervalBetweenTick
    }
    public class MtbGraphFrame: IDisposable
    {
        protected int[] dFillColor = { 127, 28, 7, 58, 116, 78, 29, 45, 123, 35, 73, 8, 49, 57, 26 };
        protected int[] dLineColor = { 64, 8, 9, 12, 18, 34 };
        protected int[] dSymbType = { 6, 12, 16, 20, 23, 26, 29 };
        protected int[] dLineType = { 1, 2, 3, 4, 5 };
        protected String xLabel = null;
        protected String yLabel = null;
        protected String secsLabel = null;
        protected String title = null;
        private static String dXLabel = null;
        private static String dYLabel = null;
        private static String dSecLabel = null;
        private static String dTitle = null;
        protected double yScaleMin = 1.23456E+30;
        protected double yScaleMax = 1.23456E+30;
        protected double secScaleMin = 1.23456E+30;
        protected double secScaleMax = 1.23456E+30;
        protected double xLabelAngle = 1.23456E+30;
        protected bool gSave;
        protected String gPath;
        protected bool expCmnd;
        protected String cmndPath;
        protected bool copyToClip;
        protected int gWidth = 576;
        protected int gHeight = 384;
        protected int d_gWidth = 576;
        protected int d_gHeight = 384;
        protected String dlgndFontName = "Segoe UI";
        protected String whereCond;

        //protected Font d_TitleFont = new Font("Segoe UI Semibold", (float)9.5, FontStyle.Bold);
        //protected Font d_TickFont = new Font("Segoe UI Semibold", 8, FontStyle.Regular);
        //protected Font d_LabFont = new Font("Segoe UI Semibold", 10, FontStyle.Regular);
        //protected Font d_lgndFont = new Font("Segoe UI Semibold", (float)7, FontStyle.Regular);
        //protected Font d_TitleFont = new Font(System.Drawing.SystemFonts.DialogFont.Name, (float)9.5, FontStyle.Bold);
        //protected Font d_TickFont = new Font(System.Drawing.SystemFonts.DialogFont.Name, 8, FontStyle.Regular);
        //protected Font d_LabFont = new Font(System.Drawing.SystemFonts.DialogFont.Name, 9, FontStyle.Bold);
        //protected Font d_lgndFont = new Font(System.Drawing.SystemFonts.DialogFont.Name, 8, FontStyle.Regular);
        protected Font d_TitleFont;
        protected Font d_TickFont;
        protected Font d_LabFont;
        protected Font d_lgndFont;

        protected double[] xRefValue = null;
        protected int[] xRefType = null;
        protected int[] xRefColor = null;
        protected double[] yRefValue = null;
        protected int[] yRefType = null;
        protected int[] yRefColor = null;
        protected double[] secsRefValue = null;
        protected int[] secsRefType = null;
        protected int[] secsRefColor = null;

        protected ScaleTickAttribute yTickAttr = ScaleTickAttribute.None;
        protected double yTickAttrValue = 1.23456E+30;
        protected ScaleTickAttribute secTickAttr = ScaleTickAttribute.None;
        protected double secTickAttrValue = 1.23456E+30;
        public MtbGraphFrame()
        {
            Form f = new Form();
            float dpiX, dpiY;
            Graphics g = f.CreateGraphics();
            dpiX = g.DpiX;
            dpiY = g.DpiY;
            int incrPercent = (dpiX == 96 ? 100 : (dpiX == 120 ? 125 : 150));
            d_TitleFont = new Font(System.Drawing.SystemFonts.DialogFont.Name, (float)(9.5 * 100 / incrPercent), FontStyle.Bold);
            d_TickFont = new Font(System.Drawing.SystemFonts.DialogFont.Name, (float)(8 * 100 / incrPercent), FontStyle.Regular);
            d_LabFont = new Font(System.Drawing.SystemFonts.DialogFont.Name, (float)(9 * 100 / incrPercent), FontStyle.Bold);
            d_lgndFont = new Font(System.Drawing.SystemFonts.DialogFont.Name, (float)(8 * 100 / incrPercent), FontStyle.Regular);

        }
        protected virtual void CopyToClipboard(String cmndStr, Mtb.Project proj, Mtb.Worksheet ws, int start, int end)
        {
            Mtb.Commands cmnds;
            cmnds = proj.Commands;

            bool flag = false;
            for (int i = start; i <= end; i++)
            {
                if (cmnds.Item(i).Name == cmndStr)
                {
                    foreach (Mtb.Output output in cmnds.Item(i).Outputs)
                    {
                        if (output.OutputType == Mtb.MtbOutputTypes.OTGraph)
                        {
                            output.Graph.CopyToClipboard();
                            flag = true;
                            break;
                        }
                    }
                    if (flag) break;
                }
            }

        }
        public void SetExportCommand(bool b, String path)
        {
            expCmnd = b;
            if (b)
            {
                if (String.IsNullOrEmpty(path))
                {
                    expCmnd = false;
                }
                else
                {
                    Regex regEx = new Regex(@"^(?:[\w]\:|\\)(\\[^\\\/:\*?<>|]+)+\.(txt|mtb)$", RegexOptions.IgnoreCase);
                    cmndPath = path;
                    Match match = regEx.Match(cmndPath);//check is a validate file extension
                    if (match.Success)
                    {
                        regEx = new Regex(@"^(?:[\w]\:|\\)(\\[^\\\/:\*?<>|]+)*\\", RegexOptions.IgnoreCase);
                        String rPath = regEx.Match(path).Groups[0].Value.ToString();
                        if (!Directory.Exists(rPath))
                        {
                            Directory.CreateDirectory(rPath);
                        }
                    }
                    else
                    {
                        gSave = false;
                    }
                }
            }

        }
        protected void ExportCommand(String str, String path, bool isNew = true)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(str);
            FileStream fs;
            if (isNew)
            {
                fs = new FileStream(cmndPath, FileMode.Create);
            }
            else
            {
                fs = new FileStream(cmndPath, FileMode.OpenOrCreate);
            }
            fs.Close();
            StreamWriter sw;
            if (isNew)
            {
                sw = new StreamWriter(cmndPath);
            }
            else
            {
                sw = new StreamWriter(cmndPath, true);
            }
            sw.Write(sb.ToString());
            sw.Close();
        }
        public void SetCopyToClipboard(bool b)
        {
            copyToClip = b;
        }
        public void SaveGraph(bool b, String path)
        {
            gSave = b;
            if (b)
            {
                if (String.IsNullOrEmpty(path))
                {
                    gSave = false;
                }
                else
                {
                    Regex regEx = new Regex(@"^(?:[\w]\:|\\)(\\[^\\\/:\*?<>|]+)+\.(JPEG|JPG)$", RegexOptions.IgnoreCase);
                    gPath = path;
                    Match match = regEx.Match(gPath);//check is a validate file extension
                    if (match.Success)
                    {
                        regEx = new Regex(@"^(?:[\w]\:|\\)(\\[^\\\/:\*?<>|]+)*\\", RegexOptions.IgnoreCase);
                        String rPath = regEx.Match(path).Groups[0].Value.ToString();
                        if (!Directory.Exists(rPath))
                        {
                            Directory.CreateDirectory(rPath);
                        }
                    }
                    else
                    {
                        gSave = false;
                    }
                }
            }

        }
        public void SetGraphSize(int width, int height)
        {
            this.gWidth = width;
            this.gHeight = height;
        }
        public void SetGraphTitle(String title)
        {
            this.title = title;
        }
        public void SetSecsAxlabel(String label)
        {
            this.secsLabel = label;
        }

        public void SetSecScaleMin(double d)
        {
            if (this.secScaleMax < 1.23456E+30 & d > this.secScaleMax)
            {
                MessageBox.Show("Invalid input value of secondary-scale minimum");
                return;
            }
            secScaleMin = d;
        }
        public void SetSecScaleMax(double d)
        {
            if (this.secScaleMin < 1.23456E+30 & d < this.secScaleMin)
            {
                MessageBox.Show("Invalid input value of secondary-scale maximum");
                return;
            }
            secScaleMax = d;
        }
        public void SetXAxlabel(String label)
        {
            this.xLabel = label;
        }
        public void SetXAxlabelAngle(double d)
        {
            this.xLabelAngle = d;
        }

        public void SetYAxlabel(String label)
        {
            this.yLabel = label;
        }
        public void SetDefaultXAxlabel()
        {
            this.xLabel = dXLabel;
        }
        public void SetDefaultXAxlabelAngle()
        {
            this.xLabelAngle = 1.23456E+30;
        }
        public void SetDefaultYAxlabel()
        {
            this.yLabel = dYLabel;
        }
        public void SetDefaultSecAxlabel()
        {
            this.secsLabel = dSecLabel;
        }
        public void SetDefaultTitle()
        {
            this.title = dTitle;
        }
        public void SetYScaleMin(double d)
        {
            if (this.yScaleMax < 1.23456E+30 & d > this.yScaleMax)
            {
                MessageBox.Show("Invalid input value of y-scale minimum");
                return;
            }
            this.yScaleMin = d;
        }
        public void SetYScaleMax(double d)
        {
            if (this.yScaleMin < 1.23456E+30 & d < this.yScaleMin)
            {
                MessageBox.Show("Invalid input value of y-scale maximum");
                return;
            }
            this.yScaleMax = d;
        }
        public void SetDefaultYScale()
        {
            this.yScaleMax = 1.23456E+30;
            this.yScaleMin = 1.23456E+30;
        }
        public void SetDefaultSecScale()
        {
            this.secScaleMax = 1.23456E+30;
            this.secScaleMin = 1.23456E+30;
        }

        public void SetXRefValue(ref double[] values, ref int[] types, ref int[] colors)
        {
            this.xRefValue = values;
            this.xRefType = types;
            this.xRefColor = colors;
        }
        public void ClearXRefValue()
        {
            this.xRefValue = null;
            this.xRefType = null;
            this.xRefColor = null;
        }
        public void SetYRefValue(ref double[] values, ref int[] types, ref int[] colors)
        {
            this.yRefValue = values;
            this.yRefType = types;
            this.yRefColor = colors;
        }
        public void ClearYRefValue()
        {
            this.yRefValue = null;
            this.yRefType = null;
            this.yRefColor = null;
        }

        public void SetSecsRefValue(ref double[] values, ref int[] types, ref int[] colors)
        {
            this.secsRefValue = values;
            this.secsRefType = types;
            this.secsRefColor = colors;
        }
        public void ClearSecsRefValue()
        {
            this.secsRefValue = null;
            this.secsRefType = null;
            this.secsRefColor = null;
        }

        public void SetYScaleTick(ScaleTickAttribute attr, double val)
        {
            this.yTickAttr = attr;
            this.yTickAttrValue = val;
        }
        public void SetSecScaleTick(ScaleTickAttribute attr, double val)
        {
            this.secTickAttr = attr;
            this.secTickAttrValue = val;
        }
        public void ResetScaleTick()
        {
            this.yTickAttr = ScaleTickAttribute.None;
            this.yTickAttrValue = 1.23456E+30;
            this.secTickAttr = ScaleTickAttribute.None;
            this.secTickAttrValue = 1.23456E+30;
        }


        public virtual void Dispose()
        {
            GC.Collect();
        }
        /*
         * Additional function
         */
        protected Size GetStringSize(Mtb.Worksheet ws, String col, Font font)
        {
            Size s = new Size(0, 0);
            String[] strArr = null;
            switch (ws.Columns.Item(col).DataType)
            {
                case Mtb.MtbDataTypes.Text:
                    strArr = ws.Columns.Item(col).GetData();
                    break;
                default:
                    break;
            }
            try
            {
                foreach (String str in strArr)
                {
                    Size tmp = TextRenderer.MeasureText(str, font);
                    if (tmp.Width > s.Width)
                    {
                        s.Width = tmp.Width;
                    }
                    if (tmp.Height > s.Height)
                    {
                        s.Height = tmp.Height;
                    }
                }
            }
            catch
            {
                return s;
            }

            return s;

        }
        protected String GetRefCmndString(double[] refValue, int[] refType = null, int[] refColor = null, int scale = 2, int primary = 1)
        {
            StringBuilder sb = new StringBuilder();
            bool[] refCheck = CheckRefCmndType(refValue, refType, refColor);

            if (refCheck[0])//是否為合格的輸入?
            {
                if (refCheck[1])//是否為單一指令可完成的參考線? 例如單一Type 和 Color
                {
                    sb.Append(" REFE " + (scale == 1 ? "1 " : "2 ") + String.Join(" ", refValue) + ";\r\n");
                    if (primary == 2) sb.Append("  SECS;\r\n");
                    if (refType != null) sb.Append("  TYPE " + String.Join(" ", refType) + ";\r\n");
                    if (refColor != null) sb.Append("  COLOR " + String.Join(" ", refColor) + ";\r\n  TCOLOR " + String.Join(" ", refColor) + ";\r\n");
                }
                else //For multiple reference line
                {
                    int[] typeArr = new int[refValue.Length];
                    int[] colorArr = new int[refValue.Length];
                    if (refType != null)//如果非 null, 建立等長度的陣列讓 for 迴圈使用
                    {
                        if (refType.Length == 1)
                        {
                            for (int i = 0; i < typeArr.Length; i++) typeArr[i] = refType[0];
                        }
                        else
                        {
                            typeArr = refType;
                        }
                    }


                    if (refColor.Length != null)
                    {
                        if (refColor.Length == 1)
                        {
                            for (int i = 0; i < colorArr.Length; i++) colorArr[i] = refColor[0];
                        }
                        else
                        {
                            colorArr = refColor;
                        }
                    }


                    for (int i = 0; i < refValue.Length; i++)
                    {
                        sb.Append(" REFE " + (scale == 1 ? "1 " : "2 ") + refValue[i] + ";\r\n");
                        if (primary == 2) sb.Append("  SECS;\r\n");
                        if (refType != null) sb.Append("  TYPE " + typeArr[i] + ";\r\n");
                        if (refColor != null) sb.Append("  COLOR " + colorArr[i] + ";\r\n  TCOLOR " + colorArr[i] + ";\r\n");
                    }
                }
            }
            else
            {
                sb.Append("###### Invalid reference setting #####\r\n");
            }
            return sb.ToString();
        }
        private bool[] CheckRefCmndType(double[] refValue, int[] refType = null, int[] refColor = null)
        {
            int typeCnt = 0;
            int colorCnt = 0;
            int cnt = refValue.Length;
            bool isOk = false;
            bool isOneType = false;
            if (refType != null) typeCnt = refType.Length;
            if (refColor != null) colorCnt = refColor.Length;

            if (typeCnt == 0)
            {
                if (colorCnt == 0)
                {
                    isOneType = true;
                    isOk = true;
                }
                else if (colorCnt == 1)
                {
                    isOneType = true;
                    isOk = true;
                }
                else if (colorCnt > 1 & colorCnt == cnt)
                {
                    isOk = true;
                }
            }
            else if (typeCnt == 1)
            {
                if (colorCnt == 0)
                {
                    isOneType = true;
                    isOk = true;
                }
                else if (colorCnt == 1)
                {
                    isOneType = true;
                    isOk = true;
                }
                else if (colorCnt > 1 & colorCnt == cnt)
                {
                    isOk = true;
                }
            }
            else if (typeCnt > 1 & typeCnt == cnt)
            {
                if (colorCnt <= 1)
                {
                    isOk = true;
                }
                else if (colorCnt > 1 & colorCnt == cnt)
                {
                    isOk = true;
                }
            }

            return new bool[] { isOk, isOneType };

        }
        protected String GetTitleVariableString(Mtb.Worksheet ws, List<String> varString)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < varString.Count; i++)
            {
                sb.Append(ws.Columns.Item(varString[i]).Label);
                if (sb.Length > 30)
                {
                    sb.Remove(27, sb.Length - 27);
                    sb.Append("...");
                    break;
                }
                else if (i == varString.Count - 1)
                {
                    sb.Append(".");
                }
                else
                {
                    sb.Append(", ");
                }
            }
            return sb.ToString();
        }
        protected Size GetLegendSize(Mtb.Worksheet ws, List<String> cols)
        {
            /*
             * 此函數只回傳變數名稱所需的大小，不含符號以及上下邊界，因為
             * trend 和 chart 的符號寬度不同，使用者需另外計算。
             */
            Size lgndSize = new Size(0, 0);
            foreach (String str in cols)
            {
                Size tmp = TextRenderer.MeasureText(ws.Columns.Item(str).Label, this.d_lgndFont);
                if (tmp.Width > lgndSize.Width)
                {
                    lgndSize.Width = tmp.Width;
                }
                lgndSize.Height = lgndSize.Height + tmp.Height;
            }
            return lgndSize;
        }

    }
}
