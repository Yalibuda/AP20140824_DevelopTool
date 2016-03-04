using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Mtb;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using MtbGraph.GraphComponent;

namespace MtbGraph
{
    public enum BarTypes
    {
        Stack,
        Cluster
    }
    public enum ScalePrimary
    {
        Primary,
        Secondary
    }
    [ClassInterface(ClassInterfaceType.None)] //自己設計接口
    public class BarLinePlot : MtbGraphFrame, IBarLinePlot
    {
        /*
         * 20150129:
         * 新增 Bar-line plot reference line 功能...以trend/bar 為 base 加入
         * 留意...因為是過渡期，所以抓取 Scale 時，不考慮 Reference 設定(自動除外)!!
         */
        public Reference BarRef { set; get; }
        public Reference TrendRef { set; get; }
        public BarLineLegendBox LegendBox { set; get; }
        public BarLinePlot()
        {
            BarRef = new Reference(ScaleType.Y_axis);
            TrendRef = new Reference(ScaleType.Y_axis);
            LegendBox = new BarLineLegendBox();
        }
        public void CreateBarLinePlot(Mtb.Project proj, Mtb.Worksheet ws, BarTypes btype = BarTypes.Stack)
        {

            /*
             * If only bar variable or line varialbe, just skip the procedure of
             * calculation for data region and legend box
             * 
             */
            if (hasBar == 0 & hasTrnd + hasTg == 0)
            {
                return;
            }
            else if (hasBar > 0 & hasTrnd + hasTg > 0)
            {
                CreateOverlayBarLinePlot(proj, ws, btype);
            }



        }
        //functions
        private void CreateOverlayBarLinePlot(Mtb.Project proj, Mtb.Worksheet ws, BarTypes btype = BarTypes.Stack)
        {
            
            List<String> barColList = new List<String>();
            List<String> trndColList = new List<String>();
            List<String> tgColList = new List<String>();
            String barStr = "";
            String trndStr = "";
            String tgStr = "";
            String labStr = "";
            double dXMin = 0.0467;
            //double dXMax = 0.9533;
            double dXMax = 0.97;
            double dYMin = 0.044;
            double dYMax = 0.93;
            double bLgndXMin;
            double bLgndXMax;
            double bLgndYMin;
            double bLgndYMax;
            double lLgndXMin;
            double lLgndXMax;
            double lLgndYMin;
            double lLgndYMax;
            int cmndCnt = 0;
            StringBuilder mtbCmnd = new StringBuilder();

            if (hasBar == 1)
            {
                barStr = String.Join(" ", this.barCols);
                barColList = da.GetMtbCols(this.barCols, ws);
            }
            if (hasTrnd == 1)
            {
                trndStr = String.Join(" ", this.trendCols);
                trndColList = da.GetMtbCols(this.trendCols, ws);
            }
            if (hasTg == 1)
            {
                tgStr = String.Join(" ", this.targetCols);
                tgColList = da.GetMtbCols(this.targetCols, ws);
            }
            if (hasLab == 1)
            {
                labStr = this.labCol[0];
            }

            /*
             * This part is try to get size of title, subtitle and footnote, and then modify
             * the location of data region
             */

            //Check Title
            Size sizeTitle = new Size(0, 0);
            if (this.title != String.Empty)
            {
                sizeTitle = TextRenderer.MeasureText((this.title == null ? mTitle : this.title), this.d_TitleFont);
                dYMax = dYMax - ((double)sizeTitle.Height / d_gHeight);
            }

            //Check subTitle
            //Check footnote


            /*
             * 計算 Legend Box 的 Size
             * Legend Box 物件在 Barline Plot 中不使用 GetCommand，因為
             * 疊圖的 LegendBox 位置要另外給。使用的目的是紀錄字型大小、
             * 是否隱藏標題等。(2015/5/12)
             */
            List<string> names = new List<string>();
            Size bLgndSize = new Size(0, 0);
            LegendBox.NotationType = LegendNotationType.Bar;
            foreach (string col in barColList)
            {
                names.Add(ws.Columns.Item(col).Label);
            }
            string[] refStr = names.ToArray();
            LegendBox.SetVariables(ref refStr);
            bLgndSize = LegendBox.GetSize();
            LegendBox.NotationType = LegendNotationType.Trend;
            Size lLgndSize = new Size(0, 0);
            names.Clear();
            foreach (string col in trndColList)
            {
                names.Add(ws.Columns.Item(col).Label);
            }
            foreach (string col in tgColList)
            {
                names.Add(ws.Columns.Item(col).Label);
            }
            refStr = names.ToArray();
            LegendBox.SetVariables(ref refStr);
            lLgndSize = LegendBox.GetSize();

            if (barColList.Count + trndColList.Count + tgColList.Count <= 3)
            {
                //bLgndYMax = 0.9767;
                bLgndYMax = 0.998;
            }
            else
            {
                bLgndYMax = Math.Min(dYMax + LegendBox.VerticalBase,1);
            }

            bLgndYMin = Math.Max(bLgndYMax - (double)bLgndSize.Height / d_gHeight, 0.005);
            bLgndXMin = (bLgndSize.Width < lLgndSize.Width) ? 0.998 - (double)lLgndSize.Width / d_gWidth : 0.998 - (double)bLgndSize.Width / d_gWidth;
            //bLgndXMax = 0.9767;
            if (bLgndXMin < 0.7) bLgndXMin = 0.7;
            bLgndXMax = 0.998;

            lLgndYMax = bLgndYMin;
            lLgndYMin = lLgndYMax - (double)lLgndSize.Height / d_gHeight;
            lLgndXMin = bLgndXMin;
            lLgndXMax = bLgndXMax;

            //Modify data region
            if (barColList.Count + trndColList.Count + tgColList.Count <= 3)
            {
                //dXMax = 0.9533;
                //dXMax = 0.97;
                dYMax = lLgndYMin - 0.015;
            }
            else
            {
                dXMax = bLgndXMin - 0.0234;
            }


            //Check primary scale label
            Size sizeLabel = new Size(0, 0);
            sizeLabel = TextRenderer.MeasureText("Label Text", this.d_LabFont);//y-axis label font: new Font("Segoe UI Semibold", 9, FontStyle.Bold)
            if (this.yLabel != String.Empty & (bScalePrimary == ScalePrimary.Primary || tScalePrimary == ScalePrimary.Primary))
            {
                dXMin = dXMin + ((double)sizeLabel.Height / d_gWidth);
            }
            //Check scendary scale label
            if (this.secsLabel != String.Empty & (bScalePrimary == ScalePrimary.Secondary || tScalePrimary == ScalePrimary.Secondary))
            {
                dXMax = dXMax - ((double)sizeLabel.Height / d_gWidth);
            }
            //Check x label
            if (this.xLabel != String.Empty)
            {
                dYMin = dYMin + ((double)sizeLabel.Height / d_gHeight);
            }


            //Collect data pool
            cmndCnt = proj.Commands.Count;
            MtbTools mtools = new MtbTools();
            String[] colStr = mtools.CreateVariableStrArray(ws, 5 +
                (isShowBDatlab & btype == BarTypes.Stack & barColList.Count > 1 ? 2 : 0), MtbVarType.Column);//加2是為了將結果堆疊
            String[] constStr = mtools.CreateVariableStrArray(ws, 12, MtbVarType.Constant);


            mtbCmnd.Append("NOTITLE\r\nBRIEF 0\r\n");
            if (btype == BarTypes.Stack)
            {
                mtbCmnd.Append("RSUM " + barStr + " " + colStr[0] + "\r\n");
            }
            else
            {
                if (barColList.Count == 1)
                {
                    mtbCmnd.Append("COPY " + barStr + " " + colStr[0] + "\r\n");
                }
                else
                {
                    mtbCmnd.Append("STACK " + barStr + " " + colStr[0] + "\r\n");
                }
            }

            if (this.bScalePrimary == this.tScalePrimary)
            {
                mtbCmnd.Append("STACK " + colStr[0] + ((hasTrnd == 1) ? " " + trndStr : String.Empty) +
                ((hasTg == 1) ? " " + tgStr : String.Empty) + " " + colStr[0] + "\r\n");
                mtbCmnd.Append("MINI " + colStr[0] + " " + constStr[0] + "\r\n");
                mtbCmnd.Append("MAXI " + colStr[0] + " " + constStr[1] + "\r\n");
                mtbCmnd.Append("LET " + constStr[0] + "=IF(" + constStr[0] + ">=0,0," + constStr[0] + ")\r\n");
                mtbCmnd.Append("GSCALE " + constStr[0] + " " + constStr[1] + ";\r\n" +
                    " NMIN 8;\r\n NMAX 15;\r\n" +
                    " SMIN " + constStr[2] + ";\r\n" + " SMAX " + constStr[3] + ";\r\n" +
                    " TMIN " + constStr[4] + ";\r\n" + " TMAX " + constStr[5] + ";\r\n" +
                    " NTIC " + constStr[6] + ".\r\n");
                mtbCmnd.Append("LET " + constStr[2] + "=IF(" + constStr[0] + ">=0,0," + constStr[2] + ")\r\n");
                mtbCmnd.Append("COPY " + constStr[4] + "-" + constStr[6] + " " + constStr[2] + " " + constStr[3] + " "
                    + constStr[7] + "-" + constStr[11] + "\r\n");
            }
            else
            {
                mtbCmnd.Append("MINI " + colStr[0] + " " + constStr[0] + "\r\n");
                mtbCmnd.Append("MAXI " + colStr[0] + " " + constStr[1] + "\r\n");
                mtbCmnd.Append("LET " + constStr[0] + "=IF(" + constStr[0] + ">=0,0," + constStr[0] + ")\r\n");
                mtbCmnd.Append("GSCALE " + constStr[0] + " " + constStr[1] + ";\r\n" +
                    " SMIN " + constStr[2] + ";\r\n" + " SMAX " + constStr[3] + ";\r\n" +
                    " TMIN " + constStr[4] + ";\r\n" + " TMAX " + constStr[5] + ";\r\n" +
                    " NTIC " + constStr[6] + ".\r\n");
                mtbCmnd.Append("LET " + constStr[2] + "=IF(" + constStr[0] + ">=0,0," + constStr[2] + ")\r\n");
                if (trndColList.Count + tgColList.Count == 1)
                {
                    mtbCmnd.Append("COPY " + ((hasTrnd == 1) ? trndStr : tgStr) + " " + colStr[0] + "\r\n");
                }
                else
                {
                    mtbCmnd.Append("STACK " + ((hasTrnd == 1) ? trndStr + " " : String.Empty) +
                        ((hasTg == 1) ? tgStr + " " : String.Empty) + colStr[0] + "\r\n");
                }

                mtbCmnd.Append("MINI " + colStr[0] + " " + constStr[0] + "\r\n");
                mtbCmnd.Append("MAXI " + colStr[0] + " " + constStr[1] + "\r\n");
                mtbCmnd.Append("GSCALE " + constStr[0] + " " + constStr[1] + ";\r\n" +
                    " SMIN " + constStr[10] + ";\r\n" + " SMAX " + constStr[11] + ";\r\n" +
                    " TMIN " + constStr[7] + ";\r\n" + " TMAX " + constStr[8] + ";\r\n" +
                    " NTIC " + constStr[9] + ".\r\n");
            }

            mtbCmnd.Append("COPY " + constStr[2] + "-" + constStr[11] + " " + colStr[1] + "\r\n");
            int n = 0;
            if (hasLab == 0)
            {
                n = ws.Columns.Item(barColList[0]).RowCount;
                mtbCmnd.Append("SET " + colStr[3] + "\r\n 1:" + n +
                    "\r\n END\r\n");
            }
            else
            {
                n = ws.Columns.Item(labCol[0]).RowCount;
                mtbCmnd.Append("TEXT " + labCol[0] + " " + colStr[3] + "\r\n");
            }

            if (barColList.Count == 1)
            {
                mtbCmnd.Append("TSET " + colStr[2] + "\r\n " + n + "(\"" + ws.Columns.Item(barColList[0]).Label + "\")\r\n END\r\n");
            }

            //Prepare stacked bar chart data label
            if (isShowBDatlab & btype == BarTypes.Stack & barColList.Count > 1)
            {
                mtbCmnd.Append("STACK " + barStr + " " + colStr[4] + ";\r\n SUBS " + colStr[5] + ".\r\n");
            }

            /*
             * Prepare tmp macro
             * 
             */
            String path;
            if (Environment.GetEnvironmentVariable("tmp").Equals(String.Empty))
            {
                path = Environment.GetEnvironmentVariable("tmp");
            }
            else { path = Environment.GetEnvironmentVariable("temp"); }

            path = path + "\\~macro.mtb";
            FileStream fs = new FileStream(path, FileMode.Create);
            fs.Close();

            StreamWriter sw = new StreamWriter(path);
            sw.Write(mtbCmnd.ToString());
            sw.Close();
            proj.ExecuteCommand("EXEC '" + path + "' 1", ws);


            if (expCmnd) ExportCommand(mtbCmnd.ToString(), cmndPath);

            /*
             * Get location value of data region
             */
            double[] tickValue;
            Size tMinSize = new Size(0, 0);
            Size tMaxSize = new Size(0, 0);
            tickValue = ws.Columns.Item(colStr[1]).GetData();
            if (this.bScalePrimary == this.tScalePrimary)
            {
                tMinSize = TextRenderer.MeasureText(tickValue[2].ToString(), this.d_TickFont);
                tMaxSize = TextRenderer.MeasureText(tickValue[3].ToString(), this.d_TickFont);
                if (this.bScalePrimary == ScalePrimary.Primary)
                {
                    dXMin = dXMin + ((double)Math.Max(tMinSize.Width, tMaxSize.Width) / d_gWidth);
                }
                else
                {
                    dXMax = dXMax - ((double)Math.Max(tMinSize.Width, tMaxSize.Width) / d_gWidth);
                }
            }
            else
            {
                if (this.bScalePrimary == ScalePrimary.Primary)
                {
                    tMinSize = TextRenderer.MeasureText(tickValue[2].ToString(), this.d_TickFont);
                    tMaxSize = TextRenderer.MeasureText(tickValue[3].ToString(), this.d_TickFont);
                    dXMin = dXMin + ((double)Math.Max(tMinSize.Width, tMaxSize.Width) / d_gWidth);
                    tMinSize = TextRenderer.MeasureText(tickValue[5].ToString(), this.d_TickFont);
                    tMaxSize = TextRenderer.MeasureText(tickValue[6].ToString(), this.d_TickFont);
                    dXMax = dXMax - ((double)Math.Max(tMinSize.Width, tMaxSize.Width) / d_gWidth);
                }
                else
                {
                    tMinSize = TextRenderer.MeasureText(tickValue[2].ToString(), this.d_TickFont);
                    tMaxSize = TextRenderer.MeasureText(tickValue[3].ToString(), this.d_TickFont);
                    dXMax = dXMax - ((double)Math.Max(tMinSize.Width, tMaxSize.Width) / d_gWidth);
                    tMinSize = TextRenderer.MeasureText(tickValue[5].ToString(), this.d_TickFont);
                    tMaxSize = TextRenderer.MeasureText(tickValue[6].ToString(), this.d_TickFont);
                    dXMin = dXMin + ((double)Math.Max(tMinSize.Width, tMaxSize.Width) / d_gWidth);
                }
            }
            if (hasLab == 1)
            {
                sizeLabel = GetStringSize(ws, colStr[3], this.d_LabFont);
                dYMin = dYMin + (double)sizeLabel.Width * Math.Abs(Math.Sin(Math.PI * (this.xLabelAngle < 1.23456E+30 ? this.xLabelAngle : 45) / 180.0)) / d_gHeight;
            }


            /*
             * Start generate graph
             */
            mtbCmnd.Clear();
            if (this.dNum < 100)
            {
                mtbCmnd.Append("FNUM " + trndStr + ";\r\n FIXED " + dNum + ".\r\n");
            }

            //Prepare datlabel position for stack bar chart            
            if (isShowBDatlab & btype == BarTypes.Stack & barColList.Count > 1)
            {
                double deltaDUnit;
                double dMin, dMax;
                double deltaFUnit;
                deltaFUnit = dYMax - dYMin;
                if (bScalePrimary == ScalePrimary.Primary)
                {
                    deltaDUnit = (this.yScaleMax != 1.23456E+30 ? this.yScaleMax : tickValue[1]) -
                        (this.yScaleMin != 1.23456E+30 ? this.yScaleMin : tickValue[0]);
                }
                else
                {
                    deltaDUnit = (this.secScaleMax != 1.23456E+30 ? this.secScaleMax : tickValue[1]) -
                        (this.secScaleMin != 1.23456E+30 ? this.secScaleMin : tickValue[0]);
                }
                mtbCmnd.Append("LET " + colStr[6] + "=" + colStr[4] + "*-" + (deltaFUnit / deltaDUnit) + "\r\n");
            }

            sw = new StreamWriter(path);
            sw.Write(mtbCmnd.ToString());
            sw.Close();
            proj.ExecuteCommand("EXEC '" + path + "' 1", ws);
            if (expCmnd) ExportCommand(mtbCmnd.ToString(), cmndPath, false);
            mTitle = "My Bar-Line Plot Of " + GetTitleVariableString(ws, barColList.Union(trndColList).ToList());
            mtbCmnd.Clear();
            mtbCmnd.Append("TITLE\r\n");
            mtbCmnd.Append("LAYOUT;\r\n");
            if (this.gSave) mtbCmnd.Append(" GSAVE \"" + gPath + "\";\r\n  JPEG;\r\n  REPL;\r\n");
            mtbCmnd.Append(" WTITLE \"" + mTitle + "\".\r\n");

            mtbCmnd.Append(" CHART (" + barStr + ")*" + ((hasLab == 0) ? colStr[3] : labStr) + ";\r\n");
            mtbCmnd.Append("  SUMM;\r\n");
            if (barColList.Count == 1)
            {
                mtbCmnd.Append("  BAR " + colStr[2] + ";\r\n");
                if (!showBarEdge) mtbCmnd.Append("    ETYPE 0;\r\n");
                mtbCmnd.Append("  GROUP " + colStr[2] + ";\r\n");
            }
            else
            {
                mtbCmnd.Append("  OVER;\r\n   VLAST;\r\n");
                mtbCmnd.Append("  BAR;\r\n   VASS;\r\n");
                if (!showBarEdge) mtbCmnd.Append("    ETYPE 0;\r\n");
            }
            if (btype == BarTypes.Stack) mtbCmnd.Append("  STACK;\r\n");
            if (bScalePrimary == ScalePrimary.Secondary)
            {
                mtbCmnd.Append("  Scale 2;\r\n   LDIS 1 0 0 0;\r\n   HDIS 1 1 1 0;\r\n");
                if (this.secScaleMin < 1.23456E+30) mtbCmnd.Append("   Min " + this.secScaleMin + ";\r\n");
                if (this.secScaleMax < 1.23456E+30) mtbCmnd.Append("   Max " + this.secScaleMax + ";\r\n");
                if (this.secTickAttr == ScaleTickAttribute.None)
                {
                    if (this.secScaleMax != 1.23456E+30 || this.secScaleMin != 1.23456E+30) mtbCmnd.Append("   NMAJ 11;\r\n");
                }
                else if (this.secTickAttr == ScaleTickAttribute.IntervalBetweenTick)
                {
                    mtbCmnd.Append("  TICK " + (this.secScaleMin < 1.23456E+30 ? this.secScaleMin : tickValue[0]) + ":" +
                        (this.secScaleMax < 1.23456E+30 ? this.secScaleMax : tickValue[1]) + "/" + this.secTickAttrValue + ";\r\n");
                    if (this.secScaleMin == 1.23456E+30) mtbCmnd.Append("   Min " + tickValue[0] + ";\r\n");
                    if (this.secScaleMax == 1.23456E+30) mtbCmnd.Append("   Max " + tickValue[1] + ";\r\n");
                }
                else
                {
                    mtbCmnd.Append("  NMAJ " + this.secTickAttrValue + ";\r\n");
                }

                mtbCmnd.Append("  AXLA 2" + (this.secsLabel == String.Empty ? "\"\"" : (this.secsLabel == null ? "" : "\"" + this.secsLabel + "\"")) + ";\r\n   ADIS 2;\r\n");
                if (BarRef.haveValues() == true)
                {
                    BarRef.Side = 2;
                    mtbCmnd.AppendLine(BarRef.GetCommand());
                }
            }
            else
            {
                mtbCmnd.Append("  Scale 2;\r\n   LDIS 1 1 1 0;\r\n   HDIS 1 0 0 0;\r\n");
                if (this.yScaleMin < 1.23456E+30) mtbCmnd.Append("   Min " + this.yScaleMin + ";\r\n");
                if (this.yScaleMax < 1.23456E+30) mtbCmnd.Append("   Max " + this.yScaleMax + ";\r\n");
                if (this.yTickAttr == ScaleTickAttribute.None)
                {
                    if (this.yScaleMax != 1.23456E+30 || this.yScaleMin != 1.23456E+30) mtbCmnd.Append("   NMAJ 11;\r\n");
                }
                else if (this.yTickAttr == ScaleTickAttribute.IntervalBetweenTick)
                {
                    mtbCmnd.Append("  TICK " + (this.yScaleMin < 1.23456E+30 ? this.yScaleMin : tickValue[0]) + ":" +
                        (this.yScaleMax < 1.23456E+30 ? this.yScaleMax : tickValue[1]) + "/" + this.yTickAttrValue + ";\r\n");
                    if (this.yScaleMin == 1.23456E+30) mtbCmnd.Append("   Min " + tickValue[0] + ";\r\n");
                    if (this.yScaleMax == 1.23456E+30) mtbCmnd.Append("   Max " + tickValue[1] + ";\r\n");
                }
                else
                {
                    mtbCmnd.Append("  NMAJ " + this.yTickAttrValue + ";\r\n");
                }
                mtbCmnd.Append("  AXLA 2" + (this.yLabel == String.Empty ? "\"\"" : (this.yLabel == null ? "" : "\"" + this.yLabel + "\"")) + ";\r\n   ADIS 1;\r\n");
                if (BarRef.haveValues() == true)
                {
                    BarRef.Side = 1;
                    mtbCmnd.AppendLine(BarRef.GetCommand());
                }
            }
            mtbCmnd.Append("  SCALE 1;\r\n   ANGLE " + (this.xLabelAngle < 1.23456E+30 ? this.xLabelAngle : 45) + ";\r\n");
            mtbCmnd.Append("  AxLa 1;\r\n   LSHOW;\r\n  TSHOW;\r\n");
            if (this.yRefValue != null)
            {
                mtbCmnd.Append(GetRefCmndString(this.yRefValue, this.yRefType, this.yRefColor));
            }
            mtbCmnd.Append("  LEGE " + bLgndXMin + " " + bLgndXMax + " " + bLgndYMin + " " + bLgndYMax + ";\r\n");
            mtbCmnd.Append("   ETYPE 0;\r\n   TYPE 0;\r\n   SECT 1;\r\n   CHHIDE;\r\n");

            //Set data labels
            if (isShowBDatlab)
            {
                if (btype == BarTypes.Stack & barColList.Count > 1)
                {
                    double[] dataArr = ws.Columns.Item(colStr[6]).GetData();
                    double[] modelArr = ws.Columns.Item(colStr[5]).GetData();
                    mtbCmnd.Append("  DATLAB " + colStr[4] + ";\r\n   PLAC 0 0;\r\n");
                    for (int i = 0; i < dataArr.Length; i++)
                    {
                        //mtbCmnd.Append("   POSI " + (i + 1) + ";\r\n    MODEL " + modelArr[i] + ";\r\n" +
                        //    "    OFFS 0 " + dataArr[i] + ";\r\n   ENDP;\r\n");
                        if (dataArr[i] < 1.23456E+30)
                        {
                            mtbCmnd.AppendLine(String.Format("   POSI {0};" + Environment.NewLine +
                                "    MODEL {1};" + Environment.NewLine +
                                "    OFFS 0 {2};" + Environment.NewLine +
                                "   ENDP;", (i + 1), modelArr[i], dataArr[i]));
                        }
                    }

                }
                else
                {
                    mtbCmnd.Append("  DATLAB;\r\n");
                }

            }

            mtbCmnd.Append("  DATA " + dXMin + " " + dXMax + " " + dYMin + " " + dYMax + ";\r\n");
            if (this.title != String.Empty)
            {
                mtbCmnd.Append("  TITLE \"" + (this.title == null ? mTitle : this.title) + "\";\r\n");
                mtbCmnd.Append("   OFFSET " + (((dXMin + dXMax) / 2) - 0.5) + " " + ((double)-sizeTitle.Height / d_gHeight) + ";\r\n");
            }
            mtbCmnd.Append(" NODT.\r\n");

            /*
             * 處理當 trend 總數 = 1，無法顯示legend box的狀況...建立一個虛擬欄位來 group 產生 legend
             */
            String[] trndgroupCol = mtools.CreateVariableStrArray(ws, 1, MtbVarType.Column);
            if (trndColList.Count + tgColList.Count == 1)
            {
                mtbCmnd.AppendLine("TSET " + trndgroupCol[0]);
                if (trndColList.Count == 1)
                {
                    mtbCmnd.AppendLine(ws.Columns.Item(trndColList[0]).RowCount + "(\"" + ws.Columns.Item(trndColList[0]).Label + "\")");
                }
                else
                {
                    mtbCmnd.AppendLine(ws.Columns.Item(tgColList[0]).RowCount + "(\"" + ws.Columns.Item(tgColList[0]).Label + "\")");
                }
                mtbCmnd.AppendLine("END");
            }
            mtbCmnd.Append(" TSPLOT " + trndStr + " " + tgStr + ";\r\n  NOEM;\r\n  NOMI;\r\n  OVER;\r\n");
            //Set symbol variable
            String symbStr = "";
            String connStr = "";
            String colorStr = "";
            List<int> symbList = new List<int>();
            List<int> colorList = new List<int>();
            List<int> connList = new List<int>();
            if (hasTrnd == 1)
            {
                for (int i = 0; i < trndColList.Count; i++) symbList.Add(dSymbType[i % this.dSymbType.Length]);

                if (this.targetColor != null)
                {
                    for (int i = 0; i < trndColList.Count; i++) colorList.Add(this.dLineColor[i % this.dLineColor.Length]);
                }
                if (this.targetType != null)
                {
                    for (int i = 0; i < trndColList.Count; i++) connList.Add(this.dLineType[i % this.dLineType.Length]);
                }
            }
            if (hasTg == 1)
            {
                for (int i = 0; i < tgColList.Count; i++) symbList.Add(0);
                if (this.targetColor != null)
                {
                    for (int i = 0; i < tgColList.Count; i++) colorList.Add(this.targetColor[i % this.targetColor.Count]);
                }
                if (this.targetType != null)
                {
                    for (int i = 0; i < tgColList.Count; i++) connList.Add(this.targetType[i % this.targetType.Count]);
                }
            }
            symbStr = String.Join(" ", symbList);
            colorStr = String.Join(" ", colorList);
            connStr = String.Join(" ", connList);

            if (trndColList.Count + tgColList.Count == 1)
            {
                mtbCmnd.AppendLine("  SYMB " + trndgroupCol[0] + ";");
            }
            else
            {
                mtbCmnd.AppendLine("  SYMB;");
            }
            mtbCmnd.Append("   TYPE " + symbStr + ";\r\n" +
                ((targetColor != null) ? "   COLOR " + colorStr + ";\r\n" : ""));

            if (trndColList.Count + tgColList.Count == 1)
            {
                mtbCmnd.AppendLine("  CONN " + trndgroupCol[0] + ";");
            }
            else
            {
                mtbCmnd.AppendLine("  CONN;");
            }

            mtbCmnd.Append(((targetType != null) ? "   TYPE " + connStr + ";\r\n" : "") +
                ((targetColor != null) ? "   COLOR " + colorStr + ";\r\n" : ""));


            mtbCmnd.Append("  SCALE 1;\r\n   MIN 0.5;\r\n   MAX " + ((double)n + 0.5) + ";\r\n" +
                "   LDIS 0 1 1 0;\r\n   HDIS 0 0 0 0;\r\n   ANGLE " + (this.xLabelAngle < 1.23456E+30 ? this.xLabelAngle : 45) + ";\r\n");
            mtbCmnd.AppendLine("   TICK 1:" + ws.Columns.Item(barColList[0]).RowCount + ";");//不確定是使用 trend 或是 target..直接使用 bar variable 的長度
            if (hasLab == 1) mtbCmnd.Append("  STAMP " + labCol[0] + ";\r\n");
            mtbCmnd.Append("  AXLA 1 " + (this.xLabel == String.Empty ? ";\r\n   ADIS 0;\r\n" :
                (hasLab == 1 ? (this.xLabel == null ? "" : "\"" + this.xLabel + "\"") : ";\r\n") + "   ADIS 1;\r\n"));
            if (this.tScalePrimary == ScalePrimary.Primary)
            {
                mtbCmnd.Append("  SCALE 2;\r\n");
                mtbCmnd.Append("   LDIS " + (this.bScalePrimary == ScalePrimary.Secondary ? "0 1 1 0;\r\n" : "0 0 0 0;\r\n"));
                mtbCmnd.Append("   HDIS 0 0 0 0;\r\n");
                if (this.yScaleMin != 1.23456E+30)
                {
                    mtbCmnd.Append("   MIN " + this.yScaleMin + ";\r\n");
                }
                else if (this.tScalePrimary == this.bScalePrimary)
                {
                    mtbCmnd.Append("   MIN " + tickValue[0] + ";\r\n");
                }
                if (this.yScaleMax != 1.23456E+30)
                {
                    mtbCmnd.Append("   MAX " + this.yScaleMax + ";\r\n");
                }
                else if (this.tScalePrimary == this.bScalePrimary)
                {
                    mtbCmnd.Append("   MAX " + tickValue[1] + ";\r\n");
                }
                //Set the tick
                if (this.yTickAttr == ScaleTickAttribute.None)
                {
                    if (this.yScaleMax != 1.23456E+30 || this.yScaleMin != 1.23456E+30) mtbCmnd.Append("   NMAJ 11;\r\n");
                }
                else if (this.yTickAttr == ScaleTickAttribute.IntervalBetweenTick)
                {
                    mtbCmnd.Append("  TICK " + (this.yScaleMin < 1.23456E+30 ? this.yScaleMin : tickValue[8]) + ":" +
                        (this.yScaleMax < 1.23456E+30 ? this.yScaleMax : tickValue[9]) + "/" + this.yTickAttrValue + ";\r\n");
                    if (this.yScaleMin == 1.23456E+30) mtbCmnd.Append("   Min " + tickValue[8] + ";\r\n");
                    if (this.yScaleMax == 1.23456E+30) mtbCmnd.Append("   Max " + tickValue[9] + ";\r\n");
                }
                else
                {
                    mtbCmnd.Append("  NMAJ " + this.yTickAttrValue + ";\r\n");
                }
                if (TrendRef.haveValues() == true)
                {
                    TrendRef.Side = 1;
                    mtbCmnd.AppendLine(TrendRef.GetCommand());
                }

                mtbCmnd.Append("  AXLA 2" + (this.yLabel == String.Empty ? "\"\"" : (this.yLabel == null ? "" : " \"" + this.yLabel + "\"")) + ";\r\n");
                mtbCmnd.Append("   ADIS " + (this.bScalePrimary == ScalePrimary.Primary ? "0" : "1") + ";\r\n");
            }
            else
            {
                mtbCmnd.Append("  SCALE 2;\r\n");
                mtbCmnd.Append("   LDIS 0 0 0 0;\r\n");
                mtbCmnd.Append("   HDIS " + (this.bScalePrimary == ScalePrimary.Primary ? "0 1 1 0;\r\n" : "0 0 0 0;\r\n"));
                if (this.secScaleMin < 1.23456E+30)
                {
                    mtbCmnd.Append("   MIN " + this.secScaleMin + ";\r\n");
                }
                else if (this.tScalePrimary == this.bScalePrimary)
                {
                    mtbCmnd.Append("   MIN " + tickValue[0] + ";\r\n");
                }
                if (this.secScaleMax < 1.23456E+30)
                {
                    mtbCmnd.Append("   MAX " + this.secScaleMax + ";\r\n");
                }
                else if (this.tScalePrimary == this.bScalePrimary)
                {
                    mtbCmnd.Append("   MAX " + tickValue[1] + ";\r\n");
                }
                //Set the tick
                if (this.secTickAttr == ScaleTickAttribute.None)
                {
                    if (this.secScaleMax != 1.23456E+30 || this.secScaleMin != 1.23456E+30) mtbCmnd.Append("   NMAJ 11;\r\n");
                }
                else if (this.secTickAttr == ScaleTickAttribute.IntervalBetweenTick)
                {
                    mtbCmnd.Append("  TICK " + (this.secScaleMin < 1.23456E+30 ? this.secScaleMin : tickValue[8]) + ":" +
                        (this.secScaleMax < 1.23456E+30 ? this.secScaleMax : tickValue[9]) + "/" + this.secTickAttrValue + ";\r\n");
                    if (this.secScaleMin == 1.23456E+30) mtbCmnd.Append("   Min " + tickValue[8] + ";\r\n");
                    if (this.secScaleMax == 1.23456E+30) mtbCmnd.Append("   Max " + tickValue[9] + ";\r\n");
                }
                else
                {
                    mtbCmnd.Append("  NMAJ " + this.secTickAttrValue + ";\r\n");
                }
                if (TrendRef.haveValues() == true)
                {
                    TrendRef.Side = 2;
                    mtbCmnd.AppendLine(TrendRef.GetCommand());
                }

                mtbCmnd.Append("  AXLA 2" + (this.secsLabel == String.Empty ? "\"\"" : (this.secsLabel == null ? "" : " \"" + this.secsLabel + "\"")) + ";\r\n");
                mtbCmnd.Append("   ADIS " + (this.bScalePrimary == ScalePrimary.Secondary ? "0" : "2") + ";\r\n");
            }
            if (isShowTDatlab & hasTrnd == 1)
            {
                mtbCmnd.Append("  DATLAB;\r\n");
                if (hasTg == 1)
                {
                    for (int i = trndColList.Count + 1; i <= trndColList.Count + tgColList.Count; i++)
                    {
                        for (int j = 1; j <= n; j++)
                        {
                            mtbCmnd.Append("   POSI " + j + " " + "\"\";\r\n    MODEL " + i + ";\r\n   ENDP;\r\n");
                        }

                    }
                }
            }

            if (isShowTgDatlab & hasTg == 1)
            {
                StringBuilder sb = new StringBuilder();
                foreach (String str in tgColList)
                {
                    sb.Append(ws.Columns.Item(str).Label + ": " + GetTargetInfo(ws.Columns.Item(str)));
                }
                mtbCmnd.Append(" FOOT \"" + sb.ToString() + "\";\r\n");
            }

            mtbCmnd.Append("  LEGE " + lLgndXMin + " " + lLgndXMax + " " + lLgndYMin + " " + lLgndYMax + ";\r\n");
            mtbCmnd.Append("   ETYPE 0;\r\n   TYPE 0;\r\n   SECT 1;\r\n   CHHIDE;\r\n");
            mtbCmnd.Append("  DATA " + dXMin + " " + dXMax + " " + dYMin + " " + dYMax + ";\r\n   TYPE 0;\r\n   ETYPE 0;\r\n");
            mtbCmnd.Append(" NODT.\r\n");

            mtbCmnd.Append("ENDL\r\n");
            mtbCmnd.Append("NOTI\r\n");
            sw = new StreamWriter(path);
            //sw = new StreamWriter(path, false, System.Text.Encoding.GetEncoding("BIG5"));
            sw.Write(mtbCmnd.ToString());
            sw.Close();
            proj.ExecuteCommand("EXEC '" + path + "' 1", ws);
            if (expCmnd) ExportCommand(mtbCmnd.ToString(), cmndPath, false);

            if (copyToClip) CopyToClipboard("LAYOUT", proj, ws, cmndCnt + 1, proj.Commands.Count);


            mtbCmnd.Clear();
            mtbCmnd.Append("ERASE " + colStr[0] + "-" + colStr[colStr.Length - 1] + " " +
                constStr[0] + "-" + constStr[constStr.Length - 1] + " " + trndgroupCol[0] + "\r\n");
            mtbCmnd.Append("TITLE\r\nBRIEF 2\r\n");

            sw = new StreamWriter(path);
            sw.Write(mtbCmnd.ToString());
            sw.Close();
            proj.ExecuteCommand("EXEC '" + path + "' 1", ws);

        }

        private String GetTargetInfo(Mtb.Column col)
        {
            String[] data = new String[1];
            StringBuilder sb = new StringBuilder();
            switch (col.DataType)
            {
                case Mtb.MtbDataTypes.Numeric:
                case Mtb.MtbDataTypes.DateTime:
                    Double[] dblArr;
                    dblArr = col.GetData();
                    data = new String[dblArr.Length];
                    for (int i = 0; i < data.Length; i++)
                    {
                        data[i] = dblArr[i].ToString();
                    }
                    break;
                case Mtb.MtbDataTypes.Text:
                    data = col.GetData();
                    break;
            }

            for (int i = 0; i < data.Length; i++)
            {
                if (i == 0)
                {
                    sb.Append(data[i]);
                }
                else
                {
                    if (data[i] != data[i - 1])
                    {
                        sb.Append(", " + data[i]);
                    }
                }
            }
            return sb.ToString();
        }

        //Attribute parameters        
        public void SetBarVariable(String inputStr)
        {
            if (!String.IsNullOrEmpty(inputStr))
            {
                barCols = da.GetMtbColInfo(inputStr);
                hasBar = 1;
            }
            else
            {
                hasBar = 0;
            }

        }
        public void RemoveBarVariable()
        {
            barCols = null;
            hasBar = 0;
        }
        public void SetBarDatLabel(bool b)
        {
            isShowBDatlab = b;
        }
        public void SetTrendDatLabel(bool b)
        {
            isShowTDatlab = b;
        }
        public void SetTrendDatLabel(bool b, int decimalNumber = 100)
        {
            isShowTDatlab = b;
            if (decimalNumber < 17) dNum = decimalNumber;
        }
        public void SetTargetDatLabel(bool b)
        {
            isShowTgDatlab = b;
        }
        public void SetLabelVariable(String inputStr)
        {
            if (!String.IsNullOrEmpty(inputStr))
            {
                labCol = da.GetMtbColInfo(inputStr);
                hasLab = 1;
            }
            else
            {
                hasLab = 0;
            }

        }
        public void RemoveLabelVariable()
        {
            labCol = null;
            hasLab = 0;
        }
        public void SetScalePrimary(ScalePrimary barScale, ScalePrimary lineScale)
        {
            bScalePrimary = barScale;
            tScalePrimary = lineScale;
        }
        public void SetTrendVariable(String inputStr)
        {
            if (!String.IsNullOrEmpty(inputStr))
            {
                trendCols = da.GetMtbColInfo(inputStr);
                hasTrnd = 1;
            }
            else
            {
                hasTrnd = 0;
            }

        }
        public void RemoveTrendVariable()
        {
            trendCols = null;
            hasTrnd = 0;
        }
        public void SetTargetVariable(String inputStr)
        {
            if (!String.IsNullOrEmpty(inputStr))
            {
                targetCols = da.GetMtbColInfo(inputStr);
                hasTg = 1;
            }
            else
            {
                hasTg = 0;
            }
        }
        public void RemoveTargetVariable()
        {
            targetCols = null;
            hasTg = 0;
        }
        public void SetTypeOfTarget(ref int[] intArr)
        {
            targetType = new List<int>();
            for (int i = 0; i < intArr.Length; i++)
            {
                this.targetType.Add(intArr[i]);
            }
        }
        public void SetColorOfTarget(ref int[] intArr)
        {
            targetColor = new List<int>();
            for (int i = 0; i < intArr.Length; i++)
            {
                this.targetColor.Add(intArr[i]);
            }
        }
        public void SetDefaultTargetAttribute()
        {
            this.targetType = null;
            this.targetColor = null;
        }

        public void EnableBarEdge(bool b)
        {
            this.showBarEdge = b;
        }


        public override void Dispose()
        {
           
           base.Dispose();
        }

        /*
         * 變數宣告
         */
        private List<String> barCols;
        private List<String> trendCols;
        private List<String> targetCols;
        private List<String> labCol;

        private ScalePrimary bScalePrimary = ScalePrimary.Primary;
        private ScalePrimary tScalePrimary = ScalePrimary.Primary;

        private bool isShowBDatlab;
        private bool isShowTDatlab;
        private bool isShowTgDatlab;

        private String mTitle = "Bar-Trend Chart";
        //private String mXLabel = null;


        private List<int> targetType = null;
        private List<int> targetColor = null;

        private DialogAppraiser da = new DialogAppraiser();

        private int dNum = 100;
        private int hasBar = 0;
        private int hasTrnd = 0;
        private int hasTg = 0;
        private int hasLab = 0;

        private bool showBarEdge = true;

        //private double xlabAglign = 45;

    }
}
