using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MtbGraph.GraphComponent;
using System.Drawing;
using System.Windows.Forms;
using System.Collections;
using System.IO;

namespace MtbGraph.MyBarChart
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class Chart : GraphFrameWork, ICOMInterop_Barchart
    {
        public BarchartXScale X_Scale { set; get; }
        public BarchartYScale Y_Scale { set; get; }
        public SimpleLegend LegendBox { set; get; }
        public Annotation Annotation { set; get; }
        public Datlab Datalabel { get; set; }

        //Bar chart 的屬性...如 Stack or Cluster/ Rows outermost or Columns outermost
        public BarChartTableArrangement TableArrangement { set; get; }
        public BarChartType ChartType { set; get; }

        public Chart()
            : base()
        {
            this.X_Scale = new BarchartXScale(ScaleType.X_axis);
            this.Y_Scale = new BarchartYScale(ScaleType.Y_axis);
            this.Annotation = new Annotation();
            this.Datalabel = new Datlab();
            this.LegendBox = new SimpleLegend();
            this.LegendBox.NotationType = LegendNotationType.Bar;
            this.TableArrangement = BarChartTableArrangement.RowsOuterMost;
            this.ChartType = BarChartType.Stack;
        }

        public Chart(Mtb.Project proj, Mtb.Worksheet ws)
            : base(proj, ws)
        {
            this.X_Scale = new BarchartXScale(ScaleType.X_axis);
            this.Y_Scale = new BarchartYScale(ScaleType.Y_axis);
            this.Annotation = new Annotation();
            this.Datalabel = new Datlab();
            this.LegendBox = new SimpleLegend();
            this.LegendBox.NotationType = LegendNotationType.Bar;
            this.TableArrangement = BarChartTableArrangement.RowsOuterMost;
            this.ChartType = BarChartType.Stack;
        }

        private List<String> variables = null;
        private MtbTools mtools = new MtbTools();
        public void SetVariable(ref object variables)
        {
            this.variables = mtools.TransObjToMtbColList(variables, ws);

        }

        private List<String> labvariable = null;
        public void SetLabelVarible(ref object variables)
        {
            labvariable = mtools.TransObjToMtbColList(variables, ws);
        }

        int datcolor = -1;
        public void SetDatalabelColor(int color)
        {
            if (color < 1 || color > 129)
            {
                this.Datalabel.Color = DatlabColor.Default;
                return;
            }
            this.datcolor = color;
            this.Datalabel.Color = DatlabColor.Custom;
        }

        public String GetCommand()
        {
            StringBuilder cmnd = new StringBuilder();

            //先設置各軸對應的變數...
            Object obj = this.variables.ToArray();
            this.Y_Scale.ChartType = this.ChartType;
            this.Y_Scale.TableArrangement = this.TableArrangement;
            this.Y_Scale.SetScaleVariable(ref obj, ws, proj);


            cmnd.AppendLine("CHART (" + String.Join(" ", this.variables) + ")*" + this.labvariable[0] + ";");
            cmnd.AppendLine(" OVER;");
            //判斷 Table arrangement
            if (this.TableArrangement == BarChartTableArrangement.RowsOuterMost)
                cmnd.AppendLine("  VLAST;");
            else
                cmnd.AppendLine("  VFIRST;");

            //如果只有一個欄位的話，使用 STACK 指令會錯誤}
            if (this.ChartType == BarChartType.Stack & this.variables.Count > 1) cmnd.AppendLine(" STACK 1;");

            //Chart 使用 (C...C)*C 的語法時，一定要搭配 Summary 指令
            cmnd.AppendLine(" SUMM;");

            //設置雙軸指令
            cmnd.Append(this.X_Scale.GetCommand());
            cmnd.Append(this.Y_Scale.GetCommand());

            /*
             * 處理 bar...
             * 在 Cluster 的時候，tick show 出所有層別，而且 bar 不會上色..因此
             * 1. 用 TShow 隱藏部分 tick
             * 2. 用 Bar 指令依類別上色
             *    - VLast: 用 VASS 指令依資料行分類
             *    - VFirst: 直接給 label 分類
             * 
             */
            if (this.ChartType == BarChartType.Cluster)
            {
                cmnd.AppendLine(" TSHOW 1;");
                if (this.TableArrangement == BarChartTableArrangement.RowsOuterMost)//VLAST                
                    cmnd.AppendLine(" BAR;" + Environment.NewLine + "  VASS;");
                else
                    cmnd.AppendLine(" BAR " + this.labvariable[0] + ";");
            }


            //if (this.LegendBox.Show)
            //{
            //    if ((this.LegendBox.HideHead == true & this.variables.Count <= 3) ||
            //        (this.LegendBox.HideHead == false & this.variables.Count <= 2))
            //        this.LegendBox.Location = Location.RightTop;
            //    if (this.LegendBox.Location != Location.Auto)
            //    {
            //        String[] colname = new String[this.variables.Count];
            //        for (int i = 0; i < this.variables.Count; i++)
            //        {
            //            colname[i] = ws.Columns.Item(this.variables[i]).Label;
            //        }
            //        this.LegendBox.SetVariables(ref colname);
            //    }
            //    cmnd.Append(LegendBox.GetCommand());
            //}
            cmnd.Append(this.LegendBox.GetCommand());

            /**************************************************************************
             * 處理 Datalab...算是這個方法的大工程
             * 這裡主要是
             * 1. 建構 DatlabModelAttribute
             * 2. 當 Stack 時，計算需要 offset 的距離
             * 
             ***************************************************************************/
            if (this.Datalabel.Show)
            {
                DatlabModelAttribute model;
                List<DatlabModelAttribute> models = new List<DatlabModelAttribute>();
                /*
                 * Datlab 的規則依照 VLast, VFirst 走...所以 start, end 給法不同
                 * VFirst 比較直覺...position 使用 row index
                 * VLast position 使用累積位置...
                 */
                if (this.TableArrangement == BarChartTableArrangement.RowsOuterMost)
                {
                    int cumulateCount = 0;
                    for (int i = 0; i < this.variables.Count; i++)
                    {
                        model = new DatlabModelAttribute();
                        model.ModelIndex = i + 1;
                        model.Start = cumulateCount + 1;
                        cumulateCount = cumulateCount + ws.Columns.Item(this.variables[i]).RowCount;
                        model.End = cumulateCount;
                        models.Add(model);
                    }
                }
                else
                {
                    for (int i = 0; i < this.variables.Count; i++)
                    {
                        model = new DatlabModelAttribute();
                        model.ModelIndex = i + 1;
                        model.Start = 1;
                        model.End = ws.Columns.Item(this.variables[i]).RowCount;
                        models.Add(model);
                    }
                }


                if (this.Datalabel.Color == DatlabColor.Custom &
                        (this.datcolor > 0 & this.datcolor < 130)) //有指定 Custom 和合法設定再修改顏色
                {
                    foreach (DatlabModelAttribute m in models)
                    {
                        m.Color = this.datcolor;
                    }
                }

                this.Datalabel.SetCustomDatlab(models);


                if (this.ChartType == BarChartType.Cluster)
                {

                    cmnd.Append(this.Datalabel.GetCommand());
                }
                else
                {
                    /*
                     * 這裡是客製化stack bar chart label, 使用
                     */
                    List<double[]> datlabInfo = GetStackDatlabInfo(this.variables, ws);

                    /*
                     * 為各 model 的每個datlab 建立 offset 的值...
                     * 需要將其標準化..乘以-0.75*(dYMax-dYMin)/(sMax-sMin)
                     * dYMax, 需要計算 Title 高度
                     * dYMin, 需要計算 Label 高度
                     * 
                     */
                    double[] offsetInDataUnit;
                    double sMin = (this.Y_Scale.Min.Value != 1.23456E+30 ?
                        this.Y_Scale.Min.Value : (this.Y_Scale.GScale.SMIN < 0 ? this.Y_Scale.GScale.SMIN : 0));
                    double sMax = (this.Y_Scale.Max.Value != 1.23456E+30 ? this.Y_Scale.Max.Value : this.Y_Scale.GScale.SMAX);
                    double dYMax = 0.93;
                    double dYMin = 0.044;
                    Size sizeText = new Size(0, 0);
                    if (this.Annotation.Title != String.Empty)//計算 Title 高度
                    {
                        sizeText = TextRenderer.MeasureText((this.Annotation.Title == null ? "Bar-Chart" : this.Annotation.Title),
                            new Font(System.Drawing.SystemFonts.DialogFont.Name,
                                (float)this.Annotation.TitleFontSize * 100 / this.incrPercent, FontStyle.Bold));
                        dYMax = dYMax - (double)sizeText.Height / 384;
                    }

                    //if (this.X_Scale.AxLab.Label != String.Empty)//計算 X-Axis label 高度
                    //{
                    //    sizeText = TextRenderer.MeasureText((this.X_Scale.AxLab.Label == null ? "X-axis label" : this.X_Scale.AxLab.Label),
                    //        new Font(System.Drawing.SystemFonts.DialogFont.Name,
                    //            (float)this.X_Scale.AxLab.FontSize * (100 / this.incrPercent), FontStyle.Bold));
                    //    dYMin = dYMin + (double)sizeText.Height / 384;
                    //}

                    /*
                     * 計算 X-Axis tick 高度...這裡要小心計算有時有誤差..
                     * 1. Datetime 抓出來會是 full format..YYYY/MM/DD HH:MM:SS
                     * 2. 有些時後，tick label 的角度不是45度..
                     */
                    String[] ticklab;
                    dynamic tmp;
                    Size tmpSize = new Size(0, 0);
                    sizeText = new Size(0, 0);
                    if (this.TableArrangement == BarChartTableArrangement.RowsOuterMost)//labCol 的值就是 tick label
                    {
                        tmp = ws.Columns.Item(this.labvariable[0]).GetData();
                        foreach (dynamic s in tmp)
                        {
                            tmpSize = TextRenderer.MeasureText(s.ToString(), new Font(System.Drawing.SystemFonts.DialogFont.Name,
                                (float)this.X_Scale.Tick.FontSize * 100 / this.incrPercent, FontStyle.Regular));
                            if (tmpSize.Width > sizeText.Width) sizeText.Width = tmpSize.Width;
                        }
                    }
                    else//variable name 就是 tick label
                    {
                        ticklab = new string[this.variables.Count];
                        foreach (String s in this.variables)
                        {
                            tmpSize = TextRenderer.MeasureText(ws.Columns.Item(s).Label, new Font(System.Drawing.SystemFonts.DialogFont.Name,
                                (float)this.X_Scale.Tick.FontSize * 100 / this.incrPercent, FontStyle.Regular));
                            if (tmpSize.Width > sizeText.Width) sizeText.Width = tmpSize.Width;
                        }
                    }
                    //計算最後的 data region 下界
                    dYMin = dYMin + (double)sizeText.Width * Math.Abs(Math.Sin(Math.PI *
                            (this.X_Scale.Tick.TickAngle < 1.23456E+30 ? this.X_Scale.Tick.TickAngle : 45) / 180.0)) / 384;

                    //處理offset 
                    double k = (dYMax - dYMin) / (sMax - sMin);
                    for (int i = 0; i < models.Count; i++)
                    {
                        models[i].Label = datlabInfo[i];
                        offsetInDataUnit = new double[datlabInfo[i].Length];

                        for (int j = 0; j < offsetInDataUnit.Length; j++)
                        {
                            offsetInDataUnit[j] = -k * datlabInfo[i][j];
                        }
                        models[i].Offset = offsetInDataUnit;
                    }
                    cmnd.Append(this.Datalabel.GetCommand());
                }
            }

            cmnd.Append(this.Annotation.GetCommand());
            if (this.isSaveGraph)
            {
                cmnd.AppendLine(" GSAVE \"" + this.pathOfSaveGraph + "\";");
                cmnd.AppendLine("  JPEG;" + Environment.NewLine + "  REPL;");
            }

            return cmnd.ToString();
        }

        private void CreateChart()
        {
            /*
             * Run Minitab Command
             * 
             */

            StringBuilder mtbCmnd = new StringBuilder();
            StringBuilder exportString;

            mtbCmnd.AppendLine("TITLE" + Environment.NewLine + "BRIEF 0");
            String cmnd = GetCommand();
            if (this.GraphSize.Width != 576 || this.GraphSize.Height != 384)
            {
                mtbCmnd.Append(cmnd);
                mtbCmnd.AppendLine("GRAPH " + (double)this.GraphSize.Width / ((double)96 * this.incrPercent / 100) + " " +
                    (double)this.GraphSize.Height / ((double)96 * this.incrPercent / 100) + ".");
            }
            else
            {
                mtbCmnd.AppendLine(cmnd.Substring(0, cmnd.Length - Environment.NewLine.Length - 1) + ".");
            }
            //Console.Write(mtbCmnd.ToString());
            /*
             * 準備暫存檔，用於執行巨集
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
            StreamWriter sw;
            sw = new StreamWriter(path);
            sw.Write(mtbCmnd.ToString());
            sw.Close();
            int cmndStart = proj.Commands.Count;
            proj.ExecuteCommand("EXEC '" + path + "' 1", ws);
            if (this.isExportCmnd) ExportCommand(mtbCmnd.ToString(), this.pathOfExportCmnd, true);
            if (this.isCopyToClipboard) CopyToClipboard("CHART", proj, ws, cmndStart, proj.Commands.Count);
        }


        public void Run()
        {
            if (this.proj == null || this.ws == null)
            {
                MessageBox.Show("Minitab Project and Worksheet cannot be null", "Bar Chart", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (this.variables == null)
            {
                MessageBox.Show("Please input at least one variable", "Bar Chart", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (this.labvariable == null)
            {
                MessageBox.Show("Label variable cannot be null", "Bar Chart", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            CreateChart();

        }

        private List<double[]> GetStackDatlabInfo(List<String> cols, Mtb.Worksheet ws)
        {
            List<double[]> datlab = new List<double[]>();
            foreach (String col in cols)
            {
                double[] data = ws.Columns.Item(col).GetData();
                datlab.Add(data);
            }
            return datlab;
        }

    }
}
