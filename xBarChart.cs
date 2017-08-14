using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mtb;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;

namespace MtbGraph
{
    public enum BarVOrder
    {
        ColumnOuterMost,
        RowOuterMost
    }
    [ClassInterface(ClassInterfaceType.None)] //自己設計接口
    public class xBarChart : MtbGraphFrame, IBarChart
    {
        public void CreateBarChart(Mtb.Project proj, Mtb.Worksheet ws,
            BarTypes bType = BarTypes.Stack, BarVOrder barVOrder = BarVOrder.RowOuterMost)
        {
            if (hasBar == 0)
            {
                return;
            }
            else
            {
                CreateChart(proj, ws, bType, barVOrder);
            }
        }
        private void CreateChart(Mtb.Project proj, Mtb.Worksheet ws,
            BarTypes bType = BarTypes.Stack, BarVOrder barVOrder = BarVOrder.RowOuterMost)
        {
            List<String> barColList = new List<String>();
            String barStr = "";
            String labStr = "";
            double dYMin = 0.044;
            double dYMax = 0.93;
            int cmndCnt = 0;

            if (hasBar == 0)
            {
                return;
            }
            else
            {
                barStr = String.Join(" ", this.barCols);
                barColList = da.GetMtbCols(this.barCols, ws);
            }

            if (hasLab == 1)
            {
                labStr = this.labCol[0];
            }

            /*
             * Declare variable for execute macro
             * 
             */
            String path;
            if (Environment.GetEnvironmentVariable("tmp").Equals(String.Empty))
            {
                path = Environment.GetEnvironmentVariable("tmp");
            }
            else { path = Environment.GetEnvironmentVariable("temp"); }
            path = path + "\\~macro.mtb";

            FileStream fs;
            fs = new FileStream(path, FileMode.Create);
            fs.Close();

            StreamWriter sw;
            StringBuilder mtbCmnd = new StringBuilder();
            MtbTools mtools = new MtbTools();
            String[] constStr = mtools.CreateVariableStrArray(ws, 5, MtbVarType.Constant);
            String[] colStr = mtools.CreateVariableStrArray(ws, 5, MtbVarType.Column);

            if (expCmnd) ExportCommand(String.Empty, cmndPath, true);
            if (isShowBDatlab & bType == BarTypes.Stack & barColList.Count > 1)
            {
                //Check Title
                Size sizeText = new Size(0, 0);
                if (this.title != String.Empty)
                {
                    sizeText = TextRenderer.MeasureText((this.title == null ? "Bar-Chart" : this.title), new Font("Segoe UI Semibold", (float)9.5, FontStyle.Bold));
                    dYMax = dYMax - ((double)sizeText.Height / d_gHeight);
                }

                sizeText = TextRenderer.MeasureText("Label Text", new Font("Segoe UI Semibold", 9, FontStyle.Bold));
                if (this.xLabel != String.Empty)
                {
                    dYMin = dYMin + ((double)sizeText.Height / d_gHeight);
                }


                double[] bDataArr = null;
                double[] tmpDataArr;
                if (barVOrder == BarVOrder.RowOuterMost)
                {
                    foreach (String str in barColList)
                    {
                        tmpDataArr = ws.Columns.Item(str).GetData();
                        if (bDataArr != null)
                        {
                            for (int i = 0; i < bDataArr.Length; i++)
                            {
                                bDataArr[i] = bDataArr[i] + tmpDataArr[i];
                            }
                        }
                        else
                        {
                            bDataArr = tmpDataArr;
                        }
                    }
                }
                else
                {
                    bDataArr = new double[barColList.Count];
                    for (int i = 0; i < barColList.Count; i++)
                    {
                        tmpDataArr = ws.Columns.Item(barColList[i]).GetData();
                        bDataArr[i] = tmpDataArr.Sum();
                    }
                }

                if (this.yRefValue != null)
                {
                    bDataArr = bDataArr.Union(yRefValue).ToArray();
                }

                double yMin = bDataArr.ToList().Min();
                double yMax = bDataArr.ToList().Max();

                //Console.WriteLine("yMin=" + yMin + ", yMax =" + yMax);

                if (Math.Abs(this.yScaleMin) < 1.23456E+30) yMin = this.yScaleMin;
                if (Math.Abs(this.yScaleMax) < 1.23456E+30) yMax = this.yScaleMax;

                mtbCmnd.Append("NOTITLE\r\nBRIEF 0\r\n");
                mtbCmnd.Append("GSCALE " + (this.yScaleMin == 1.23456E+30 ? (yMin >= 0 ? "0" : yMin.ToString()) : this.yScaleMin.ToString())
                    + " " + (this.yScaleMax == 1.23456E+30 ? yMax.ToString() : this.yScaleMax.ToString()) + ";\r\n");
                mtbCmnd.Append(" SMIN " + constStr[0] + ";\r\n  SMAX " + constStr[1] + ";\r\n");
                mtbCmnd.Append(" TMIN " + constStr[2] + ";\r\n  TMAX " + constStr[3] + ";\r\n");
                mtbCmnd.Append(" NTICK " + constStr[4] + ".\r\n");

                if (Math.Abs(this.yScaleMin) < 1.23456E+30)
                {
                    mtbCmnd.Append("COPY " + this.yScaleMin + " " + constStr[0] + "\r\n");
                }
                else if (yMin >= 0)
                {
                    mtbCmnd.Append("COPY 0 " + constStr[0] + "\r\n");
                }
                if (Math.Abs(this.yScaleMax) < 1.23456E+30) mtbCmnd.Append("COPY " + this.yScaleMax + " " + constStr[1] + "\r\n");
                mtbCmnd.Append("COPY " + constStr[0] + "-" + constStr[4] + " " + colStr[0] + "\r\n");
                if (hasLab == 1)
                {
                    mtbCmnd.Append("TEXT " + labCol[0] + " " + colStr[4] + "\r\n");
                }
                else
                {
                    mtbCmnd.Append("SET " + colStr[4] + "\r\n 1:" + ws.Columns.Item(barColList[0]).RowCount +
                        "\r\n END\r\n");
                    mtbCmnd.Append("TEXT " + colStr[4] + " " + colStr[4] + "\r\n");
                }

                /*
                 * Execute macro
                 */
                sw = new StreamWriter(path);
                sw.Write(mtbCmnd.ToString());
                sw.Close();
                proj.ExecuteCommand("EXEC '" + path + "' 1", ws);
                if (expCmnd) ExportCommand(mtbCmnd.ToString(), cmndPath, false);
            }



            mtbCmnd.Clear();
            if (isShowBDatlab & bType == BarTypes.Stack & barColList.Count > 1)
            {
                Size sizeText = new Size(0, 0);
                sizeText = GetStringSize(ws, colStr[4], this.d_LabFont);
                dYMin = dYMin + (double)sizeText.Width * Math.Abs(Math.Sin(Math.PI * (this.xLabelAngle < 1.23456E+30 ? this.xLabelAngle : 45) / 180.0)) / d_gHeight;
                double[] scaleInfo = ws.Columns.Item(colStr[0]).GetData();
                double k = (dYMax - dYMin) / (scaleInfo[1] - scaleInfo[0]);

                //Console.WriteLine("(dYMax - dYMin)/(yMax-yMin) = " + k + ", where dYMax=" + scaleInfo[1] + ", dYMin=" + scaleInfo[0]);
                mtbCmnd.Append("STACK " + barStr + " " + colStr[1] + ";\r\n SUBS " + colStr[2] + ".\r\n");
                mtbCmnd.Append("LET " + colStr[3] + "=" + colStr[1] + "*-1*" + k + "\r\n");
                sw = new StreamWriter(path);
                sw.Write(mtbCmnd.ToString());
                sw.Close();
                proj.ExecuteCommand("EXEC '" + path + "' 1", ws);
                if (expCmnd) ExportCommand(mtbCmnd.ToString(), cmndPath, false);
            }

            if (hasLab == 0) mtbCmnd.Append("SET " + colStr[4] + "\r\n 1:" + ws.Columns.Item(barColList[0]).RowCount +
                "\r\n END\r\n");
            mtbCmnd.Append("TITL\r\n");
            mtbCmnd.Append("CHART (" + barStr + ")*" + (hasLab == 0 ? colStr[4] : labCol[0]) + ";\r\n");
            //Get wtitle string
            mtbCmnd.Append(" WTIT \"Bar Chart Of " + GetTitleVariableString(ws, barColList) + "\";\r\n");
            if (this.gSave) mtbCmnd.Append(" GSAVE \"" + this.gPath + "\";\r\n  REPL;\r\n  JPEG;\r\n");

            mtbCmnd.Append(" SUMM;\r\n Overlay;\r\n  " + (barVOrder == BarVOrder.RowOuterMost ? "VLAST;\r\n" : "VFIRST;\r\n"));
            if (bType == BarTypes.Stack & barColList.Count > 1) mtbCmnd.Append(" STACK;\r\n");
            if (this.xLabelAngle < 1.23456E+30) mtbCmnd.Append(" SCALE 1;\r\n  ANGLE " + this.xLabelAngle + ";\r\n");

            if (this.xLabel == String.Empty)
            {
                mtbCmnd.Append(" AxLa 1;\r\n  LSHOW;\r\n");
            }
            else if (this.xLabel != null)
            {
                mtbCmnd.Append(" AxLa 1;\r\n  LABEL \"" + this.xLabel + "\";\r\n  LSHOW 1;" + Environment.NewLine);
            }

            if (bType == BarTypes.Cluster) mtbCmnd.Append(" TSHOW 1;\r\n");

            if (this.yLabel == String.Empty)
            {
                mtbCmnd.Append(" AxLa 2;\r\n  ADIS 0;\r\n");
            }
            else if (this.yLabel != null)
            {
                mtbCmnd.Append(" AxLa 2 \"" + this.yLabel + "\";\r\n");
            }

            if (this.yScaleMin != 1.23456E+30 || this.yScaleMax != 1.23456E+30)
            {
                mtbCmnd.Append(" SCALE 2;\r\n");
                if (Math.Abs(this.yScaleMin) < 1.23456E+30) mtbCmnd.Append("  MIN " + this.yScaleMin + ";\r\n");
                if (Math.Abs(this.yScaleMax) < 1.23456E+30) mtbCmnd.Append("  MAX " + this.yScaleMax + ";\r\n");
            }

            if (this.yRefValue != null)
            {
                mtbCmnd.Append(GetRefCmndString(this.yRefValue, this.yRefType, this.yRefColor));
            }

            if (bType == BarTypes.Cluster & barColList.Count > 1)
            {
                if (barVOrder == BarVOrder.ColumnOuterMost)
                {
                    mtbCmnd.Append(" BAR " + (hasLab == 0 ? colStr[4] : labCol[0]) + ";" + Environment.NewLine);
                }
                else
                {
                    mtbCmnd.Append(" BAR;" + Environment.NewLine + "  Vassign;" + Environment.NewLine);
                }
            }
            else
            {
                mtbCmnd.Append(" Bar;\r\n");
            }


            if (isShowBDatlab)
            {
                if (bType == BarTypes.Stack & barColList.Count > 1)
                {

                    double[] modelArr = ws.Columns.Item(colStr[2]).GetData();
                    double[] offsetArr = ws.Columns.Item(colStr[3]).GetData();
                    if (barVOrder == BarVOrder.RowOuterMost)
                    {
                        mtbCmnd.Append(" DATLAB " + colStr[1] + ";\r\n  PLAC 0 0;\r\n");
                        for (int i = 0; i < ws.Columns.Item(colStr[1]).RowCount; i++)
                        {
                            mtbCmnd.Append("  POSI " + (i + 1) + ";\r\n   MODEL " + modelArr[i] + ";\r\n");
                            mtbCmnd.Append("   OFFS 0 " + offsetArr[i] + ";\r\n   ENDP;\r\n");
                        }
                    }
                    else
                    {
                        int n = ws.Columns.Item(barColList[0]).RowCount;
                        double[] datArr = ws.Columns.Item(colStr[1]).GetData();
                        mtbCmnd.Append(" DATLAB;\r\n  PLAC 0 0;\r\n");
                        for (int i = 0; i < modelArr.Length; i++)
                        {
                            mtbCmnd.Append("  POSI " + ((i + 1) % n != 0 ? ((i + 1) % n).ToString() : n.ToString()) + " \"" + datArr[i] + "\";\r\n");
                            mtbCmnd.Append("   MODEL " + modelArr[i] + ";\r\n");
                            mtbCmnd.Append("   OFFS 0 " + offsetArr[i] + ";\r\n   ENDP;\r\n");
                        }
                    }
                }
                else
                {
                    mtbCmnd.Append(" DATLAB;\r\n");
                }
            }

            if (this.title == String.Empty)
            {
                mtbCmnd.Append(" NODT.");
            }
            else
            {
                mtbCmnd.Append(" TITLE " + (this.title == null ? ".\r\n" : "\"" + this.title + "\".\r\n"));
            }
            mtbCmnd.Append("NOTI\r\n");

            sw = new StreamWriter(path);
            sw.Write(mtbCmnd.ToString());
            sw.Close();
            proj.ExecuteCommand("EXEC '" + path + "' 1", ws);
            if (expCmnd) ExportCommand(mtbCmnd.ToString(), cmndPath, false);
            if (copyToClip) CopyToClipboard("CHART", proj, ws, cmndCnt + 1, proj.Commands.Count);

            mtbCmnd.Clear();
            mtbCmnd.Append("ERASE " + colStr[0] + "-" + colStr[colStr.Length - 1] + " " +
                constStr[0] + "-" + constStr[constStr.Length - 1] + "\r\n");
            mtbCmnd.Append("TITLE\r\nBRIEF 2\r\n");
            sw = new StreamWriter(path);
            sw.Write(mtbCmnd.ToString());
            sw.Close();
            proj.ExecuteCommand("EXEC '" + path + "' 1", ws);
        }


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
        private List<String> barCols;
        private List<String> labCol;
        private bool isShowBDatlab;
        private DialogAppraiser da = new DialogAppraiser();
        private int hasBar = 0;
        private int hasLab = 0;
    }
}
