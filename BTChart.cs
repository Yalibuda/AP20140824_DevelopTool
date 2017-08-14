using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;


namespace MtbGraph
{


    [ClassInterface(ClassInterfaceType.None)] //自己設計接口
    class BTChart : MtbGraphFrame//, IBTChart
    {

        public BTChart()
            : base()
        {

        }

        public void CreateBarTrendChart(Mtb.Project proj, Mtb.Worksheet ws, BarTypes btype)
        {
            /**
             * Varify input variable
             */
            try
            {
                if (String.Join(" ", barCols) == String.Empty)
                {
                    throw new ArgumentNullException("Please specify at least one bar variable", new Exception());
                }
                if (trendCols == null)
                {
                    trendCols = new List<String>();
                    trendCols.Add("");
                }
                if (targetCols == null)
                {
                    targetCols = new List<String>();
                    targetCols.Add("");
                }

                if (String.Join(" ", trendCols) == String.Empty & String.Join(" ", targetCols) == String.Empty)
                {
                    throw new ArgumentNullException("Please specify at least one line variable.", new Exception());
                }

            }
            catch (ArgumentNullException ae)
            {
                throw ae;
            }
            catch (Exception e)
            {
                throw e;
            }

            try
            {
                if (btype == MtbGraph.BarTypes.Stack)
                {
                    CreateStackBarTrendChart(proj, ws);
                }
                else if (btype == MtbGraph.BarTypes.Cluster)
                {
                    CreateClusterBarTrendChart(proj, ws);
                }
            }
            catch (Exception e)
            {
                throw e;
            }



        }
        private void CreateClusterBarTrendChart(Mtb.Project proj, Mtb.Worksheet ws)
        {

            int colCnt;
            int cmndCnt;

            colCnt = ws.Columns.Count;
            cmndCnt = proj.Commands.Count; //it is used to determine the start of search for the CopyToClipboard method

            String bars;
            String trends;
            String targets;
            List<String> barCols;
            List<String> trendCols;
            List<String> targetCols;

            bars = String.Join(" ", this.barCols);
            trends = String.Join(" ", this.trendCols);
            targets = String.Join(" ", this.targetCols);

            int nB = 0;
            int nT = 0;
            int nTg = 0;
            int bLen = 0;
            int tick = 0;

            barCols = da.GetMtbCols(this.barCols, ws);
            nB = barCols.Count();
            if (trends != String.Empty)
            {
                trendCols = da.GetMtbCols(this.trendCols, ws);
                nT = trendCols.Count;
            }
            else
            {
                trendCols = new List<String>();
                trendCols.Add("");
            }

            if (targets != String.Empty)
            {
                targetCols = da.GetMtbCols(this.targetCols, ws);
                nTg = targetCols.Count;
            }
            else
            {
                targetCols = new List<String>();
                targetCols.Add("");

            }
            bLen = ws.Columns.Item(barCols[0]).RowCount;

            tick = ((Double)nB / 2) > Math.Floor((Double)nB / 2) ? ((nB + 1) / 2) : (nB / 2);

            // Create minitab variables
            String[] colStr;
            String[] propCols;

            MtbTools mTool = new MtbTools();
            colStr = mTool.CreateVariableStrArray(ws, 4 + nT + nTg, MtbVarType.Column);
            propCols = mTool.CreateVariableStrArray(ws, 5, MtbVarType.Column);


            //Set name for temporary column
            for (int i = 0; i < colStr.Length; i++)
            {
                if (i == 3) { ws.Columns.Item(colStr[i]).Name = "~bar"; }
                else if (i > 3 & i <= 3 + nT) { ws.Columns.Item(colStr[i]).Name = "~trend_" + (i - 3); }
                else if (i > 3 + nT & i <= 3 + nT + nTg) { ws.Columns.Item(colStr[i]).Name = "~target_" + (i - 3 - nT); }
            }

            //Start build macros
            StringBuilder mtbCmnd = new StringBuilder();
            mtbCmnd.Append("NOTITLE\r\nBRIEF 0\r\n");
            //Check if specify decimal number of trend and target
            if (dNum < 100 & trends != String.Empty)
            {
                mtbCmnd.Append("FNUM " + trends + ";\r\n FIXED " + dNum + ".\r\n");
            }
            //Create empty column
            mtbCmnd.Append("NAME " + colStr[0] + " " + "\"~empty\"\r\n");
            mtbCmnd.Append("SET " + colStr[0] + "\r\n  " + bLen + "('*')\r\n" + "END\r\n");

            //Create a index as the x variable in scatter plot
            mtbCmnd.Append("SET " + colStr[1] + "\r\n" + "  1:" + (nB + 1) * bLen + "\r\nEND\r\n");
            //Create stacked bar column for cluster bar

            mtbCmnd.Append("ROWTOC " + bars + " " + colStr[0] + " " + colStr[3] + ";\r\n  CSUB " + colStr[2] + ".\r\n");

            String[] tmpArr = new String[nB + 1];
            for (int i = 0; i < tmpArr.Length; i++) tmpArr[i] = colStr[0];

            if (nT > 0)
            {
                for (int i = 0; i < trendCols.Count; i++)
                {
                    tmpArr[tick - 1] = trendCols[i];
                    mtbCmnd.Append("ROWTOC " + String.Join(" ", tmpArr) + " " + colStr[4 + i] + "\r\n");
                }
            }
            if (nTg > 0)
            {
                for (int i = 0; i < targetCols.Count; i++)
                {
                    tmpArr[tick - 1] = targetCols[i];
                    mtbCmnd.Append("ROWTOC " + String.Join(" ", tmpArr) + " " + colStr[4 + nT + i] + "\r\n");
                }
            }

            //Create properties column            
            mtbCmnd.Append("SET " + propCols[0] + "\r\n  " + nB + "(1) " + (nT + nTg) * nB + "(0)" + "\r\nEND\r\n");

            int[] propArray = new int[nB + nT + nTg];
            int m = 0;
            //Set project color
            for (int i = 0; i < propArray.Length; i++)
            {
                if (i < nB)
                {
                    m = i % this.dFillColor.Length;
                    propArray[i] = this.dFillColor[m];
                }
                else
                {
                    propArray[i] = 0;
                }
            }
            ws.Columns.Item(propCols[1]).SetData(propArray);
            //Set line type
            propArray = new int[1 + nT + nTg];
            for (int i = 0; i < propArray.Length; i++)
            {
                if (i > 0)
                {
                    if (i > nT & this.targetType != null)
                    {
                        m = (i - 1) % this.targetType.Count;
                        propArray[i] = this.targetType[m];
                    }
                    else
                    {
                        m = (i - 1) % this.dLineType.Length;
                        propArray[i] = this.dLineType[m];
                    }
                }
                else
                {
                    propArray[i] = 0;
                }
            }
            ws.Columns.Item(propCols[2]).SetData(propArray);

            //Set line color
            for (int i = 0; i < propArray.Length; i++)
            {
                if (i > 0)
                {
                    if (i > nT & this.targetColor != null)
                    {
                        m = (i - 1) % this.targetColor.Count;
                        propArray[i] = this.targetColor[m];
                    }
                    else
                    {
                        m = (i - 1) % this.dLineColor.Length;
                        propArray[i] = this.dLineColor[m];
                    }
                }
                else
                {
                    propArray[i] = 0;
                }
            }
            ws.Columns.Item(propCols[3]).SetData(propArray);

            //Set symbol type
            for (int i = 0; i < propArray.Length; i++)
            {
                if (i > 0 & i < 1 + nT)
                {
                    m = (i - 1) % this.dSymbType.Length;
                    propArray[i] = this.dSymbType[m];
                }
                else
                {
                    propArray[i] = 0;
                }
            }
            ws.Columns.Item(propCols[4]).SetData(propArray);


            //Eliminate non-ecessary value for projStack, and check lengths of all needed columns.
            mtbCmnd.Append("CODE (\"~empty\") \"\" " + colStr[2] + " " + colStr[2] + "\r\n");

            //Prepare tmp macro1
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


            //Check if export command
            StringBuilder expCmndStr = new StringBuilder();
            if (expCmnd)
            {
                expCmndStr.Append(mtbCmnd.ToString());

            }


            //Create Cluster Bar Chart
            mtbCmnd.Clear();
            trends = String.Empty;
            if (nT > 0) trends = (nT > 1) ? colStr[4] + "-" + colStr[3 + nT] : colStr[4];

            targets = String.Empty;
            if (nTg > 0) targets = (nTg > 1) ? colStr[4 + nT] + "-" + colStr[3 + nT + nTg] : colStr[4 + nT];

            int l = 0;
            String tmpStr = "Cluster Bar-Trend Chart";
            if (this.title == null) tmpStr = this.title; //Get Title
            l = ws.Columns.Item(colStr[3]).RowCount;
            mtbCmnd.Append("TITLE\r\n");
            mtbCmnd.Append("PLOT (" + colStr[3] + " " + trends + " " + targets + ")*" + colStr[1] + ";\r\n");
            mtbCmnd.Append(" WTITLE \"Bar-Trend Chart " + tmpStr + "\";\r\n");
            mtbCmnd.Append(" OVER;\r\n  NOEM;\r\n  NOMI;\r\n");
            mtbCmnd.Append(" PROJ " + colStr[2] + ";\r\n");
            mtbCmnd.Append("  TYPE " + propCols[0] + ";\r\n" + "  SIZE 7;\r\n  COLO " + propCols[1] + ";\r\n");
            mtbCmnd.Append(" SCALE 1;\r\n" + "  TICK " + tick + ":" + l + "/" + (nB + 1) + ";\r\n");
            mtbCmnd.Append("  Mini " + -0.5 + ";\r\n  Maxi " + (l + 0.5) + ";\r\n");

            //Set label
            List<String> labCol;
            String labs;
            try
            {
                labs = String.Join(" ", this.labCol);
            }
            catch { labs = String.Empty; }

            if (labs != String.Empty)
            {
                labCol = da.GetMtbCols(this.labCol, ws);
                mtbCmnd.Append("  LABEL " + labCol[0] + ";\r\n");
                if (this.xLabel == null)
                {
                    mtbCmnd.Append("  AxLa 1 \"" + ws.Columns.Item(labCol[0]).Label + "\";\r\n");
                }
                else
                {
                    mtbCmnd.Append("  AxLa 1 \"" + this.xLabel + "\";\r\n");
                }
            }
            else
            {
                mtbCmnd.Append(" AxLa 1 \"" + this.xLabel + "\";\r\n");
            }
            //Check Min-Max of y scale
            if (this.yScaleMin != 1.23456E+30 || this.yScaleMax != 1.23456E+30)
            {
                mtbCmnd.Append(" SCALE 2;\r\n");
                if (this.yScaleMin != 1.23456E+30) mtbCmnd.Append("  MINI " + this.yScaleMin + ";\r\n");
                if (this.yScaleMax != 1.23456E+30) mtbCmnd.Append("  MAXI " + this.yScaleMax + ";\r\n");
            }



            //Check secondary scale
            if (bScalePrimary != tScalePrimary)
            {
                mtbCmnd.Append(" SCALE 2;\r\n  SECS ");
                if (bScalePrimary == ScalePrimary.Secondary)
                {
                    mtbCmnd.Append(colStr[3] + ";\r\n");
                    if (this.secScaleMax != 1.23456E+30) mtbCmnd.Append("Maxi " + this.secScaleMax + ";\r\n");
                    if (this.secScaleMin != 1.23456E+30) mtbCmnd.Append("Mini " + this.secScaleMin + ";\r\n");
                    mtbCmnd.Append(" AxLab 2 \"" + (this.secsLabel == null ? "Bar" : this.secsLabel) + "\";\r\n  SECS;\r\n");
                    if (nT + nTg == 1)
                    {
                        mtbCmnd.Append("  AxLab 2 \"" + (this.yLabel == null ? (nT == 1 ?
                            ws.Columns.Item(trendCols[0]).Name : ws.Columns.Item(targetCols[0]).Name) : this.yLabel) + "\";\r\n");
                    }
                    else
                    {
                        mtbCmnd.Append("  AxLab 2 \"" + (this.yLabel == null ? "Line Data" : this.yLabel) + "\";\r\n");
                    }

                }
                if (tScalePrimary == ScalePrimary.Secondary)
                {
                    mtbCmnd.Append(trends + " " + targets + ";\r\n");
                    if (this.secScaleMax != 1.23456E+30) mtbCmnd.Append("Maxi " + this.secScaleMax + ";\r\n");
                    if (this.secScaleMin != 1.23456E+30) mtbCmnd.Append("Mini " + this.secScaleMin + ";\r\n");
                    mtbCmnd.Append("  AxLab 2 \"" + (this.yLabel == null ? "Bar" : this.yLabel) + "\";\r\n");
                    if (nT + nTg == 1)
                    {
                        mtbCmnd.Append("  AxLab 2 \"" + (this.secsLabel == null ? (nT == 1 ?
                            ws.Columns.Item(trendCols[0]).Name : ws.Columns.Item(targetCols[0]).Name) : this.secsLabel) + "\";\r\n  SECS;\r\n");
                    }
                    else
                    {
                        mtbCmnd.Append("  AxLab 2 \"" + (this.secsLabel == null ? "Line Data" : this.secsLabel) + "\";\r\n");
                    }
                }
            }

            mtbCmnd.Append(" CONN;\r\n  TYPE " + propCols[2] + ";\r\n  COLO " + propCols[3] + ";\r\n");
            mtbCmnd.Append(" SYMB;\r\n  TYPE " + propCols[4] + ";\r\n  COLO " + propCols[3] + ";\r\n");


            //Legend handler...
            StringBuilder sb = new StringBuilder();
            Size legendSize = new Size();
            if (nB + nT + nTg <= 3)
            {
                mtbCmnd.Append("SUBT \"\";\r\n");//增加放 Legend box 的空間
                if (this.title == String.Empty) mtbCmnd.Append("SUBT \"\";\r\n");
                String[] nmArr = new String[nB + nT + nTg];
                for (int i = 0; i < nmArr.Length; i++)
                {
                    if (i < nB)
                    {
                        nmArr[i] = ws.Columns.Item(barCols[i]).Label;
                    }
                    else if (i >= nB & i < nB + nT)
                    {
                        nmArr[i] = ws.Columns.Item(trendCols[i - nB]).Label;
                    }
                    else if (i >= nB + nT & i < nB + nT + nTg)
                    {
                        nmArr[i] = ws.Columns.Item(targetCols[i - nB - nT]).Label;
                    }
                }
                legendSize = GetLegendSize(nmArr, new Font("Segoe UI", 8, FontStyle.Regular));
                //隨著手動修改Graph size...legend box 位置可能會跑掉@@
                legendSize.Width = (int)((40 + legendSize.Width) * Math.Round((double)this.gHeight / 384, 1, MidpointRounding.ToEven));
                legendSize.Height = (int)((10 + legendSize.Height) * Math.Round((double)this.gHeight / 384, 1, MidpointRounding.ToEven));
                //legendSize.Width = 40 + legendSize.Width;
                //legendSize.Height = 10 + legendSize.Height;

            }
            sb.Append(" LEGE " + ((nB + nT + nTg <= 3) ? (0.9767 - (Double)legendSize.Width / this.gWidth) + " 0.9767 " + (0.9767 - (Double)legendSize.Height / this.gHeight) + " 0.9767;\r\n  PSIZE 8;\r\n" : ";\r\n"));
            sb.Append("  TFONT \"Segoe UI\";\r\n");
            sb.Append("  SECT 1;\r\n   CHHIDE;\r\n");
            sb.Append("   RHIDE " + (nB + 1) + ":" + (1 + nT + nTg) * nB + ";\r\n   CHIDE 2;\r\n");
            sb.Append("  SECT 2;\r\n   CHHIDE;\r\n   RHIDE 1;\r\n");
            //Get name of line variable...
            String[] lNm = new String[nT + nTg];
            for (int i = 0; i < lNm.Length; i++)
            {
                lNm[i] = (i < nT) ? ws.Columns.Item(trendCols[i]).Label : ws.Columns.Item(targetCols[i - nT]).Label;
            }
            //Set text in legend
            for (int i = 0; i < lNm.Length; i++)
            {
                sb.Append("   BTEXT " + (i + 2) + " 2 \"" + lNm[i] + "\";\r\n");
            }
            mtbCmnd.Append(sb.ToString());

            //Datlabel handler...

            if (isShowBDatlab || isShowTDatlab)
            {
                sb.Clear();
                sb.Append(" DatLab;\r\n  PLAC 0 0;\r\n");
                for (int i = 1; i <= 1 + nT + nTg; i++)
                {
                    if (i == 1)
                    {
                        if (!isShowBDatlab)
                        {
                            for (int j = 1; j <= (nB + 1) * bLen; j++) sb.Append("  POSI " + j + " \"\";\r\n   MODEL " + i + ";\r\n  ENDP;\r\n");
                        }
                    }
                    else if (i > 1 & i <= 1 + nT)
                    {
                        if (!isShowTDatlab)
                        {
                            for (int j = 1; j <= (nB + 1) * bLen; j++) sb.Append("  POSI " + j + " \"\";\r\n   MODEL " + i + ";\r\n  ENDP;\r\n");
                        }
                    }
                    else
                    {
                        for (int j = 1; j <= (nB + 1) * bLen; j++) sb.Append("  POSI " + j + " \"\";\r\n   MODEL " + i + ";\r\n  ENDP;\r\n");
                    }

                }
                mtbCmnd.Append(sb.ToString());
            }
            if (isShowTgDatlab & nTg > 0)
            {
                sb.Clear();
                for (int i = 0; i < targetCols.Count; i++)
                {
                    sb.Append(ws.Columns.Item(targetCols[i]).Label + ": " + GetTargetInfo(ws.Columns.Item(targetCols[i])));
                }
                mtbCmnd.Append(" FOOT \"" + sb.ToString() + "\";\r\n");
            }


            //Check if saving graph
            if (gSave)
            {
                mtbCmnd.Append(" GSAVE \"" + gPath + "\";\r\n  JPEG;\r\n REPL;\r\n");
            }
            //Check title
            if (this.title == null)
            {
                mtbCmnd.Append(" TITL \"Bar-Trend Chart\";\r\n");
            }
            else if (this.title != String.Empty)
            {
                mtbCmnd.Append(" TITL \"" + this.title + "\";\r\n");
            }
            if (this.gWidth != 576 || this.gHeight != 384)
            {
                mtbCmnd.Append("GRAPH " + (this.gWidth / 96) + " " + (this.gHeight / 96) + ";\r\n");
            }
            mtbCmnd.Append("NODT.\r\n");


            sw = new StreamWriter(path);
            sw.Write(mtbCmnd.ToString());
            sw.Close();
            proj.ExecuteCommand("EXEC '" + path + "' 1", ws);

            //Check if copy to clipboard
            if (copyToClip) CopyToClipboard("PLOT", proj, ws, cmndCnt + 1, proj.Commands.Count);

            //Check if export command             
            if (expCmnd) ExportCommand(expCmndStr.Append(mtbCmnd.ToString()).ToString(), cmndPath);

            //Delete variables
            mtbCmnd.Clear();
            mtbCmnd.Append("NOTI\r\nBRIEF 0\r\n");
            mtbCmnd.Append("ERASE " + colStr[0] + "-" + colStr[colStr.Length - 1] + " " +
                propCols[0] + "-" + propCols[propCols.Length - 1] + "\r\n");
            mtbCmnd.Append("TITLE\r\nBRIEF 2\r\n");
            sw = new StreamWriter(path);
            sw.Write(mtbCmnd.ToString());
            sw.Close();
            proj.ExecuteCommand("EXEC '" + path + "' 1", ws);

            //Delete temporary file...
            File.Delete(path);
        }
        private void CreateStackBarTrendChart(Mtb.Project proj, Mtb.Worksheet ws)
        {
            int colCnt;
            int cmndCnt;

            colCnt = ws.Columns.Count;
            cmndCnt = proj.Commands.Count; //it is used to determine the start of search for the CopyToClipboard method

            String bars;
            String trends;
            String targets;
            List<String> barCols;
            List<String> trendCols;
            List<String> targetCols;

            bars = String.Join(" ", this.barCols);
            trends = String.Join(" ", this.trendCols);
            targets = String.Join(" ", this.targetCols);

            int nB = 0;
            int nT = 0;
            int nTg = 0;
            int bLen = 0;

            barCols = da.GetMtbCols(this.barCols, ws);
            nB = barCols.Count();
            if (trends != String.Empty)
            {
                trendCols = da.GetMtbCols(this.trendCols, ws);
                nT = trendCols.Count;
            }
            else
            {
                trendCols = new List<String>();
                trendCols.Add("");
            }

            if (targets != String.Empty)
            {
                targetCols = da.GetMtbCols(this.targetCols, ws);
                nTg = targetCols.Count;
            }
            else
            {
                targetCols = new List<String>();
                targetCols.Add("");

            }
            bLen = ws.Columns.Item(barCols[0]).RowCount;

            // Create minitab variables
            String[] barNew;
            String[] propCols;
            MtbTools mTool = new MtbTools();
            barNew = mTool.CreateVariableStrArray(ws, nB, MtbVarType.Column);
            propCols = mTool.CreateVariableStrArray(ws, 6, MtbVarType.Column);
            for (int i = 0; i < barNew.Length; i++)
            {
                ws.Columns.Item(barNew[i]).Name = "~" + ws.Columns.Item(barCols[i]).Name + "~";
            }

            StringBuilder mtbCmnd = new StringBuilder();
            mtbCmnd.Append("NOTITLE\r\nBRIEF 0\r\n");
            //Check if specify decimal number of trend and target
            if (dNum < 100 & trends != String.Empty)
            {
                mtbCmnd.Append("FNUM " + trends + ";\r\n FIXED " + dNum + ".\r\n");
            }

            //Copy variable to a Matrix and Create an upper triangular matrix of 1's
            String[] mat;
            int m = 0;
            mat = mTool.CreateVariableStrArray(ws, 3, MtbVarType.Matrix);

            int[] arrayIndex = new int[nB * nB];
            for (int i = 0; i < arrayIndex.Length; i = i + nB)
            {
                for (int j = 0; j <= Math.Floor((Double)i / nB); j++) arrayIndex[i + j] = 1;

            }
            ws.Matrices.Item(mat[0]).SetData(arrayIndex, nB, nB);
            mtbCmnd.Append("COPY " + bars + " " + mat[1] + "\r\n");
            mtbCmnd.Append("MULT " + mat[1] + " " + mat[0] + " " + mat[2] + "\r\n");
            mtbCmnd.Append("COPY " + mat[2] + " " + barNew[0] + "-" + barNew[barNew.Length - 1] + "\r\n");
            mtbCmnd.Append("TSET " + propCols[0] + "\r\n 1(\"Bar\")" + bLen + "\r\n END\r\n");

            //Create properties columns
            int[] propArray = new int[nB + nT + nTg];
            mtbCmnd.Append("SET " + propCols[1] + "\r\n " + nB + "(1) " + (nT + nTg) + "(0)\r\n END\r\n");

            //Set proj color
            for (int i = 0; i < propArray.Length; i++)
            {
                if (i < nB)
                {
                    propArray[i] = dFillColor[((nB - 1) - i) % dFillColor.Length];
                }
                else
                {
                    propArray[i] = 0;
                }
            }
            ws.Columns.Item(propCols[2]).SetData(propArray);
            //Set conn Type
            for (int i = 0; i < propArray.Length; i++)
            {
                if (i >= nB)
                {
                    if (i >= nB + nT & this.targetType != null)
                    {
                        propArray[i] = this.targetType[(i - nB) % this.targetType.Count];
                    }
                    else
                    {
                        propArray[i] = this.dLineType[(i - nB) % this.dLineType.Length];
                    }
                }
                else
                {
                    propArray[i] = 0;
                }
            }
            ws.Columns.Item(propCols[3]).SetData(propArray);

            //Set conn line and symbol color, target and trend are use the same color setting
            for (int i = 0; i < propArray.Length; i++)
            {
                if (i >= nB)
                {
                    if (i >= nB + nT & this.targetColor != null)
                    {
                        propArray[i] = this.targetColor[(i - nB) % this.targetColor.Count];
                    }
                    else
                    {
                        propArray[i] = this.dLineColor[(i - nB) % this.dFillColor.Length];
                    }
                }
                else
                {
                    propArray[i] = 0;
                }
            }
            ws.Columns.Item(propCols[4]).SetData(propArray);

            //Set symbol type, target and trend are use the same color setting
            for (int i = 0; i < propArray.Length; i++)
            {
                if (i >= nB & i < nB + nT)
                {
                    propArray[i] = dSymbType[(i - nB) % dFillColor.Length];
                }
                else
                {
                    propArray[i] = 0;
                }
            }
            ws.Columns.Item(propCols[5]).SetData(propArray);
            mtbCmnd.Append("TITL\r\n");
            mtbCmnd.Append("TSPLOT " + ((nB == 1) ? barNew[0] : (barNew[barNew.Length - 1] + "-" + barNew[0])) + " " +
                trends + " " + targets + ";\r\n");
            mtbCmnd.Append(" WTITLE \"Stack Bar-Trend Chart\";\r\n");
            mtbCmnd.Append(" OVERl;\r\n NOEM;\r\n NOMI;\r\n");


            List<String> labCol;
            String labs;
            try
            {
                labs = String.Join(" ", this.labCol);
            }
            catch { labs = String.Empty; }
            if (!String.IsNullOrEmpty(labs))
            {
                labCol = da.GetMtbCols(this.labCol, ws);
                mtbCmnd.Append(" STAMP " + labCol[0] + ";\r\n");
                if (this.xLabel == null)
                {
                    mtbCmnd.Append("  AxLa 1 \"" + ws.Columns.Item(labCol[0]).Label + "\";\r\n");
                }
                else
                {
                    mtbCmnd.Append("  AxLa 1 \"" + this.xLabel + "\";\r\n");
                }
            }
            else
            {
                mtbCmnd.Append(" AxLa 1 \"\";\r\n");
            }
            mtbCmnd.Append(" PROJ " + propCols[0] + ";\r\n  TYPE " +
                propCols[1] + ";\r\n  COLOR " + propCols[2] + ";\r\n  Size 10;\r\n");
            mtbCmnd.Append(" CONN;\r\n  TYPE " + propCols[3] + ";\r\n  COLOR " + propCols[4] + ";\r\n");
            mtbCmnd.Append(" SYMB;\r\n  TYPE " + propCols[5] + ";\r\n  COLOR " + propCols[4] + ";\r\n");

            //Check Min-Max of y scale
            if (this.yScaleMin != 1.23456E+30 || this.yScaleMax != 1.23456E+30)
            {
                mtbCmnd.Append(" SCALE 2;\r\n");
                if (this.yScaleMin != 1.23456E+30) mtbCmnd.Append("  MINI " + this.yScaleMin + ";\r\n");
                if (this.yScaleMax != 1.23456E+30) mtbCmnd.Append("  MAXI " + this.yScaleMax + ";\r\n");
            }


            //Check secondary scale
            if (bScalePrimary != tScalePrimary)
            {
                mtbCmnd.Append(" SCAL 2;\r\n  SECS ");
                if (bScalePrimary == ScalePrimary.Secondary)
                {
                    mtbCmnd.Append(((nB == 1) ? barNew[0] : (barNew[0] + "-" + barNew[barNew.Length - 1])) + ";\r\n");
                    if (this.secScaleMax != 1.23456E+30) mtbCmnd.Append("Maxi " + this.secScaleMax + ";\r\n");
                    if (this.secScaleMin != 1.23456E+30) mtbCmnd.Append("Mini " + this.secScaleMin + ";\r\n");
                    mtbCmnd.Append(" AxLab 2 \"" + (this.secsLabel == null ? (nB == 1 ? ws.Columns.Item(barCols[0]).Label : "Bars") :
                        this.secsLabel) + "\";\r\n  SECS;\r\n");
                    if (nT + nTg == 1)
                    {
                        mtbCmnd.Append("  AxLab 2 \"" + (this.yLabel == null ? (nT == 1 ?
                            ws.Columns.Item(trendCols[0]).Name : ws.Columns.Item(targetCols[0]).Name) : this.yLabel) + "\";\r\n");
                    }
                    else
                    {
                        mtbCmnd.Append("  AxLab 2 \"" + (this.yLabel == null ? "Line Data" : this.yLabel) + "\";\r\n");
                    }

                }
                if (tScalePrimary == ScalePrimary.Secondary)
                {
                    mtbCmnd.Append(trends + " " + targets + ";\r\n");
                    if (this.secScaleMax != 1.23456E+30) mtbCmnd.Append("Maxi " + this.secScaleMax + ";\r\n");
                    if (this.secScaleMin != 1.23456E+30) mtbCmnd.Append("Mini " + this.secScaleMin + ";\r\n");
                    mtbCmnd.Append(" AxLab 2 \"" + (this.yLabel == null ? (nB == 1 ? ws.Columns.Item(barCols[0]).Label : "Bars") :
                            this.yLabel) + "\";\r\n ");
                    if (nT + nTg == 1)
                    {
                        mtbCmnd.Append("  AxLab 2 \"" + (this.secsLabel == null ? (nT == 1 ?
                            ws.Columns.Item(trendCols[0]).Name : ws.Columns.Item(targetCols[0]).Name) : this.secsLabel) + "\";\r\n  SECS;\r\n");
                    }
                    else
                    {
                        mtbCmnd.Append("  AxLab 2 \"" + (this.secsLabel == null ? "Line Data" : this.secsLabel) + "\";\r\n  SECS;\r\n");
                    }
                }

            }



            //Legend handler...
            StringBuilder sb = new StringBuilder();
            Size legendSize = new Size();
            if (nB + nT + nTg <= 3)
            {
                mtbCmnd.Append("SUBT \"\";\r\n");//增加放 Legend box 的空間
                if (this.title == String.Empty) mtbCmnd.Append("SUBT \"\";\r\n");
                String[] nmArr = new String[nB + nT + nTg];
                for (int i = 0; i < nmArr.Length; i++)
                {
                    if (i < nB)
                    {
                        nmArr[i] = ws.Columns.Item(barCols[i]).Label;
                    }
                    else if (i >= nB & i < nB + nT)
                    {
                        nmArr[i] = ws.Columns.Item(trendCols[i - nB]).Label;
                    }
                    else if (i >= nB + nT & i < nB + nT + nTg)
                    {
                        nmArr[i] = ws.Columns.Item(targetCols[i - nB - nT]).Label;
                    }
                }
                legendSize = GetLegendSize(nmArr, new Font("Segoe UI", 8, FontStyle.Regular));
                //隨著手動修改Graph size...legend box 位置可能會跑掉@@
                legendSize.Width = (int)((40 + legendSize.Width) * Math.Round((double)this.gHeight / 384, 1, MidpointRounding.ToEven));
                legendSize.Height = (int)((10 + legendSize.Height) * Math.Round((double)this.gHeight / 384, 1, MidpointRounding.ToEven));

            }
            sb.Append(" LEGE " + ((nB + nT + nTg <= 3) ? (0.9767 - ((Double)legendSize.Width / this.gWidth)) + " 0.9767 " + (0.9767 - ((Double)legendSize.Height / this.gHeight)) + " 0.9767;\r\n  PSIZE 8;\r\n" : ";\r\n"));
            sb.Append("  TFONT \"Segoe UI\";\r\n");
            sb.Append("  SECT 1;\r\n   CHHIDE;\r\n");
            sb.Append("   RHIDE " + (nB + 1) + ":" + (nB + nT + nTg) + ";\r\n");
            sb.Append("   CHIDE 3;\r\n");

            //Get name of bar variable
            String[] barNm = new String[nB];
            for (int i = 0; i < barCols.Count(); i++)
            {
                sb.Append("   BTEXT " + (i + 1) + " 2 \"" + ws.Columns.Item(barCols[barCols.Count - 1 - i]).Label + "\";\r\n");
            }
            sb.Append("  SECT 2;\r\n   CHHIDE;\r\n   RHIDE 1:" + nB + ";\r\n");
            mtbCmnd.Append(sb.ToString());

            //Data label handler...
            if (isShowBDatlab || isShowTDatlab)
            {
                sb.Clear();
                sb.Append(" DatLab;\r\n  PLAC 0 0;\r\n");
                for (int i = 1; i <= nB + nT + nTg; i++)
                {
                    if (i <= nB)
                    {
                        if (!isShowBDatlab)
                        {
                            for (int j = 1; j <= bLen; j++) sb.Append("  POSI " + j + " \"\";\r\n   MODEL " + i + ";\r\n  ENDP;\r\n");
                        }
                        else
                        {
                            for (int j = 1; j <= bLen; j++)
                            {
                                sb.Append("  POSI " + j + "\"" + Math.Round(ws.Columns.Item(barCols[nB - i]).GetData(j, 1), 2) +
                                    "\";\r\n   MODEL " + i + ";\r\n  ENDP;\r\n");
                            }
                        }
                    }
                    else if (i > nB & i <= nB + nT)
                    {
                        if (!isShowTDatlab)
                        {
                            for (int j = 1; j <= bLen; j++) sb.Append("  POSI " + j + " \"\";\r\n   MODEL " + i + ";\r\n  ENDP;\r\n");
                        }
                    }
                    else
                    {
                        for (int j = 1; j <= bLen; j++) sb.Append("  POSI " + j + " \"\";\r\n   MODEL " + i + ";\r\n  ENDP;\r\n");
                    }

                }
                mtbCmnd.Append(sb.ToString());
            }
            //Get target information
            if (isShowTgDatlab & nTg > 0)
            {
                sb.Clear();
                for (int i = 0; i < targetCols.Count; i++)
                {
                    sb.Append(ws.Columns.Item(targetCols[i]).Label + ": " + GetTargetInfo(ws.Columns.Item(targetCols[i])));
                }
                mtbCmnd.Append(" FOOT \"" + sb.ToString() + "\";\r\n");
            }

            //Check if saving graph
            if (gSave)
            {
                mtbCmnd.Append(" GSAVE \"" + gPath + "\";\r\n  JPEG;\r\n REPL;\r\n");
            }
            //Check title
            if (this.title == null)
            {
                mtbCmnd.Append(" TITL \"Bar-Trend Chart\";\r\n");
            }
            else if (this.title != String.Empty)
            {
                mtbCmnd.Append(" TITL \"" + this.title + "\";\r\n");
            }
            if (this.gWidth != 576 || this.gHeight != 384)
            {
                mtbCmnd.Append("GRAPH " + (this.gWidth / 96) + " " + (this.gHeight / 96) + ";\r\n");
            }


            mtbCmnd.Append("NODT.\r\n");
            mtbCmnd.Append("TITL\r\nBRIEF 2\r\n");


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

            //Check if copy to clipboard
            if (copyToClip) CopyToClipboard("TSPLOT", proj, ws, cmndCnt + 1, proj.Commands.Count);

            //Check if export command           
            if (expCmnd) ExportCommand(mtbCmnd.ToString(), cmndPath);

            //Delete variables...
            mtbCmnd.Clear();
            mtbCmnd.Append("NOTIL\r\nBRIEF 0\r\n");
            mtbCmnd.Append("ERASE " + barNew[0] + "-" + barNew[barNew.Length - 1] + " " +
                propCols[0] + "-" + propCols[propCols.Length - 1] + " " +
                mat[0] + "-" + mat[mat.Length - 1] + "\r\n");
            mtbCmnd.Append("TITL\r\nBRIEF 2\r\n");

            sw = new StreamWriter(path);
            sw.Write(mtbCmnd.ToString());
            sw.Close();
            proj.ExecuteCommand("EXEC '" + path + "' 1", ws);


            //Delete temporary file
            File.Delete(path);

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

        public void SetBarVariable(String inputStr)
        {
            barCols = da.GetMtbColInfo(inputStr);

        }

        public void SetBarDatLabel(bool b)
        {
            isShowBDatlab = b;
        }
        public void SetTrendDatLabel(bool b)
        {
            isShowTDatlab = b;
        }
        public void SetTrendDatLabel(bool b, int decimalNumber)
        {
            isShowTDatlab = b;
            dNum = decimalNumber;
        }
        public void SetTargetDatLabel(bool b)
        {
            isShowTgDatlab = b;
        }
        public void SetLabelVariable(String inputStr)
        {
            labCol = da.GetMtbColInfo(inputStr);
        }
        public void SetScalePrimary(ScalePrimary barScale, ScalePrimary lineScale)
        {
            bScalePrimary = barScale;
            tScalePrimary = lineScale;
        }
        public void SetTrendVariable(String inputStr)
        {
            trendCols = da.GetMtbColInfo(inputStr);
        }

        public void SetTargetVariable(String inputStr)
        {
            targetCols = da.GetMtbColInfo(inputStr);
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

        private Size GetLegendSize(String[] strArr, Font f)
        {
            List<Size> size = new List<Size>();
            for (int i = 0; i < strArr.Length; i++)
            {
                size.Add(TextRenderer.MeasureText(strArr[i], f));
            }
            Size maxSize = new Size();
            for (int i = 0; i < size.Count; i++)
            {
                if (i == 0)
                {
                    maxSize.Width = size[i].Width;
                    maxSize.Height = size[i].Height;
                }
                else
                {
                    if (size[i].Width > maxSize.Width) maxSize.Width = size[i].Width;
                    //if (size[i].Height > maxSize.Height) maxSize.Height = size[i].Height;
                    maxSize.Height = maxSize.Height + size[i].Height;
                }

            }
            return maxSize;


        }


        private List<String> barCols;
        private ScalePrimary bScalePrimary;
        private List<String> labCol;
        private int datLabDig;
        private bool isShowBDatlab;
        private bool isShowTDatlab;
        private bool isShowTgDatlab;
        private List<String> trendCols;
        private List<String> targetCols;
        private String mTitle = "Bar-Trend Chart";
        private ScalePrimary tScalePrimary;
        private DialogAppraiser da = new DialogAppraiser();
        private List<int> targetType = null;
        private List<int> targetColor = null;
        private int dNum = 100;



    }
}
