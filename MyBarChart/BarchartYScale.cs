using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MtbGraph.GraphComponent;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;

namespace MtbGraph.MyBarChart
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class BarchartYScale : ContinuousScale, ICOMInterop_BarchartYScale
    {        
        public BarchartYScale(ScaleType scale_axis)
            : base(scale_axis)
        {
            if (scale_axis != ScaleType.Y_axis) return;
            this.scale_axis = scale_axis;
            this.ChartType = BarChartType.Stack;
            this.TableArrangement = BarChartTableArrangement.RowsOuterMost;
        }
        public BarChartType ChartType { set; get; }
        public BarChartTableArrangement TableArrangement { set; get; }

        protected override void CalculateTickValue(List<string> varCols, Mtb.Worksheet ws, Mtb.Project proj)
        {
            if (this.ChartType == BarChartType.Cluster)
            {
                base.CalculateTickValue(varCols, ws, proj);
            }
            else if (this.ChartType == BarChartType.Stack)
            {
                if (ws == null || proj == null) return;

                List<double> datas = new List<double>();
                double[] data = new double[0];
                double min;
                double max;
                if (this.TableArrangement == BarChartTableArrangement.ColsOuterMost)
                {
                    foreach (String col in varCols)
                    {
                        data = ws.Columns.Item(col).GetData();
                        datas.Add(data.Sum());
                    }
                    min = datas.Min();
                    max = datas.Max();                    
                }
                else
                {
                    double[] tmp;
                    int rows = ws.Columns.Item(varCols[0]).RowCount;
                    int cols = varCols.Count;
                    foreach (String col in varCols)
                    {
                        tmp = ws.Columns.Item(col).GetData();
                        data = data.Concat<double>(tmp).ToArray();
                    }
                    Matrix<double> mat = new DenseMatrix(rows, cols, data);
                    Matrix<double> iden = Matrix<double>.Build.Dense(cols, 1, 1);
                    Matrix<double> stackvalue = mat * iden;
                    min = stackvalue.ToColumnWiseArray().Min();
                    max = stackvalue.ToColumnWiseArray().Max();
                }
                this.GScale.Run(min, max, proj, ws);


            }


        }
    }
}
