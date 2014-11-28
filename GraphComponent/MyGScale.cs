using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    public class MyGScale
    {

        public double SMIN { private set; get; }
        public double SMAX { private set; get; }
        public double NTICKS { set; get; }
        public double TMIN { set; get; }
        public double TMAX { set; get; }
        public double TINCREMENT { set; get; }
        public MyGScale()
        {
            Initialize();
        }
        public MyGScale(double min, double max)
        {
            Initialize();
            Run(min, max);

        }
        public MyGScale(double min, double max, Mtb.Project proj, Mtb.Worksheet ws)
        {
            Initialize();
            Run(min,max,proj,ws);
        }

        public void Initialize()
        {
            this.SMAX = 1.23456E+30;
            this.SMIN = 1.23456E+30;
            this.TMIN = 1.23456E+30;
            this.TMAX = 1.23456E+30;
            this.TINCREMENT = 1.23456E+30;            
        }

        public void Run(double min, double max)
        {
            double dMin = min;
            double dMax = max;

            if (min == max)
            {
                dMin = min * 0.99;
                dMax = max * 1.01;
            }
            else if (max < min)
            {
                dMax = min;
                dMin = max;
            }

            if (dMax > 0)
            {
                dMax = dMax + (dMax - dMin) * 0.01;
            }
            else if (dMax < 0)
            {
                dMax = Math.Min(dMax + (dMax - dMin) * 0.01, 0);
            }
            else
            {
                dMax = 0;
            }

            if (dMin > 0)
            {
                dMin = Math.Max(dMin - (dMax - dMin) * 0.01, 0);
            }
            else if (dMin < 0)
            {
                dMin = dMin - (dMax - dMin) * 0.01;
            }
            else
            {
                dMin = 0;
            }
            if (dMax == 0 & dMin == 0) dMax = 1;

            double dPower = Math.Log10(dMax - dMin);
            double d = Math.Pow(10, dPower - Math.Truncate(dPower));
            double dScale = d;
            Console.WriteLine("dScale is {0}", d);

            if (d <= 2.5 & d > 0)
            {
                dScale = 0.2;
            }
            else if (d > 2.5 & d <= 5)
            {
                dScale = 0.5;
            }
            else if (d > 5 & d <= 7.5)
            {
                dScale = 1;
            }
            else
            {
                dScale = 2;
            }

            dScale = dScale * Math.Pow(10, Math.Truncate(dPower));
            this.SMAX = dScale * (Math.Truncate(dMax / dScale) + 1);
            this.SMIN = dScale * Math.Truncate(dMin / dScale);
        }

        public void Run(double min, double max, Mtb.Project proj, Mtb.Worksheet ws)
        {
            MtbTools mtools = new MtbTools();
            int colcnt = ws.Columns.Count;
            String[] constStr = mtools.CreateVariableStrArray(ws, 6, MtbVarType.Constant);
            String[] colStr = mtools.CreateVariableStrArray(ws, 1, MtbVarType.Column);

            StringBuilder cmnd = new StringBuilder();
            cmnd.AppendLine("NOTITLE" + Environment.NewLine + "BRIEF 0");
            cmnd.AppendLine("GSCALE " + min + " " + max + ";");
            cmnd.AppendLine(" NTICK " + constStr[0] + ";");
            cmnd.AppendLine(" TMIN " + constStr[1] + ";");
            cmnd.AppendLine(" TMAX " + constStr[2] + ";");
            cmnd.AppendLine(" TINCREMENT " + constStr[3] + ";");
            cmnd.AppendLine(" SMIN " + constStr[4] + ";");
            cmnd.AppendLine(" SMAX " + constStr[5] + ".");
            cmnd.AppendLine("COPY " + constStr[0] + "-" + constStr[5] + " " + colStr[0]);
            cmnd.AppendLine("TITLE" + Environment.NewLine + "BRIEF 2");


            //建立巨集檔案
            String path;
            if (Environment.GetEnvironmentVariable("tmp").Equals(String.Empty))
                path = Environment.GetEnvironmentVariable("tmp");
            else
                path = Environment.GetEnvironmentVariable("temp");
            path = path + "\\~gscaleMacro.mtb";
            FileStream fs = new FileStream(path, FileMode.Create);
            fs.Close();
            StreamWriter sw = new StreamWriter(path);
            sw.Write(cmnd.ToString());
            sw.Close();

            //執行 Minitab 指令
            proj.ExecuteCommand("EXEC '" + path + "' 1", ws);

            double[] scaleInfo = ws.Columns.Item(colStr[0]).GetData();
            this.NTICKS = scaleInfo[0];
            this.TMIN = scaleInfo[1];
            this.TMAX = scaleInfo[2];
            this.TINCREMENT = scaleInfo[3];
            this.SMIN = scaleInfo[4];
            this.SMAX = scaleInfo[5];

            //清除資料                
            foreach (string s in constStr.Reverse())
                ws.Constants.Item(s).Delete();

            foreach (string s in colStr.Reverse())
                ws.Columns.Item(s).Delete();
        }

        public MyGScale Clone()
        {
            return (MyGScale)this.MemberwiseClone();
        }





    }
}
