using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mtb;
using System.IO;

namespace MtbGraph
{
    internal class SPCTrend
    {

        public String varCol { get; set; }
        public String subgp { get; set; }

        public void CreateSPCTrend(Mtb.Project proj, Mtb.Worksheet ws)
        {
            String path;
            if (Environment.GetEnvironmentVariable("tmp").Equals(String.Empty))
            {
                path = Environment.GetEnvironmentVariable("tmp");
            }
            else { path = Environment.GetEnvironmentVariable("temp"); }

            path = path + "\\~macro.mtb";

            StringBuilder mtbCmnd = new StringBuilder();
            mtbCmnd.Capacity = 49152;

            MtbTools mtbTool = new MtbTools();
            String[] colStr = mtbTool.CreateVariableStrArray(ws, 11, MtbVarType.Column);
            mtbCmnd.Append("NOTITLE\r\nBRIEF 0\r\n");
            try
            {
                int subgpSize;
                int l;
                int m;
                int k;
                subgpSize = Convert.ToInt32(subgp);
                l = ws.Columns.Item(varCol).RowCount;
                m = (int)Math.Floor((double)l / subgpSize);
                k = l % subgpSize;
                mtbCmnd.Append("SET " + colStr[0] + "\r\n");
                mtbCmnd.Append("1(1:" + m + ")" + subgpSize + (k > 0 ? "1(" + m + 1 + ")" + k : "") + "\r\n");
                mtbCmnd.Append("END\r\n");

            }
            catch (Exception e)
            {
                mtbCmnd.Append("LET " + colStr[0] + "=LAG(" + subgp + ",1)\r\n");
                mtbCmnd.Append("LET " + colStr[0] + "=" + colStr[0] + "<>" + subgp + "\r\n");
                mtbCmnd.Append("LET " + colStr[0] + "=PARS(" + colStr[0] + ")\r\n");
            }
            mtbCmnd.Append("");
            mtbCmnd.Append("");
            mtbCmnd.Append("TITLE\r\nBRIEF 2\r\n");

            FileStream fs = new FileStream(path, FileMode.Create);
            fs.Close();

            StreamWriter sw = new StreamWriter(path);
            sw.Write(mtbCmnd.ToString());
            sw.Close();

            proj.ExecuteCommand("EXEC '" + path + "' 1", ws);
            //"EXECUTE '" & path & "' 1", ws


        }
        public void Dispose()
        {
            GC.Collect();
        }
    }
}
