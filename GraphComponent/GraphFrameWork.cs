using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MtbGraph.GraphComponent
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)] 
    public class GraphFrameWork : ICOMInterop_GFrameWork, IDisposable
    {
        protected Mtb.Project proj;
        protected Mtb.Worksheet ws;
        protected int incrPercent=100;
        public GraphFrameWork(Mtb.Project proj, Mtb.Worksheet ws)
        {
            this.proj = proj;
            this.ws = ws;
            //這裡是用於像 Bar chart 的 offset datlab 或 bar-line plot 計算都可以用到
            Form form = new Form();
            Graphics g = form.CreateGraphics();
            float dpiX;
            dpiX = g.DpiX;
            incrPercent = (dpiX == 96 ? 100 : (dpiX == 120 ? 125 : 150));
            GraphSize = new Size(576, 384);
        }
        /*
         * 為VB6 設計的建構子，因此 VB6使用時須自行設定 Minitab 參數
         */
        public GraphFrameWork()
        {
            //這裡是用於像 Bar chart 的 offset datlab 或 bar-line plot 計算都可以用到
            Form form = new Form();
            Graphics g = form.CreateGraphics();
            float dpiX;            
            dpiX = g.DpiX;
            incrPercent = (dpiX == 96 ? 100 : (dpiX == 120 ? 125 : 150));
            GraphSize = new Size(576, 384);
        }

        public void SetMtbEnvironment(Mtb.Project proj, Mtb.Worksheet ws)
        {
            this.proj = proj;
            this.ws = ws;
        }
        

        /************************************************************
         * 
         * 一般 IO 工具，一些方法需要設為 protected 讓 extend 的 class 可以使用
         * 
         *************************************************************/
        
        /***************************
         * 是否儲存圖形 
         */
        protected bool isSaveGraph = false;
        protected String pathOfSaveGraph = null;
        public void SaveGraph(bool b, String outputPath)
        {
            isSaveGraph = b;
            if (b)
            {
                if (String.IsNullOrEmpty(outputPath))
                {
                    isSaveGraph = false;
                }
                else
                {
                    Regex regEx = new Regex(@"^(?:[\w]\:|\\)(\\[^\\\/:\*?<>|]+)+\.(JPEG|JPG)$", RegexOptions.IgnoreCase);
                    pathOfSaveGraph = outputPath;
                    Match match = regEx.Match(pathOfSaveGraph);//check is a validate file extension
                    if (match.Success)
                    {
                        regEx = new Regex(@"^(?:[\w]\:|\\)(\\[^\\\/:\*?<>|]+)*\\", RegexOptions.IgnoreCase);
                        String rPath = regEx.Match(outputPath).Groups[0].Value.ToString();
                        if (!Directory.Exists(rPath))
                        {
                            Directory.CreateDirectory(rPath);
                        }
                    }
                    else
                    {
                        isSaveGraph = false;
                    }
                }
            }

        }


         /***************************
         * 輸出 Minitab 指令 
         */
        protected bool isExportCmnd = false;
        protected String pathOfExportCmnd;//用於紀錄 Command 的 Output 路徑
        public void SetExportCommand(bool b, String outputPath = null)
        {
            isExportCmnd = b;
            if (b)
            {
                if (String.IsNullOrEmpty(outputPath))
                {
                    isExportCmnd = false;
                }
                else
                {
                    // 判斷路徑和檔案格式是否合格
                    Regex regEx = new Regex(@"^(?:[\w]\:|\\)(\\[^\\\/:\*?<>|]+)+\.(txt|mtb)$", RegexOptions.IgnoreCase);
                    pathOfExportCmnd = outputPath;
                    Match match = regEx.Match(pathOfExportCmnd);//check is a validate file extension
                    if (match.Success)
                    {
                        regEx = new Regex(@"^(?:[\w]\:|\\)(\\[^\\\/:\*?<>|]+)*\\", RegexOptions.IgnoreCase);
                        String rPath = regEx.Match(outputPath).Groups[0].Value.ToString();
                        if (!Directory.Exists(rPath))
                        {
                            Directory.CreateDirectory(rPath);
                        }
                    }
                    else
                    {
                        isExportCmnd = false;
                    }
                }
            }

        }
        protected void ExportCommand(String inputStr, String outputPath, bool createNewFile = true)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(inputStr);
            FileStream fs;
            if (createNewFile)
            {
                fs = new FileStream(pathOfExportCmnd, FileMode.Create);
            }
            else
            {
                fs = new FileStream(pathOfExportCmnd, FileMode.OpenOrCreate);
            }
            fs.Close();
            StreamWriter sw;
            if (createNewFile)
            {
                sw = new StreamWriter(pathOfExportCmnd);
            }
            else
            {
                sw = new StreamWriter(pathOfExportCmnd, true);
            }
            sw.Write(sb.ToString());
            sw.Close();
        }
        
        
        /***************************
         * 複製圖形至剪貼簿
         */
        protected bool isCopyToClipboard = false;
        public void CopyGraphToClipboard(bool b)
        {
            isCopyToClipboard = b;
        }
        protected virtual void CopyToClipboard(String cmndName, Mtb.Project proj, Mtb.Worksheet ws, int start, int end)
        {
            Mtb.Commands cmnds;
            cmnds = proj.Commands;

            bool flag = false;
            for (int i = start+1; i <= end; i++)
            {
                if (cmnds.Item(i).Name == cmndName)
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
            //下段方式可以用於不輸入 start 和 end 的狀況
            //
            //foreach (Mtb.Command c in cmnds)
            //{
            //  ....
            //}
        }
        

        /***************************
         * 回收囉
         */ 
        public void Dispose()
        {
            this.proj = null;
            this.ws = null;
            GC.Collect();
        }


        protected int[] dFillColor = { 127, 28, 7, 58, 116, 78, 29, 45, 123, 35, 73, 8, 49, 57, 26 };
        protected int[] dLineColor = { 64, 8, 9, 12, 18, 34 };
        protected int[] dSymbType = { 6, 12, 16, 20, 23, 26, 29 };
        protected int[] dLineType = { 1, 2, 3, 4, 5 };
        public Size GraphSize { set; get; }

    }
}
