using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MtbGraph.GraphComponent;
using MtbGraph.Tool;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections;

namespace MtbGraph.MyTrend
{
    [ClassInterface(ClassInterfaceType.None)]
    public class Trend : GraphFrameWork, ICOMInterop_Trend
    {

        public Line Line { set; get; }
        public CategoricalScale X_Scale { set; get; }
        public ContinuousScale Y_Scale { set; get; }
        public SimpleLegend LegendBox { set; get; }
        public Annotation Annotation { set; get; }
        public Datlab Datalabel { get; set; }
        public TargetAttribute TargetAttribute { set; get; }
        public Trend()
            : base()
        {
            this.Line = new Line();
            this.X_Scale = new CategoricalScale(ScaleType.X_axis);
            this.Y_Scale = new ContinuousScale(ScaleType.Y_axis);
            this.Annotation = new Annotation();
            this.Datalabel = new Datlab();
            this.LegendBox = new SimpleLegend();
            this.TargetAttribute = new TargetAttribute();
        }

        public Trend(Mtb.Project proj, Mtb.Worksheet ws)
            : base(proj, ws)
        {
            this.Line = new Line();
            this.X_Scale = new CategoricalScale(ScaleType.X_axis);
            this.Y_Scale = new ContinuousScale(ScaleType.Y_axis);
            this.Annotation = new Annotation();
            this.Datalabel = new Datlab();
            this.LegendBox = new SimpleLegend();
            this.TargetAttribute = new TargetAttribute();
        }

        //設定 Trend 欄位
        private List<String> variables = null;
        private List<String> secvariables = null;
        private MtbTools mtools = new MtbTools();
        public void SetVariable(ref object variables, ScaleType scaletype = ScaleType.Y_axis)
        {
            switch (scaletype)
            {
                case ScaleType.Y_axis:
                    this.variables = mtools.TransObjToMtbColList(variables, ws);
                    break;
                case ScaleType.Secondary_Y_axis:
                    this.secvariables = mtools.TransObjToMtbColList(variables, ws);
                    break;
                case ScaleType.X_axis:
                    throw new ArgumentException("Invalid input variable scale, Y or Secondary Y was expected");
                    return;
            }
        }

        //設定 Label 欄位
        private List<String> labvariable = null;
        public void SetLabelVarible(ref object variables)
        {
            labvariable = mtools.TransObjToMtbColList(variables, ws);
        }

        //設定 Target 欄位
        private List<String> targets = null;
        private List<String> sectargets = null;
        public void SetTargetVariable(ref object targets, ScaleType scaletype = ScaleType.Y_axis)
        {
            switch (scaletype)
            {
                case ScaleType.Y_axis:
                    this.targets = mtools.TransObjToMtbColList(targets, ws);
                    break;
                case ScaleType.Secondary_Y_axis:
                    this.sectargets = mtools.TransObjToMtbColList(targets, ws);
                    break;
                case ScaleType.X_axis:
                    throw new ArgumentException("Invalid input variable scale, Y or Secondary Y was expected.");
                    return;
            }
        }

        //設定 Group 欄位(只給開放一個)
        private List<String> groupvariable = null;
        public void SetGroupVariable(String column)
        {
            this.groupvariable = mtools.TransObjToMtbColList(column, ws);
        }

        private List<int> targetColor = null;
        public void SetTargetColor(ref Object colors)
        {
            Type t = colors.GetType();
            List<int> list = new List<int>();
            if (t.IsArray)
            {
                try
                {
                    System.Collections.IEnumerable enumerable = colors as System.Collections.IEnumerable;
                    foreach (object o in enumerable)
                    {
                        list.Add(Convert.ToInt16(o.ToString()));
                    }
                    targetColor = list;
                }
                catch
                {
                    //throw new ArgumentException("Invalid input of target color");
                }
            }
            else
            {
                list.Add(Convert.ToInt16(colors.ToString()));
                targetColor = list;
            }
        }


        private List<int> targetType = null;
        public void SetTargetType(ref Object linetype)
        {
            Type t = linetype.GetType();
            List<int> list = new List<int>();
            Console.WriteLine("The type of element " + Type.GetTypeCode(t));
            if (t.IsArray)
            {
                try
                {
                    System.Collections.IEnumerable enumerable = linetype as System.Collections.IEnumerable;
                    foreach (object o in enumerable)
                    {
                        list.Add(Convert.ToInt16(o.ToString()));
                    }
                    targetType = list;
                }
                catch
                {
                    //throw new ArgumentException("Invalid input of type ar target connect line");
                }

            }
            else
            {
                //throw new ArgumentException("Invalid input type of target connect line type");
                list.Add(Convert.ToInt16(linetype.ToString()));
                targetColor = list;
            }

        }

        public String GetCommand()
        {

            /*
             * 整理出 Primary 和 Secondary variable 有哪些，繪製時需要
             * 1. 所有變數，排列方式為 primary, secondary, primary_target, secondary_target
             * 2. 主軸變數(設定Scale用)
             * 3. 副軸變數(設定Scale用)
             * 
             * 考慮是否需要特別做出 Variable 和 Target 的 List...因為 Symbol 用數量即可決定
             */
            if (this.variables == null) return @"#Please input at least one variable."; //這部分應該在此方法外就做完

            //Variable 部分
            List<String> varCols = this.variables;
            if (this.secvariables != null) varCols = varCols.Concat(this.secvariables).ToList();

            //Target 部分
            List<String> targetCols = new List<string>();
            if (this.targets != null) targetCols = targetCols.Concat(this.targets).ToList();
            if (this.sectargets != null) targetCols = targetCols.Concat(this.sectargets).ToList();

            //所有變數
            List<String> allCols = varCols;
            if (targetCols.Count > 0) allCols = allCols.Concat(targetCols).ToList();
            Object obj = allCols.ToArray();
            this.X_Scale.SetScaleVariable(ref obj, ws); //這裡要Setvariable的原因是為了要讓 X 軸的 tick 數可被控制(抓資料數)

            //主軸變數
            List<String> prmyvarCols = this.variables;
            if (this.targets != null) prmyvarCols = prmyvarCols.Concat(this.targets).ToList();
            obj = prmyvarCols.ToArray();
            this.Y_Scale.SetScaleVariable(ref obj, ws, proj);

            //副軸變數
            List<String> secvarCols = new List<string>();
            if (this.secvariables != null) secvarCols = secvarCols.Concat(this.secvariables).ToList();
            if (this.sectargets != null) secvarCols = secvarCols.Concat(this.sectargets).ToList();
            if (secvarCols.Count > 0)
                this.Y_Scale.SecsScale.SetScaleVariable(ref obj, ws, proj);


            StringBuilder cmnd = new StringBuilder();
            cmnd.AppendLine("TSPLOT " + String.Join(" ", allCols.ToArray()) + ";");
            cmnd.AppendLine(" OVER;");
            if (this.labvariable != null) cmnd.AppendLine(" STAMP " + String.Join(" ", this.labvariable) + ";");
            cmnd.Append(this.X_Scale.GetCommand());
            cmnd.Append(this.Y_Scale.GetCommand());

            int trendCount = this.variables.Count + (this.secvariables == null ? 0 : this.secvariables.Count);
            int targetCount = (this.targets == null ? 0 : this.targets.Count) + (this.sectargets == null ? 0 : this.sectargets.Count);
            int[] array1;
            int[] array2;

            /*
             * 當有 Target 時，需要對變數修改 Symbol，甚至是顏色
             */
            if (targetCount > 0)
            {
                //調整 Symbol type
                array1 = new int[trendCount];
                array2 = new int[targetCount];
                if (this.Line.Symbols.GetTypes() != null)
                {
                    //需要檢查 symbol type 是 array 或是單一值...然後塞回 array1 中
                    Type t = this.Line.Symbols.GetTypes().GetType();
                    if (t.IsArray)
                    {
                        IEnumerable tmp = this.Line.Symbols.GetTypes() as IEnumerable;
                        List<int> typeArray = new List<int>();
                        foreach (object o in tmp) typeArray.Add(Convert.ToInt16(o));
                        for (int i = 0; i < trendCount; i++)
                        {
                            array1[i] = typeArray[i % typeArray.Count];
                        }
                    }
                    else
                    {
                        int stypes = Convert.ToInt16(this.Line.Symbols.GetTypes());
                        for (int i = 0; i < trendCount; i++) array1[i] = stypes;
                    }

                }
                else
                {
                    for (int i = 0; i < trendCount; i++) array1[i] = dSymbType[i % this.dSymbType.Length];
                }

                for (int i = 0; i < targetCount; i++) array2[i] = 0;
                obj = array1.Concat(array2).ToArray();
                this.Line.Symbols.SetType(ref obj);

                /* 調整 Connectline color 和 type，當 Target 有指定顏色或類型時...
                 * 不用調整 Symbol 顏色..因為 Target 沒 Symbol.../_\
                 * 
                 * 未來可考慮改用 TargetAttribute物件的屬性...即將該類別內的SetTargetcolor 這些
                 * 方法拿掉，這裡就改用判斷 TargetAttribute 處理，包含將 dynamic 轉為 array 再丟
                 * 到 Line 裡面
                 */

                if (targetColor != null)
                {
                    for (int i = 0; i < trendCount; i++) array1[i] = dLineColor[i % this.dLineColor.Length];
                    for (int i = 0; i < targetCount; i++) array2[i] = targetColor[i % targetColor.Count];
                    obj = array1.Concat(array2).ToArray();
                    this.Line.Connectlines.SetColor(ref obj);
                }

                if (targetType != null)
                {
                    for (int i = 0; i < trendCount; i++) array1[i] = dLineType[i % this.dLineType.Length];
                    for (int i = 0; i < targetCount; i++) array2[i] = targetType[i % targetType.Count];
                    obj = array1.Concat(array2).ToArray();
                    this.Line.Connectlines.SetType(ref obj);
                }

            }
            else
            {
                //調整 Symbol type
                array1 = new int[trendCount];
                array2 = new int[targetCount];
                if (this.Line.Symbols.GetTypes() != null)
                {
                    //需要檢查 symbol type 是 array 或是單一值...然後塞回 array1 中
                    Type t = this.Line.Symbols.GetTypes().GetType();
                    if (t.IsArray)
                    {
                        IEnumerable tmp = this.Line.Symbols.GetTypes() as IEnumerable;
                        List<int> typeArray = new List<int>();
                        foreach (object o in tmp) typeArray.Add(Convert.ToInt16(o));
                        for (int i = 0; i < trendCount; i++)
                        {
                            array1[i] = typeArray[i % typeArray.Count];
                        }
                    }
                    else
                    {
                        int stypes = Convert.ToInt16(this.Line.Symbols.GetTypes());
                        for (int i = 0; i < trendCount; i++) array1[i] = stypes;
                    }
                    obj = array1;
                    this.Line.Symbols.SetType(ref obj);
                }

            }
            cmnd.Append(this.Line.GetCommand());
            if (this.Datalabel.Show)
            {
                Datalabel.LabelType = DatlabType.Value;
                DatlabModelAttribute model;
                List<DatlabModelAttribute> models = new List<DatlabModelAttribute>();
                if (this.Datalabel.Color == DatlabColor.Custom) //有指定 Custom 再修改顏色
                {
                    for (int i = 0; i < trendCount; i++)
                    {
                        model = new DatlabModelAttribute();
                        model.ModelIndex = i + 1;
                        model.Color = this.dLineColor[i % this.dLineColor.Length];
                        model.Start = 1;
                        model.End = ws.Columns.Item(varCols[i]).RowCount;
                        models.Add(model);
                    }
                    this.Datalabel.SetCustomDatlab(models);
                }

                if (targetCount > 0)
                {
                    models = new List<DatlabModelAttribute>();
                    for (int i = 0; i < targetCount; i++)
                    {
                        model = new DatlabModelAttribute();
                        model.ModelIndex = trendCount + i + 1;
                        model.Start = 1;
                        model.End = ws.Columns.Item(targetCols[i]).RowCount;
                        models.Add(model);
                    }
                    this.Datalabel.SetDatlabInvisible(models);
                }
                cmnd.Append(Datalabel.GetCommand());

                //處理 Target 的註記
                if (targetCount > 0)
                {
                    StringBuilder fnotes = new StringBuilder();
                    for (int i = 0; i < targetCount; i++)
                    {
                        fnotes.Append(ws.Columns.Item(targetCols[i]).Label + ": " + GetTargetInfo(targetCols[i], ws) + "; ");
                    }
                    this.Annotation.AddFootnote(fnotes.ToString());
                }

            }

            if (this.LegendBox.Show)
            {
                this.LegendBox.GraphSize = this.GraphSize;
                if ((this.LegendBox.HideHead == true & allCols.Count <= 3) ||
                    (this.LegendBox.HideHead == false & allCols.Count <= 2))
                    this.LegendBox.Location = Location.RightTop;
                if (this.LegendBox.Location != Location.Auto)
                {
                    String[] colname = new String[allCols.Count];
                    for (int i = 0; i < allCols.Count; i++)
                    {
                        colname[i] = ws.Columns.Item(allCols[i]).Label;
                    }
                    this.LegendBox.SetVariables(ref colname);
                }
                cmnd.Append(LegendBox.GetCommand());
            }
            cmnd.Append(this.Annotation.GetCommand());

            if (this.isSaveGraph)
            {
                cmnd.AppendLine(" GSAVE \"" + this.pathOfSaveGraph + "\";");
                cmnd.AppendLine("  JPEG;" + Environment.NewLine + "  REPL;");
            }

            return cmnd.ToString();
        }

        private String GetTargetInfo(String variables, Mtb.Worksheet ws)
        {
            Mtb.Column col = ws.Columns.Item(variables);
            Mtb.MtbDataTypes t = col.DataType;
            dynamic data = col.GetData();

            if (t == Mtb.MtbDataTypes.DataUnassigned || t == Mtb.MtbDataTypes.Text) return null;

            StringBuilder info = new StringBuilder();
            for (int i = 0; i < data.Length; i++)
            {
                if (i == 0)
                {
                    info.Append(data[i]);
                }
                else
                {
                    if (data[i] != data[i - 1])
                    {
                        info.Append(", " + data[i]);
                    }
                }
            }
            return info.ToString();
        }


        public void StackedDataTrend()
        {
            /*
             * 首先確認下列事項才能用 Stacked data 畫 TSPLOT
             * 1. 每一個 unstack 的 label 的 distinct row number = row number  
             * 
             */
            dynamic data = this.ws.Columns.Item(this.variables[0]).GetData();
            dynamic subs = this.ws.Columns.Item(this.groupvariable[0]).GetData();
            StackData stackdata = new StackData(data, subs);
            var unstackdata =
                from m in stackdata.Data
                group m by m.Subscript into gp
                select new { ID = gp.Key, Value = gp };

            data = this.ws.Columns.Item(this.labvariable[0]).GetData();
            stackdata = new StackData(data, subs);
            var unstacklab =
                from m in stackdata.Data
                group m by m.Subscript into gp
                select new { ID = gp.Key, Value = gp };

            //tally label column            
            IEnumerable enumerable = data as IEnumerable;
            List<dynamic> la = new List<dynamic>();
            foreach (object o in enumerable)
            {
                la.Add(o);
            }
            dynamic tallyLabel = la.Distinct().ToArray();


            //建立新工作表
            Mtb.Worksheet unstackWs = this.proj.Worksheets.Add(1);
            unstackWs.Name = "Summary_" + this.ws.Name + "_tmp";
            unstackWs.MakeWorksheetActive(0);

            //新增 k+1 個欄位...k+1 個放 unstack 後的 data ，剩下一個是放 tally 後的 label 資料
            MtbTools mtbTools = new MtbTools();
            String[] summaryColArray = mtbTools.CreateVariableStrArray(unstackWs, unstackdata.Count() + 1, MtbVarType.Column);

            unstackWs.Columns.Item(summaryColArray[0]).SetData(tallyLabel);
            unstackWs.Columns.Item(summaryColArray[0]).Name = ws.Columns.Item(this.labvariable[0]).Name;

            Mtb.Column col;
            List<dynamic> groupedData = new List<dynamic>();

            /*
             * Convert data ...
             * Create source table，這裡需要確保 unstacklabel 中沒有重複的 item (還沒做)...
             */
            Dictionary<dynamic, dynamic> mydictionarySource;
            for (int i = 0; i < unstacklab.Count(); i++)
            {
                col = unstackWs.Columns.Item(summaryColArray[i + 1]);
                col.Name = unstackdata.ElementAt(i).ID.ToString();
                mydictionarySource = new Dictionary<dynamic, dynamic>();
                for (int j = 0; j < unstacklab.ElementAt(i).Value.Count(); j++)
                {
                    mydictionarySource.Add(unstacklab.ElementAt(i).Value.ElementAt(j).Data,
                        unstackdata.ElementAt(i).Value.ElementAt(j).Data);
                }
                groupedData = new List<object>();
                foreach (dynamic item in tallyLabel)
                {
                    try { groupedData.Add(mydictionarySource[item]); }
                    catch (KeyNotFoundException e)
                    {
                        groupedData.Add(1.23456E+30);
                    }
                }
                col.SetData(groupedData.ToArray());
            }


            /*
             * 如果有 Target，則先以 label 取 distinct...一個 label 
             * 只能有一個值，如果有兩個則丟出訊息
             * 
             */
            List<String> invalidTargetCol = null;
            List<String> validTargetCol = null;
            if (this.targets != null)
            {
                /*
                 * 逐條 Target 檢驗，將合格的 Target 名稱放入 validTargetCol 清單中，同
                 * 時也記錄不合格的 Target 名稱。
                 */

                subs = ws.Columns.Item(this.labvariable[0]).GetData();
                bool flag;
                validTargetCol = new List<string>();
                invalidTargetCol = new List<string>();

                foreach (dynamic tg in this.targets)
                {
                    col = ws.Columns.Item(tg);
                    data = col.GetData();
                    stackdata = new StackData(data, subs);
                    groupedData = new List<dynamic>();
                    var unstackTg =
                        from m in stackdata.Data
                        group m by m.Subscript into gp
                        select new
                        {
                            ID = gp.Key,
                            Value = gp.First(),
                            Value_Max = gp.Max(t => t.Data),
                            Value_Min = gp.Min(t => t.Data)
                        };
                    flag = true;
                    foreach (var item in unstackTg) //檢查該 Target 內容是否合格
                    {
                        if (Convert.ToDouble(item.Value_Max) - Convert.ToDouble(item.Value_Min) > 1E-17)
                        {
                            invalidTargetCol.Add(this.ws.Columns.Item(tg).Name);
                            flag = false;
                            break;
                        }
                    }
                    if (flag) //如果該數據合格，則加入target 清單中，並整理至新工作表
                    {
                        validTargetCol.Add(this.ws.Columns.Item(tg).Label);
                        String nm = this.ws.Columns.Item(tg).Label;
                        unstackWs.Columns.Add(Quantity: 1).Name = nm;
                        foreach (var item in unstackTg) groupedData.Add(item.Value.Data);
                        unstackWs.Columns.Item(nm).SetData(groupedData.ToArray());
                    }
                }
            }

            //建立新的 Trend 物件            
            Trend trend = new Trend(this.proj, unstackWs);
            //複製工程
            trend.Line = this.Line.Clone();
            trend.LegendBox = this.LegendBox.Clone();
            trend.X_Scale = (CategoricalScale)this.X_Scale.Clone();
            trend.Y_Scale = (ContinuousScale)this.Y_Scale.Clone();
            trend.Annotation = this.Annotation.Clone();
            trend.Datalabel = this.Datalabel; //---> 還沒有寫 Clone 方法..需要嗎?
            trend.GraphSize = this.GraphSize;
            //開始設定轉換後參數
            String[] copyArray = new String[summaryColArray.Length - 1];
            Array.Copy(summaryColArray, 1, copyArray, 0, summaryColArray.Length - 1);
            Object obj = copyArray;
            trend.SetVariable(ref obj);
            obj = summaryColArray[0];
            trend.SetLabelVarible(ref obj); //Unstack 後的變數

            /*
             * 如果沒有指定則將 legend box 的section title 設為原本變數的 column label
             */
            if (this.LegendBox.SectTitle == null)
                trend.LegendBox.SectTitle = ws.Columns.Item(this.variables[0]).Label;

            //Graph framewoek 資訊
            trend.CopyGraphToClipboard(this.isCopyToClipboard);
            trend.SetExportCommand(this.isExportCmnd, this.pathOfExportCmnd);
            trend.SaveGraph(this.isSaveGraph, this.pathOfSaveGraph);



            if (invalidTargetCol != null)
            {
                trend.Annotation.AddFootnote("Invalid target variable: " + String.Join(", ", invalidTargetCol), 20, true);
            }
            if (validTargetCol != null)
            {
                obj = validTargetCol.ToArray();
                trend.SetTargetVariable(ref obj);
                trend.targetColor = this.targetColor;
                trend.targetType = this.targetType;
            }

            trend.Run();

        }

        private void CreateTSPlot()
        {
            /*
             * Run Minitab Command
             * 
             */

            StringBuilder mtbCmnd = new StringBuilder();
            StringBuilder exportString;

            mtbCmnd.AppendLine("NOTITLE" + Environment.NewLine + "BRIEF 0");
            String cmnd = GetCommand();

            if (this.GraphSize.Width != 576 || this.GraphSize.Height != 384)
            {
                mtbCmnd.Append(cmnd);
                mtbCmnd.AppendLine("GRAPH " + (double)this.GraphSize.Width / (96 * this.incrPercent / 100) + " " +
                    (double)this.GraphSize.Height / (96 * this.incrPercent / 100) + ".");
            }
            else
            {
                mtbCmnd.AppendLine(cmnd.Substring(0, cmnd.Length - Environment.NewLine.Length - 1) + ".");
            }
           
            Console.Write(mtbCmnd.ToString());
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
            if (this.isCopyToClipboard) CopyToClipboard("TSPLOT", proj, ws, cmndStart, proj.Commands.Count);
        }


        public void Run()
        {
            /*
             * 先檢查有無worksheet, proj 等物件
             */
            if (this.ws == null || this.proj == null)
            {
                throw new ArgumentNullException("Do not set Minitab Project or Worksheet.");
                return;
            }
            /*
             * 是否可以進行繪圖
             * 
             */
            if (this.variables == null && this.labvariable == null)
            {
                throw new ArgumentNullException("Variable or Label is null.");
                return;
            }
            if (this.variables.Count == 1)
            {
                if (groupvariable != null)
                {
                    if (this.groupvariable.Count == 1)
                    {
                        /*
                         * 使用 Stacked data trend 邏輯
                         */
                        this.StackedDataTrend();
                    }
                    else
                    {
                        return;//表示太多 group 欄位...目前只開放一組
                    }
                }
                else
                {
                    /*
                     * 使用 summary data trend 模式
                     */
                    this.CreateTSPlot();
                }
            }
            else
            {
                /*
                 * 使用 summary data trend 模式
                 */
                this.CreateTSPlot();
            }
        }
    }
}

