using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class Datlab : ICOMInterop_Datlab, IDatLabel
    {

        public bool Show { get; set; }
        public DatlabType LabelType { get; set; }
        public DatlabColor Color { set; get; }
        public DatlabPlace Place { set; get; }

        /*
         * 設定要不顯示 Datlab 的 model, position (start:end)
         * 因為要給定 Label，所以每一個 Position 都需要單獨處理
         */
        List<DatlabModelAttribute> invisibleDatlabModel = null;
        public void SetDatlabInvisible(List<DatlabModelAttribute> modelAttribute)
        {
            invisibleDatlabModel = modelAttribute;
        }

        /*
         * 設定客製化 Datlab 的 model, position (start:end)
         */
        List<DatlabModelAttribute> datlabModelAttribute = null;
        public void SetCustomDatlab(List<DatlabModelAttribute> modelAttribute)
        {
            datlabModelAttribute = modelAttribute;
        }


        String labColumn = String.Empty;
        public void SetLabelFromColumn(string col = null)
        {

            if (String.IsNullOrEmpty(col))
            {
                labColumn = String.Empty;
            }
            else
            {
                labColumn = col;
                this.LabelType = DatlabType.LabFromColumn;
            }

        }

        public string GetCommand()
        {
            StringBuilder cmnd = new StringBuilder();
            if (Show)
            {
                switch (this.LabelType)
                {
                    case DatlabType.LabFromColumn:
                        if (labColumn == String.Empty)
                        {
                            return @"#未指定 Datlabel 欄位";
                        }
                        else
                        {
                            cmnd.AppendLine(" DATLAB " + labColumn + ";");
                        }
                        break;
                    case DatlabType.Value:
                    case DatlabType.RowNum:
                        cmnd.AppendLine(" DATLAB;");
                        break;
                }
                if (this.Place == DatlabPlace.Center)
                {
                    cmnd.AppendLine("  PLACE 0 0;");
                }
                else if (this.Place == DatlabPlace.Below)
                {
                    cmnd.AppendLine("  PLACE 0 -1;");
                }

                if (this.datlabModelAttribute != null)
                {
                    foreach (DatlabModelAttribute model in datlabModelAttribute)
                    {
                        if (model.Start != null & model.End != null & model.ModelIndex != null)
                        {
                            if (model.Label != null)//表示有指定 label
                            {
                                /*
                                 * 每一個 model 的 label 用一個陣列表示，因為可能是從
                                 * Minitab column 中取出，所以屬性是 dynamic
                                 */
                                Type t = model.Label.GetType();
                                if (t.IsArray)
                                {
                                    System.Collections.IList ilist = model.Label as System.Collections.IList;
                                    List<String> labels = new List<String>();
                                    foreach (object o in ilist) labels.Add(o.ToString());//將 datlab 資料納入

                                    /*
                                     * 建一組長度和 labels 一樣的List 來放 Offset資料，因為 Offset 可能是陣列(在
                                     * Stack bar custom data label)
                                     */
                                    List<double> offsets = new List<double>();
                                    if (model.Offset.GetType().IsArray)
                                    {
                                        ilist = model.Offset as System.Collections.IList;
                                        foreach (object o in ilist) offsets.Add(Convert.ToDouble(o));
                                    }
                                    else
                                    {
                                        for (int i = 0; i < labels.Count; i++) offsets.Add((double)model.Offset);
                                    }

                                    for (int i = (int)model.Start; i <= (int)model.End; i++)//Minitab 是以1為base開始算...
                                    {
                                        cmnd.AppendLine("  POSI " + i + " \"" + labels[(i-1) % labels.Count] + "\";");
                                        cmnd.AppendLine("   MODEL " + model.ModelIndex + ";");
                                        if (model.Color != null) cmnd.AppendLine("   TCOLOR " + model.Color + ";");
                                        if (model.Size != null) cmnd.AppendLine("   TSIZE " + model.Size + ";");
                                        if (model.Offset != null) cmnd.AppendLine("   OFFSET 0 " + offsets[(i - 1) % offsets.Count] + ";");
                                    }
                                }
                            }
                            else
                            {
                                cmnd.AppendLine("  POSI " + model.Start + ":" + model.End + ";");
                                cmnd.AppendLine("   MODEL " + model.ModelIndex + ";");
                                if (model.Color != null) cmnd.AppendLine("   TCOLOR " + model.Color + ";");
                                if (model.Size != null) cmnd.AppendLine("   TSIZE " + model.Size + ";");
                                if (model.Offset != null) cmnd.AppendLine("   OFFSET 0 " + model.Offset + ";");
                            }
                            cmnd.AppendLine("  ENDP;");
                        }
                    }
                }

                if (this.invisibleDatlabModel != null)//有指定隱藏, 通常是Target 
                {
                    foreach (DatlabModelAttribute model in invisibleDatlabModel)
                    {
                        if (model.Start != null & model.End != null)
                        {
                            for (int i = (int)model.Start; i <= (int)model.End; i++)
                            {
                                cmnd.AppendLine("  POSI " + i + " \"\";");
                                cmnd.AppendLine("   MODEL " + model.ModelIndex + ";");
                            }
                            cmnd.AppendLine("  ENDP;");
                        }
                    }
                }
            }
            return cmnd.ToString();
        }


        public void SetDefault()
        {
            this.LabelType = DatlabType.Value;
            this.Show = false;
            this.labColumn = String.Empty;
            this.invisibleDatlabModel = new List<DatlabModelAttribute>();
            this.datlabModelAttribute = new List<DatlabModelAttribute>();
        }

    }
}
