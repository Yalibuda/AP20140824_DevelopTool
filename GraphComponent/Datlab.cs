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
                if (this.datlabModelAttribute != null)
                {
                    foreach (DatlabModelAttribute model in datlabModelAttribute)
                    {
                        if (model.Start != null & model.End != null)
                        {
                            cmnd.AppendLine("  POSI " + model.Start + ":" + model.End + ";");
                            cmnd.AppendLine("  MODEL " + model.ModelIndex + ";");
                            if (model.Color != null) cmnd.AppendLine("   TCOLOR " + model.Color + ";");
                            if (model.Size != null) cmnd.AppendLine("   TSIZE " + model.Size + ";");
                            if (model.Offset != 0) cmnd.AppendLine("   OFFSET 0 " + model.Offset + ";");
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
