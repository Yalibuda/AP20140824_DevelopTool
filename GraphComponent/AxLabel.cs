using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class AxLabel : ICOMInterop_AxLab
    {
        private ScaleType scale_axis;
        public AxLabel(ScaleType scale_axis)
        {
            this.scale_axis = scale_axis;
            this.Label = null;
            this.FontSize = 11;
        }
        public String Label { set; get; }
        public double FontSize { set; get; }

        public void SetDefault()
        {
            this.Label = null;
            this.FontSize = 11;
        }

        public AxLabel Clone()
        {
            AxLabel axla = new AxLabel(this.scale_axis);
            axla.Label = this.Label;
            axla.FontSize = this.FontSize;
            return axla;
        }

        public String GetCommand()
        {
            StringBuilder cmnd = new StringBuilder();
            int k = 0;
            switch (scale_axis)
            {
                case ScaleType.X_axis:               
                    k = 1;
                    break;
                case ScaleType.Y_axis:
                case ScaleType.Secondary_Y_axis:
                    k = 2;
                    break;
            }
            if (this.Label != null)
            {
                cmnd.AppendLine(" AxLa " + k + " \"" + this.Label + "\";");
                if (scale_axis == ScaleType.Secondary_Y_axis)
                {
                    cmnd.AppendLine("  SECS;");
                }
                if (this.FontSize != 11) cmnd.AppendLine("  PSIZE " + this.FontSize + ";");
            }
            
            return cmnd.ToString();
            
        }
    }
}
