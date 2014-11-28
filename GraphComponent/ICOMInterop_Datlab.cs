using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    public interface ICOMInterop_Datlab
    {
        bool Show { set; get; }
        DatlabColor Color { set; get; }
        //DatlabType LabelType { set; get; }
        //void SetDatlabInvisible(List<DatlabModelAttribute> modelAttribute); //VB6 無法顯示此方法
        //void SetCustomDatlab(List<DatlabModelAttribute> modelAttribute); //VB6 無法顯示此方法
        void SetDefault();
    }
}
