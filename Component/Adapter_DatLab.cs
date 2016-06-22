using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component
{
    internal class Adapter_DatLab : IDatlab
    {
        Mtblib.Graph.Component.Datlab _datlab;
        public Adapter_DatLab(Mtblib.Graph.Component.Datlab datlab)
        {
            _datlab = datlab;            
        }

        public bool Visible
        {
            get { return _datlab.Visible; }
            set { _datlab.Visible = value; }
        }
        
        public float FontSize
        {
            get
            {
                return _datlab.FontSize;
            }
            set
            {
                _datlab.FontSize = value;
            }
        }
    }
}
