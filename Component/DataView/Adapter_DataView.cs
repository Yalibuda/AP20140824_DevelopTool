using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.DataView
{
    internal class Adapter_DataView : IDataView
    {
        Mtblib.Graph.Component.IDataView _dataview;
        
        public Adapter_DataView(Mtblib.Graph.Component.IDataView dataview)
        {
            _dataview = dataview;
        }
       
        public bool Visible
        {
            get
            {
                return _dataview.Visible;
            }
            set
            {
                _dataview.Visible = value;
            }
        }

        public void SetSize(dynamic var)
        {
            _dataview.Size = var;
        }

        public void SetType(dynamic var)
        {
            _dataview.Type = var;
        }

        public void SetColor(dynamic var)
        {
            _dataview.Color = var;
        }
    }
}
