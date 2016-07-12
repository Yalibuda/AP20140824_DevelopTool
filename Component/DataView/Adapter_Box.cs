using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.DataView
{
    internal class Adapter_Box : IBox
    {
        Mtblib.Graph.Component.IDataView _dataview;
        Adapter_Box(Mtblib.Graph.Component.IDataView dataview)
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

        public void SetEType(dynamic var)
        {
            _dataview.EType = var;
        }

        public void SetEColor(dynamic var)
        {
            _dataview.EColor = var;
        }

        public void SetESize(dynamic var)
        {
            _dataview.ESize = var;
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
