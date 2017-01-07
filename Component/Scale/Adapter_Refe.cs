using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Scale
{
    internal class Adapter_Refe : IRefe
    {
        Mtblib.Graph.Component.Scale.IRefe _refe;
        public Adapter_Refe(Mtblib.Graph.Component.Scale.IRefe refe)
        {
            _refe = refe;
        }

        public void AddValue(double value)
        {
            if (_refe.Values == null)
            {
                _refe.Values = value;
                return;
            }
            List<string> valueStr = ((string[])_refe.Values).ToList();
            valueStr.Add(value.ToString());
            _refe.Values = valueStr.ToArray();
        }

        public void SetLabel(dynamic var)
        {
            _refe.Labels = var;

        }

        public void RemoveAll()
        {
            _refe.Values = null;
        }

        public float FontSize
        {
            get
            {
                return _refe.FontSize;
            }
            set
            {
                _refe.FontSize = value;
            }
        }

        public void SetSize(dynamic var)
        {
            _refe.Size = var;

        }

        public void SetType(dynamic var)
        {
            _refe.Type = var;
        }

        public void SetColor(dynamic var)
        {
            _refe.Color = var;
        }

    }
}
