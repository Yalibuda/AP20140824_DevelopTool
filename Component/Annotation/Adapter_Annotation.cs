using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Annotation
{
    internal class Adapter_Annotation : IAnnotation
    {
        Mtblib.Graph.Component.Annotation.Annotation _annotation;
        public Adapter_Annotation(Mtblib.Graph.Component.Annotation.Annotation annotation)
        {
            _annotation = annotation;
        }

        public bool Visible
        {
            get { return _annotation.Visible; }
            set { _annotation.Visible = value; }
        }

        public void SetColor(dynamic var)
        {
            _annotation.Color = var;
        }

        public void SetSize(dynamic var)
        {
            _annotation.Size = var;
        }

        public void SetType(dynamic var)
        {
            _annotation.Type = var;
        }
    }
}
