﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.DataView
{
    public interface IDataView
    {
        bool Visible { set; get; }
        void SetType(dynamic var);
        void SetColor(dynamic var);
        void SetSize(dynamic var);
    }
}
