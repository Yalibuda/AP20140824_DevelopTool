﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    public class DatlabModelAttribute
    {
        public int? ModelIndex { set; get; }
        public dynamic Label { set; get; }
        public int? Color { set; get; }
        public int? Size { set; get; }
        public dynamic Offset { set; get; }        
        public int? Start { set; get; }
        public int? End { set; get; }
    }
}
