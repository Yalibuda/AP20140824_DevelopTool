using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    public interface ICOMInterop_GFrameWork
    {
        void SaveGraph(bool b, String outputPath);
        void SetExportCommand(bool b, String outputPath = null);
        void CopyGraphToClipboard(bool b);
        void Dispose();
    }
}
