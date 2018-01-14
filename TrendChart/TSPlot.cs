using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.TrendChart
{
    public class TSPlot : ITSPlot, IDisposable
    {

        protected Mtblib.Graph.TimeSeriesPlot.TSPlot _tsplot;
        protected Mtb.Project _proj;
        protected Mtb.Worksheet _ws;
        public TSPlot()
        {

        }
        public TSPlot(Mtb.Project proj, Mtb.Worksheet ws)
        {
            SetMtbEnvironment(proj, ws);
        }

        public void SetVariables(dynamic var)
        {
            Variables = var;
        }
        protected dynamic Variables
        {
            set { _tsplot.Variables = value; }
            get { return _tsplot.Variables; }
        }

        public virtual void SetGroupingBy(dynamic var)
        {
            GroupingBy = var;
        }
        protected dynamic GroupingBy
        {
            get
            {
                return _tsplot.GroupingVariables;
            }
            set
            {
                _tsplot.GroupingVariables = value;
            }
        }

        public void SetStamp(dynamic var)
        {
            Stamp = var;
        }
        protected dynamic Stamp
        {
            get
            {
                return _tsplot.Stamp;
            }
            set
            {
                _tsplot.Stamp = value;
            }
        }

        public Component.Scale.IContScale XScale
        {
            get { return new Component.Scale.Adapter_ContScale(_tsplot.XScale); }

        }

        public Component.Scale.IContScale YScale
        {
            get { return new Component.Scale.Adapter_ContScale(_tsplot.YScale); }
        }

        public Component.IDatlab Datlab
        {
            get { return new Component.Adapter_DatLab(_tsplot.DataLabel); }
        }

        public Component.Region.ILegend Legend
        {
            get { return new Component.Region.Adapter_Legend(_tsplot.Legend); }
        }
        public Component.Region.IRegion DataRegion
        {
            get { return new Component.Region.Adapter_Region(_tsplot.DataRegion); }
        }
        public Component.Region.IGraph Graph
        {
            get { return new Component.Region.Adapter_Graph(_tsplot.GraphRegion); }
        }

        public Component.ILabel Title
        {
            get { return new Component.Adapter_Lab(_tsplot.Title); }
        }

        public Component.IFootnote Footnotes
        {
            get { return new Component.Adapter_Footnote(_tsplot.FootnoteLst); }
        }

        public string GSave
        {
            get
            {
                return _tsplot.GraphPath;
            }
            set
            {
                _tsplot.GraphPath = value;
            }
        }

        public void SetMtbEnvironment(Mtb.Project proj, Mtb.Worksheet ws)
        {
            _proj = proj;
            _ws = ws;
            SetDefault();
        }

        protected virtual void SetDefault()
        {
            _tsplot = new Mtblib.Graph.TimeSeriesPlot.TSPlot(_proj, _ws);

        }
        protected virtual string GetCommand()
        {
            return null;
        }

        public virtual void Run()
        {
            StringBuilder cmnd = new StringBuilder();
            string macroPath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("mytsplot.mac", GetCommand());
            cmnd.AppendLine("notitle");
            cmnd.AppendLine("brief 0");
            cmnd.AppendFormat("%\"{0}\" {1};\r\n", macroPath,
                string.Join(" &\r\n", ((Mtb.Column[])Variables).Select(x => x.SynthesizedName).ToArray()));
            if (GroupingBy != null)
            {
                cmnd.AppendFormat("group {0};\r\n", string.Join(" &\r\n",
                    ((Mtb.Column[])GroupingBy).Select(x => x.SynthesizedName).ToArray()));
            }

            cmnd.AppendLine(".");
            cmnd.AppendLine("title");
            cmnd.AppendLine("brief 2");
            string fpath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("mycode.mtb", cmnd.ToString());

            _proj.ExecuteCommand(string.Format("exec \"{0}\" 1", fpath));
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free other state (managed objects).
                _tsplot.Dispose();
            }
            // Free your own state (unmanaged objects).
            // Set large fields to null.
            Variables = null;
            GroupingBy = null;
            _proj = null;
            _ws = null;
            GC.Collect();

        }
        ~TSPlot()
        {
            Dispose(false);
        }
    }
}
