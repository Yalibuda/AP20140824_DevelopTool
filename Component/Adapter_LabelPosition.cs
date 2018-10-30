using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component
{
    public class Adapter_LabelPosition : ILabelPosition
    {
        Mtblib.Graph.Component.LabelPosition _labelposition;
        public Adapter_LabelPosition(Mtblib.Graph.Component.LabelPosition labelposition)
        {
            _labelposition = labelposition;
        }
        public bool Visible
        {
            get
            {
                return _labelposition.Visible;
            }
            set
            {
                _labelposition.Visible = value;
            }
        }
        public float FontSize
        {
            get
            {
                return _labelposition.FontSize;
            }
            set
            {
                _labelposition.FontSize = value;
            }
        }
        public int FontColor
        {
            get
            {
                return _labelposition.FontColor;
            }
            set
            {
                _labelposition.FontColor = value;
            }
        }
        public int Model
        {
            get { return _labelposition.Model; }
            set { _labelposition.Model = value; }
        } 
        public int[] RowId
        {
            get
            {
                return _labelposition.RowId;
            }
            set
            {
                if (value != null & value is Array)
                {
                    if (value.Length > 1) throw new ArgumentException("TextPosition 只能輸入一個 RowId");
                }
                _labelposition.RowId = value;
            }
        }
        public string Text
        {
            get
            {
                return _labelposition.Text;
            }
            set
            {
                _labelposition.Text = value;
            }
        }

    }
}
