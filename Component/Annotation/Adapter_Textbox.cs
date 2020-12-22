using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Annotation
{
    internal class Adapter_Textbox : ITextBox
    {
        Mtblib.Graph.Component.Annotation.Textbox _textbox;
        public Adapter_Textbox(Mtblib.Graph.Component.Annotation.Textbox textbox)
        {
            _textbox = textbox;
        }
        public string Text
        {
            get { return _textbox.Text; }
            set { _textbox.Text = value; }
        }
        public int Unit
        {
            get { return _textbox.Unit; }
            set { _textbox.Unit = value; }
        }
       
        //public void SetTextSize(dynamic var)
        //{
        //    _textbox.Size = var;
        //}
        public void SetCoordinate(params object[] args)
        {
            _textbox.SetCoordinate(args);
        }
        public string[] GetCoordinate()
        {
            return _textbox.GetCoordinate();
        }

        protected string[] _boxposition = null;
        public string[] Boxposition
        {
            get { return _boxposition; }
            set { _boxposition = value; }
        }
        public void SetBoxposition(params object[] args)
        {
            
            if (args.Length != 4) throw new ArgumentException("有不正確的參數個數，必須為 4 個!");

            string[] boxposizion;
            boxposizion = args.Where(x => x != null).Select(x => x.ToString()).
                Where(x => !string.IsNullOrEmpty(x) && !string.IsNullOrWhiteSpace(x)).ToArray();

            if (boxposizion.Length != args.Length)
            {
                throw new ArgumentException("座標不可包含 null 或空白。");
            }
            Boxposition = boxposizion;
        }
        public string GetCommand()
        {
            if (_textbox.GetCoordinate() == null || string.IsNullOrEmpty(_textbox.Text)) return "";
            StringBuilder cmnd = new StringBuilder();
            cmnd.AppendLine("Text &");
            cmnd.AppendLine(string.Join(" &\r\n", _textbox.GetCoordinate()) + " &");
            cmnd.AppendLine(string.Format("{0};", _textbox.Text));
            // textbox size
            cmnd.AppendLine("Box &");
            foreach (string str in Boxposition) cmnd.AppendFormat("   {0} & \r\n", str);
            cmnd.AppendLine("   ;");
            cmnd.AppendLine(string.Format("Unit {0};", Unit));
            //cmnd.AppendLine(string.Format("PSize {0};", _textbox.Size));
            return cmnd.ToString();
            #region Closed
            //if (FontColor > 0) cmnd.AppendLine(string.Format("TColor {0};", FontColor));
            //if (FontSize > 0) cmnd.AppendLine(string.Format("PSize {0};", FontSize));

            //if (Angle < MtbTools.MISSINGVALUE) cmnd.AppendLine(string.Format("Angle {0};", Angle));
            //if (Bold) cmnd.AppendLine("Bold;");
            //if (Italic) cmnd.AppendLine("Italic;");
            //if (Underline) cmnd.AppendLine("Under;");

            //if (Offset != null)
            //{
            //    cmnd.AppendLine(string.Format("Offset {0} {1};", Offset[0], Offset[1]));
            //}

            //if (Placement != null)
            //{
            //    cmnd.AppendLine(string.Format("Placement {0} {1};", Placement[0], Placement[1]));
            //}
            #endregion
        }
    }
}
