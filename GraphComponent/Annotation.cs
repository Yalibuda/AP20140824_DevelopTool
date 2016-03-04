using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class Annotation : ICOMInterop_Annotation
    {
        public Annotation()
        {
            SetDefault();
        }

        public Annotation Clone()
        {
            Annotation annotation = new Annotation();
            annotation.Title = this.Title;
            annotation.TitleFontSize = this.TitleFontSize;
            if (_footnote != null)
            {
                foreach (INotation item in _footnote)
                {
                    annotation.AddFootnote(item.Clone());
                }
            }
          
            return annotation;
        }


        /**
         * 
         * 新版 Annotation 做法: 用 Title、Footnote 物件去控制
         * 
         */
        private INotation _title;
        private INotation[] _footnote;


        /// <summary>
        /// 設定或取得 Title 的文字
        /// </summary>
        public String Title
        {
            set
            {
                _title.Text = value;
            }
            get
            {
                return _title.Text;
            }
        }

        /// <summary>
        /// 設定或取得 Title 大小
        /// </summary>
        public float TitleFontSize
        {
            set
            {
                _title.Size = value;
            }
            get
            {
                return _title.Size;
            }
        }

        //private List<String> footnotes;
        public void AddFootnote(String footnote)
        {
            Footnote f = new Footnote();
            f.Text = footnote;
            AddFootnote(f);
            //AddFootnote(footnote, -1, false);

        }

        public void AddFootnote(INotation footnote)
        {
            List<INotation> l = new List<INotation>();

            if (_footnote != null)
                l = _footnote.ToList();
            INotation f = footnote.Clone();
            l.Add(f);
            _footnote = l.ToArray();
        }

        //private List<int> footnoteColor;
        //private List<bool> footnoteItalic;
        //public void AddFootnote(String footnote, int color = -1, bool italic = false)
        //{
        //    List<INotation> fList = new List<INotation>();
        //    if (_footnote != null) fList = _footnote.ToList();

        //    Footnote f = new Footnote();
        //    f.Text = footnote;
        //    f.Color = color;
        //    f.Italic = italic;

        //    footnotes.Add(footnote);
        //    footnoteColor.Add(color);
        //    footnoteItalic.Add(italic);
        //}

        public void RemoveFootnoteAt(int i)
        {
            List<INotation> l = _footnote.ToList();
            l.RemoveAt(i);
            _footnote = l.ToArray();
        }

        public void ClearFootnote()
        {
            _footnote = null;         
        }

        public void SetDefault()
        {
            _title = new Title();
            _footnote = null;
        }

        public String GetCommand()
        {
            StringBuilder cmnd = new StringBuilder();
            if (this.Title == String.Empty)
            {
                cmnd.AppendLine(" NODT;");
            }
            else if (this.Title != null)
            {
                cmnd.AppendLine(" TITLE \"" + this.Title + "\";");
                if (this.TitleFontSize != 13) cmnd.AppendLine("  PSIZE " + this.TitleFontSize + ";");
            }

            if (_footnote != null && _footnote.Length > 0)
            {
                foreach (INotation item in _footnote)
                {
                    cmnd.AppendLine(string.Format(" FOOT \"{0}\";", item.Text));
                    if (item.Color != -1) cmnd.AppendLine(string.Format("  TCOLOR {0};", item.Color));
                    if (item.Size != 9) cmnd.AppendLine(string.Format("  PSIZE {0};", item.Size));
                    if (item.Italic) cmnd.AppendLine("  ITALIC;");
                    if (item.Bold) cmnd.AppendLine("  BOLD;");
                }
            }
            

            return cmnd.ToString();
        }
    }
}
