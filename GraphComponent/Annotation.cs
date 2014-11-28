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
            Title = null;
            footnotes = new List<string>();
            footnoteColor = new List<int>();
            footnoteItalic = new List<bool>();
            this.TitleFontSize = 13;
            this.FootnoteFontSize = 9;
        }

        public Annotation Clone()
        {
            Annotation annotation = new Annotation();
            annotation.Title = this.Title;
            annotation.TitleFontSize = this.TitleFontSize;
            annotation.FootnoteFontSize = this.FootnoteFontSize;
            if (this.footnotes != null)
            {
                for (int i = 0; i < this.footnotes.Count; i++)
                    annotation.AddFootnote(this.footnotes[i], this.footnoteColor[i], this.footnoteItalic[i]);
            }

            return annotation;
        }

        public String Title { set; get; }
        public double TitleFontSize { set; get; }

        public double FootnoteFontSize { set; get; }

        private List<String> footnotes;
        public void AddFootnote(String footnote)
        {
            AddFootnote(footnote, -1, false);
        }
        private List<int> footnoteColor;
        private List<bool> footnoteItalic;
        public void AddFootnote(String footnote, int color = -1, bool italic = false)
        {
            footnotes.Add(footnote);
            footnoteColor.Add(color);
            footnoteItalic.Add(italic);
        }

        public void RemoveFootnoteAt(int i)
        {
            footnotes.RemoveAt(i);
            footnoteColor.RemoveAt(i);
            footnoteItalic.RemoveAt(i);
        }

        public void ClearFootnote()
        {
            footnotes.Clear();
            footnoteColor.Clear();
            footnoteItalic.Clear();
        }

        public void SetDefault()
        {
            Title = null;
            ClearFootnote();
            this.TitleFontSize = 13;
            this.FootnoteFontSize = 9;
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

            if (this.footnotes.Count > 0)
            {
                for (int i = 0; i < footnotes.Count; i++)
                {
                    cmnd.AppendLine(" FOOT \"" + footnotes[i] + "\";");
                    if (footnoteColor[i] != -1) cmnd.AppendLine("  TCOLOR " + footnoteColor[i] + ";");
                    if (footnoteItalic[i]) cmnd.AppendLine("  ITALIC;");
                    if (FootnoteFontSize != 9) cmnd.AppendLine("  PSIZE " + this.FootnoteFontSize + ";");

                }
            }

            return cmnd.ToString();
        }
    }
}
