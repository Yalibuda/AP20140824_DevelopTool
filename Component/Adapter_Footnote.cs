using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component
{
    internal class Adapter_Footnote : IFootnote
    {
        List<Mtblib.Graph.Component.Footnote> _footnotes;
        public Adapter_Footnote(List<Mtblib.Graph.Component.Footnote> footnotes)
        {
            _footnotes = footnotes;

        }

        /// <summary>
        /// 加入註解，合法的輸入為單一(string) 或文字陣列(string[])
        /// </summary>
        /// <param name="footnote"></param>
        public void AddFootnote(dynamic footnote)
        {
            if (footnote == null) return;
            if (footnote is string)
            {
                Mtblib.Graph.Component.Footnote f = new Mtblib.Graph.Component.Footnote();
                f.Text = footnote;
                f.FontSize = this.FontSize;
                f.FontColor = this.FontColor;
                _footnotes.Add(f);
            }
            else
            {

                _footnotes.AddRange(((string[])footnote).Select(x => new Mtblib.Graph.Component.Footnote()
                {
                    Text = x,
                    FontSize = this.FontSize,
                    FontColor = this.FontColor
                }).ToList());

            }
        }

        /// <summary>
        /// 清除所有 Footnote 內容
        /// </summary>
        public void RemoveAll()
        {
            _footnotes.Clear();
        }

        public int FontColor
        {
            get
            {
                if (_footnotes != null && _footnotes.Count > 0)
                    return _footnotes[0].FontColor;
                else
                    return -1;
            }
            set
            {
                if (_footnotes != null && _footnotes.Count > 0)
                {
                    for (int i = 0; i < _footnotes.Count; i++)
                    {
                        _footnotes[i].FontColor = value;
                    }
                }
                else
                {
                    _footnotes.Add(
                        new Mtblib.Graph.Component.Footnote()
                        {
                            Text = string.Empty,
                            FontColor = value,
                            FontSize = this.FontSize
                        }
                        );
                }
            }
        }

        public float FontSize
        {
            get
            {
                if (_footnotes != null && _footnotes.Count > 0)
                    return _footnotes[0].FontSize;
                else
                    return -1;
            }
            set
            {
                if (_footnotes != null && _footnotes.Count > 0)
                {
                    for (int i = 0; i < _footnotes.Count; i++)
                    {
                        _footnotes[i].FontSize = value;
                    }
                }
                else
                {
                    _footnotes.Add(
                        new Mtblib.Graph.Component.Footnote()
                        {
                            Text = string.Empty,
                            FontColor = this.FontColor,
                            FontSize = value
                        }
                        );
                }
            }
        }
    }
}
