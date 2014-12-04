using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Mtb;
using System.Text.RegularExpressions;

namespace MtbGraph
{
    internal class DialogAppraiser
    {
        internal DialogAppraiser() { }
        internal String GetColumnName(String str, Mtb.Worksheet ws)
        {
            String colName = null;
            try
            {
                if (str.Substring(0, 1) == "'" & str.Substring(str.Length - 1, 1) == "'")
                {
                    colName = str.Substring(1, str.Length - 2);
                }
                else
                {
                    colName = str;
                }
                ws.Columns.Item(colName);
            }
            catch (Exception e)
            {
                MessageBox.Show("#" + e.HResult + " - " + e.Message, mTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return colName;
        }
        internal String GetColumnName(String str, Mtb.Worksheet ws, Boolean getID)
        {
            String colName = null;
            if (String.IsNullOrEmpty(str))
            {
                return colName;
            }
            try
            {
                if (str.Substring(0, Math.Min(1, str.Length)) == "'" & str.Substring(Math.Max(0, str.Length - 1), Math.Min(1, str.Length)) == "'")
                {
                    colName = str.Substring(1, str.Length - 2);
                }
                else
                {
                    colName = str;
                }
                if (getID)
                {
                    colName = ws.Columns.Item(colName).SynthesizedName;
                }
                else
                {
                    colName = ws.Columns.Item(colName).Label;
                }

            }
            catch (Exception e)
            {
                MessageBox.Show("#" + e.HResult + " - " + e.Message, mTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return colName;
        }
        internal int GetColumnIndex(String str, Mtb.Worksheet ws)
        {
            String colID;
            int index = 0;
            try
            {
                colID = GetColumnName(str, ws, true);
                for (int i = 0; i < ws.Columns.Count; i++)
                {
                    if (ws.Columns.Item(i + 1).SynthesizedName == colID)
                    {
                        index = i + 1;
                        break;
                    }
                }
            }
            catch { }
            return index;
        }
        internal Boolean IsColumnExist(String col)
        {
            return false;
        }
        internal Boolean IsWorksheetExist(String ws)
        {
            return false;
        }

        internal List<String> GetMtbCols(List<String> l, Mtb.Worksheet ws)
        {
            Regex regEx;
            regEx = new Regex(@"([^'#\s-]+|'[^'#]+')-([^'#\s-]+|'[^'#]+')");
            List<String> colInfo = new List<String>();
            String mItem1;
            String mItem2;
            int k1;
            int k2;
            try
            {
                foreach (String s in l)
                {
                    if (regEx.IsMatch(s))
                    {
                        Match m = regEx.Match(s);
                        mItem1 = GetColumnName(m.Groups[1].Value, ws, true);
                        mItem2 = GetColumnName(m.Groups[2].Value, ws, true);
                        k1 = GetColumnIndex(mItem1, ws);
                        k2 = GetColumnIndex(mItem2, ws);
                        if (k1 <= k2)
                        {
                            for (int i = k1; i <= k2; i++)
                            {
                                colInfo.Add(ws.Columns.Item(i).SynthesizedName);
                            }
                        }
                        else
                        {
                            for (int i = k1; i >= k2; i--)
                            {
                                colInfo.Add(ws.Columns.Item(i).SynthesizedName);
                            }
                        }
                    }
                    else
                    {
                        colInfo.Add(GetColumnName(s, ws, true));
                    }
                }
            }
            catch(Exception e)
            {
                throw e;
            }
            return colInfo;
        }
        internal List<String> GetMtbColInfo(String inputStr)
        {
            Regex regEx;

            Boolean b;
            //Check validation of String
            regEx = new Regex(@"''|#|\*");
            b = regEx.IsMatch(inputStr);
            if (b)
            {
                MessageBox.Show("Invalid character",mTitle,MessageBoxButtons.OK,MessageBoxIcon.Warning );
                return null;
            }

            //Get each element by patterm [^'#\s-]+|'[^'#]+'
            regEx = new Regex(@"(([^'#\s-]+|'[^'#]+')-([^'#\s-]+|'[^'#]+'))|[^'#\s-]+|'[^'#]+'");
            List<String> l = new List<String>();
            try
            {
                if (regEx.IsMatch(inputStr))
                {
                    foreach (Match m in regEx.Matches(inputStr))
                    {
                        l.Add(m.Value);
                    }
                }

            }
            catch(Exception e)
            {
                throw e;         
            }
            return l;
        }

        //Local variables
        private String mTitle = "Dialog Appraiser";

    }
}
