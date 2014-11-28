using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Mtb;
using System.Collections;

namespace MtbGraph
{
    enum MtbVarType
    {
        Column,
        Constant,
        Matrix
    }
    internal class MtbTools
    {
        internal String[] CreateVariableStrArray(Mtb.Worksheet ws, int num, MtbVarType mType)
        {
            int cnt = 0;
            String[] varStr = new String[num]; //num have to large than 1
            try
            {
                switch (mType)
                {
                    case MtbVarType.Column:
                        cnt = ws.Columns.Count;
                        ws.Columns.Add(Quantity: num);
                        for (int i = 0; i < varStr.Length; i++)
                        {
                            varStr[i] = ws.Columns.Item(cnt + 1 + i).SynthesizedName;
                        }
                        break;
                    case MtbVarType.Constant:
                        cnt = ws.Constants.Count;
                        ws.Constants.Add(Quantity: num);
                        for (int i = 0; i < varStr.Length; i++)
                        {
                            varStr[i] = ws.Constants.Item(cnt + 1 + i).SynthesizedName;
                        }
                        break;
                    case MtbVarType.Matrix:
                        cnt = ws.Matrices.Count;
                        ws.Matrices.Add(Quantity: num);
                        for (int i = 0; i < varStr.Length; i++)
                        {
                            varStr[i] = ws.Matrices.Item(cnt + 1 + i).SynthesizedName;
                        }
                        break;
                    default:
                        break;
                }

                return varStr;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Source + "-" + ex.Message);
                return null;

            }

        }

        public List<String> TransObjToMtbColList(Object varCols, Mtb.Worksheet ws)
        {
            if (varCols == null || ws == null) return null;

            Type t = varCols.GetType();
            List<String> cols = new List<String>();
            DialogAppraiser da = new DialogAppraiser();
            if (t.IsArray)
            {
                try
                {
                    IEnumerable enumerable = varCols as IEnumerable;
                    foreach (object o in enumerable)
                    {
                        cols.Add(o.ToString());
                    }
                    cols = da.GetMtbCols(cols, ws);

                }
                catch
                {
                    throw new ArgumentException("Invalid input of scale variables");
                    return null;
                }

            }
            else if (Type.GetTypeCode(t) == TypeCode.String)
            {
                cols = da.GetMtbColInfo(varCols.ToString());
                cols = da.GetMtbCols(cols, ws);
            }
            else
            {
                cols = null;
            }
            return cols;
        }

    }
}
