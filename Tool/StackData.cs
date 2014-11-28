using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Tool
{
    class StackData
    {
        public StackData()
        {
            this.Data = new List<GroupedData>();
        }
        public StackData(dynamic datas, dynamic subscript)
        {
            this.Data = new List<GroupedData>();
            SetData(datas, subscript);
        }
        public List<GroupedData> Data { set; get; }
        public void SetData(dynamic datas, dynamic subscript)
        {
            /*
             * 只允許datas 和 subscript 以單值輸入(如 datas = 1, subscript = "A")或均為陣列，就算是未定資料數的狀況，
             * 理當也是用陣列去接數據。所以當 datas = int[1]{10} 和 subscript = 1 的狀況不被允許。
             * 
             */ 

            Type t1 = datas.GetType();
            Type t2 = subscript.GetType();

            if (t1.IsArray & t2.IsArray)
            {
                IList ilist1 = datas as IList;
                IList ilist2 = subscript as IList;
                if (ilist1.Count != ilist2.Count)
                {
                    return;
                }
                else
                {
                    GroupedData singledata;
                    List<GroupedData> _data = new List<GroupedData>();
                    try
                    {
                        for (int i = 0; i < ilist1.Count; i++)
                        {
                            singledata = new GroupedData();
                            singledata.Data = ilist1[i];
                            singledata.Subscript = ilist2[i];
                            _data.Add(singledata);
                        }
                    }
                    catch
                    {
                        Console.WriteLine("Opps, something wrong 1.");
                        return;
                    }

                    Data = _data;                    
                }

            }
            else
            {
                GroupedData d = new GroupedData();
                switch (Type.GetTypeCode(t1))
                {
                    case TypeCode.Int16:
                    case TypeCode.Double:
                    case TypeCode.Decimal:
                    case TypeCode.Single:
                    case TypeCode.Int32:
                    case TypeCode.DateTime:
                    case TypeCode.String:
                        d.Data = datas;
                        break;
                    default:
                        Console.WriteLine("Invalid type of data");
                        return;
                        break;
                }
                switch (Type.GetTypeCode(t2))
                {
                    case TypeCode.Int16:
                    case TypeCode.Double:
                    case TypeCode.Decimal:
                    case TypeCode.Single:
                    case TypeCode.Int32:
                    case TypeCode.DateTime:
                    case TypeCode.String:
                        d.Subscript = subscript;
                        break;
                    default:
                        Console.WriteLine("Invalid type of data");
                        return;
                        break;
                }

                this.Data = new List<GroupedData>();
                this.Data.Add(d);
            }
        }
    }
}
