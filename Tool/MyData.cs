using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mtb;
using System.Collections;
namespace MtbGraph.Tool
{
    public enum DataType
    {
        Numeric, Text
    }

    class MyData
    {
        public String id { set; get; }
        public Object data { private set; get; }
        public DataType dtype { private set; get; }
        public Object cate { private set; get; }
        public DataType catetype { private set; get; }
        public MyData()
        {
            id = null;
            data = null;
            cate = null;
            dtype = DataType.Numeric;
            catetype = DataType.Text;
        }

        public MyData(Object data, Object group)
        {
            //Set data
            SetData(data);
            //Set group
            SetCategory(data);
        }

        public MyData(String mtbColID, Object data, Object group)
        {
            this.id = mtbColID;
            //Set data
            SetData(data);
            //Set group
            SetCategory(data);
        }

        public void SetData(Object data)
        {
            this.data = data;
            TypeCode t = Type.GetTypeCode(data.GetType());
            switch (t)
            {
                case TypeCode.String:
                    dtype = DataType.Text;
                    break;
                case TypeCode.Int16:
                case TypeCode.Double:
                case TypeCode.Decimal:
                case TypeCode.Single:
                    dtype = DataType.Numeric;
                    break;

            }
        }

        public void SetCategory(Object group)
        {
            this.cate = group;
            TypeCode t = Type.GetTypeCode(group.GetType());
            switch (t)
            {
                case TypeCode.String:
                    catetype = DataType.Text;
                    break;
                case TypeCode.Int16:
                case TypeCode.Double:
                case TypeCode.Decimal:
                case TypeCode.Single:
                    catetype = DataType.Numeric;
                    break;
            }
        }

    }

    class MyDatas
    {
        public List<MyData> datas { private set; get; }
        public MyDatas()
        {
            datas = new List<MyData>();
        }
        public void SetDatas(List<Object> data, List<Object> gp)
        {
            if (data.Count != gp.Count)
            {
                throw new ArgumentException("Length of data and gp are not equal");
                return;
            }
            MyData m;
            for (int i = 0; i < data.Count; i++)
            {
                m = new MyData();
                m.SetCategory(gp[i]);
                m.SetData(data[i]);
                datas.Add(m);
            }
        }
        public void SetDatas(List<String> id, List<Object> data, List<Object> gp)
        {
            if (data.Count != gp.Count || data.Count != id.Count || id.Count != gp.Count)
            {
                throw new ArgumentException("Length of data, group and id are not equal!");
                return;
            }
            MyData m;
            for (int i = 0; i < data.Count; i++)
            {
                m = new MyData();
                m.id = id[i];
                m.SetCategory(gp[i]);
                m.SetData(data[i]);
                datas.Add(m);
            }
        }
        public MyDatas(List<Object> data, List<Object> gp)
        {
            datas = new List<MyData>();
            SetDatas(data, gp);
        }
        public MyDatas(List<String> id, List<Object> data, List<Object> gp)
        {
            datas = new List<MyData>();
            SetDatas(id, data, gp);
        }
    }

    class MyDataCompareByData : IComparer<MyData>
    {

        public int Compare(MyData x, MyData y)
        {
            try
            {
                double dx = Convert.ToDouble(x.data);
                double dy = Convert.ToDouble(y.data);
                if (dx < dy)
                {
                    return -1;
                }
                else if (dx > dy)
                {
                    return 1;
                }
                else
                {
                    return 0;
                }
            }
            catch(Exception e)
            {
                CaseInsensitiveComparer c = new CaseInsensitiveComparer();
                return c.Compare(x.data,y.data);
            }

        }
    }




}
