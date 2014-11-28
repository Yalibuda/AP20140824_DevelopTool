using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Tool
{
    /*
     * 一般不會直接使用到此類別，這只是紀錄 Stack Data 中的某一個資料，
     * 已經設計 StackedData 類別，讓使用者直接輸入 Array 來處理 Stack
     * Data(此類別中即使用 GroupedData 類別)。
     * 
     */ 
    public class GroupedData
    {
        public object Data { set; get; }
        public object Subscript { set; get; }
    }
}
