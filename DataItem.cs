using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLS
{
   
    public class DataItem
    {
        public DateTime date { get; set; }
        public decimal ptdValue { get; set; }
        public decimal sumValue { get; set; }
        public DataItem()
        {
        }

        public DataItem(DateTime dt, decimal val, decimal sum)
        {
            date = dt;
            ptdValue = val;
            sumValue = sum;
        }

    }
}
