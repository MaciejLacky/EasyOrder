using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCompare.FileModels
{
    public  class FileStocks
    {
        public string ean { get; set; }
        public string quantity { get; set; }
        public string priceNetto { get; set; }
        public string productName { get; set; }

    }
}
