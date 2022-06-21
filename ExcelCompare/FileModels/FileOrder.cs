using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCompare.FileModels
{
    public class FileOrder
    {
        public string ean { get; set; }
        public string ProductName { get; set; }
        public string quantityStocks { get; set; }
        public string priceNettoStocks { get; set; }
        public string priceNettoOffers {get; set;}
        public string quantityToOrder { get; set; }
        
    }
}
