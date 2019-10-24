using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PWC_to_Jobboss
{
    public class ExcelLine
    {
        public string Name { get; set; }
        public string PurchaseDoc { get; set; }
        public string Item { get; set; }
        public string Material { get; set; }
        public string ShortText { get; set; }
        public string E { get; set; }
        public decimal NetPrice { get; set; }
        public string NetPriceCurrency { get; set; }
        public decimal DocItem { get; set; }
        public string DocItemQty { get; set; }
        public decimal OutstQty { get; set; }
        public DateTime StateDelDate { get; set; }
        public DateTime DeliveryDate { get; set; }
        public string P { get; set; }
        public string Status { get; set; }

        public string CompanyName { get; set; }
        public string LnMeso { get; set; }
        public string Description { get; set; }
        public string SO { get; set; }
        public string Job { get; set; }
        public string Shipped { get; set; }
        public string PromisedDate { get; set; }

        public string DeliveryDateString
        {
            get
            {
                return DeliveryDate.ToString("dd-MMM-yyyy");
            }
        }
    }
}
