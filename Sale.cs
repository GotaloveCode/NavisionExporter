using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ayoti
{
    class Sale
    {
        public int Id { get; set; }
        public string Product { get; set; }
        public string Category { get; set; }
        public double Quantity { get; set; }
        public string Customer { get; set; }
        public string SalesPerson { get; set; }
        public bool Uploaded { get; set; }
    }
}
