using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Inventory.RestAPI.Entities
{
    public class RawMaterialInventory
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string ProductImage { get; set; }
        public int Quantity { get; set; }
        public double UnitPrice { get; set; }
        public int Threshold { get; set; }
    }
}
