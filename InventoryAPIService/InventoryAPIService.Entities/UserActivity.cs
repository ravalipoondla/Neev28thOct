using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Inventory.RestAPI.Entities
{
    public class UserActivity
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int quantity { get; set; }
        public decimal price { get; set; }
    }
}
