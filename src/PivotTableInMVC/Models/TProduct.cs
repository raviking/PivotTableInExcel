using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PivotTableInMVC.Models
{
    public class TProduct
    {
        public string Category { get; set; }
        public string Place { get; set; }
        public string Name { get; set; }
        public long Price { get; set; }
        public long NoOfUnits { get; set; }
    }
}