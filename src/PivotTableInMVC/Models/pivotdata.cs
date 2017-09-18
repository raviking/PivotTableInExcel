using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PivotTableInMVC.Models
{
    public class pivotdata
    {
        public List<TProduct> PivotData()
        {
            List<TProduct> _list = new List<TProduct>()
            {
                new TProduct {Category="Clothing",Place="Hyderabad",Name="LEVIS",Price=3000,NoOfUnits=52 },
                new TProduct {Category="Clothing",Place="Hyderabad",Name="Buffallo",Price=10000,NoOfUnits=12 },
                new TProduct {Category="Clothing",Place="Banglore",Name="FM",Price=3200,NoOfUnits=5 },
                new TProduct {Category="Clothing",Place="Banglore",Name="PUMA",Price=6400,NoOfUnits=10 },
                new TProduct {Category="Clothing",Place="Banglore",Name="LEVIS",Price=3400,NoOfUnits=20 },
                new TProduct {Category="Clothing",Place="Banglore",Name="Buffallo",Price=34400,NoOfUnits=30 },
                new TProduct {Category="Electronics",Place="Banglore",Name="IPhone",Price=72000,NoOfUnits=1 },
                new TProduct {Category="Electronics",Place="Banglore",Name="LED TV",Price=20000,NoOfUnits=4 },
                new TProduct {Category="Electronics",Place="Banglore",Name="SAMSUNG",Price=300000,NoOfUnits=5 },
                new TProduct {Category="Electronics",Place="Banglore",Name="IPhone",Price=7200,NoOfUnits=1 },
                new TProduct {Category="Electronics",Place="Hyderabad",Name="Fridge",Price=150000,NoOfUnits=10 },
                new TProduct {Category="Electronics",Place="Hyderabad",Name="Laptops",Price=40000,NoOfUnits=15 },
                new TProduct {Category="Electronics",Place="Hyderabad",Name="Laptops",Price=30000,NoOfUnits=6 },
                new TProduct {Category="Electronics",Place="Hyderabad",Name="Laptops",Price=78347,NoOfUnits=8 },
            };
            return _list;
        }
    }
}