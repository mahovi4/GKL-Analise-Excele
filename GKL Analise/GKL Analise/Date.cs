using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GKL_Analise
{
    internal class Date
    {
        public int Year { get; }

        public int Month { get; }

        public Dictionary<Product, int> Products { get; set; }

        public Date(int year, int month)
        {
            Year = year;
            Month = month;
        }
    }
}
