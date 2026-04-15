using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GKL_Analise
{
    internal class Product
    {
        public string Name { get; }

        public Gabaryte Gabaryte { get; }

        public int Complexity { get; }

        public Stvorka Active { get; }

        public Stvorka Passive { get; }

        public Framuga Framuga { get; }

        public Framuga LeftVstavka { get; }

        public Framuga RightVstavka { get; }

        public Product(string name, Gabaryte gabaryte, int complexity, Stvorka active, Stvorka passive, Framuga framuga, Framuga leftVstavka, Framuga rightVstavka)
        {
            Name = name;
            Gabaryte = gabaryte;
            Complexity = complexity;
            Active = active;
            Passive = passive;
            Framuga = framuga;
            LeftVstavka = leftVstavka;
            RightVstavka = rightVstavka;
        }
    }
}
