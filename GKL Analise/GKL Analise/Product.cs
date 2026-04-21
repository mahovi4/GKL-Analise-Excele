using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GKL_Analise
{
    internal class Product : IConstruction
    {
        public string Name { get; }

        public Gabaryte Gabaryte { get; private set; }

        public double Complexity { get; }

        public Stvorka Active { get; }

        public Stvorka Passive { get; }

        public Stvorka Framuga { get; }

        public Stvorka LeftVstavka { get; }

        public Stvorka RightVstavka { get; }

        public Product(string name, Gabaryte gabaryte, double complexity, Stvorka active, Stvorka passive, Stvorka framuga, Stvorka leftVstavka, Stvorka rightVstavka)
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

        public bool IsPassiv => Passive != null;
        public bool IsFramuga => Framuga != null;
        public bool IsLeftVstavka => LeftVstavka != null;
        public bool IsRightVstavka => RightVstavka != null;


        public bool IsEqualConctruction(IConstruction construction)
        {
            if(!construction.Gabaryte.Equals(Gabaryte)) return false;

            return true;
        }

        public override string ToString()
        {
            return $"{Name}({Gabaryte})";
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if(!(obj is Product product)) return false;

            if(!product.Name.Equals(Name)) return false;
            if(!product.Gabaryte.Equals(Gabaryte)) return false;
            if(product.Complexity != Complexity) return false;
            if(!product.Active.Equals(Active)) return false;

            if(product.IsPassiv != IsPassiv) return false;
            if((product.IsPassiv && IsPassiv))
                if(!product.Passive.Equals(Passive)) return false;

            if(product.IsFramuga != IsFramuga) return false;
            if ((product.IsFramuga && IsFramuga))
                if (!product.IsFramuga.Equals(IsFramuga)) return false;

            if (product.IsLeftVstavka != IsLeftVstavka) return false;
            if ((product.IsLeftVstavka && IsLeftVstavka))
                if (!product.IsLeftVstavka.Equals(IsLeftVstavka)) return false;

            if (product.IsRightVstavka != IsRightVstavka) return false;
            if ((product.IsRightVstavka && IsRightVstavka))
                if (!product.RightVstavka.Equals(RightVstavka)) return false;

            return true;
        }

        public void SetGabaryte(Gabaryte gabaryte)
        {
            Gabaryte = gabaryte;
        }
    }
}
