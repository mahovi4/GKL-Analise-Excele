using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GKL_Analise
{
    internal class Framuga : IConstruction, IPolotno
    {
        public string Name { get; }

        public int Square { get; }

        public int VirezSq { get; }

        public Gabaryte Gabaryte { get; private set; }


        public Framuga(string name, int square, int virez)
        {
            Name = name;
            Square = square;
            VirezSq = virez;
        }

        public void SetGabaryte(Gabaryte gabaryte)
        {
            Gabaryte = gabaryte;
        }

        public override string ToString()
        {
            return $"{Gabaryte} [{VirezSq}]";
        }

        public override int GetHashCode()
        {
            return Name.GetHashCode() + Gabaryte.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if(!(obj is Framuga framuga)) return false;

            return framuga.Gabaryte.Equals(Gabaryte) && framuga.VirezSq == VirezSq;
        }

        public void setGabaryte(Gabaryte gabaryte)
        {
            throw new NotImplementedException();
        }
    }
}
