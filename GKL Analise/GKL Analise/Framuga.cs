using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GKL_Analise
{
    internal class Framuga
    {
        public int Square { get; }

        public int Virez { get; }

        public Gabaryte Gabaryte { get; private set; }

        public Framuga(int square, int virez)
        {
            Square = square;
            Virez = virez;
        }

        public void SetGabaryte(Gabaryte gabaryte)
        {
            Gabaryte = gabaryte;
        }
    }
}
