using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GKL_Analise
{
    internal class Stvorka
    {
        public Gabaryte Gabaryte { get; }

        public int Virez { get; }

        public Stvorka(Gabaryte gabaryte, int virez)
        {
            Gabaryte = gabaryte;
            Virez = virez;
        }
    }
}
