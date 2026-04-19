using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GKL_Analise
{
    internal class Stvorka : IConstruction, IPolotno
    {
        public string Name { get; }

        public Gabaryte Gabaryte { get; private set; }

        public int VirezSq { get; }

        public Stvorka(string name, int square, EGabaryteDirection direction, int delimeter, int virezSq)
        {
            Name = name;

            var dim = square * 1000000 / delimeter;
            Gabaryte = new Gabaryte(
                direction == EGabaryteDirection.Height ? delimeter : dim, 
                direction == EGabaryteDirection.Width ? delimeter : dim);

            VirezSq = virezSq;
        }

        public Stvorka(string name, Gabaryte gabaryte, int virez)
        {
            Name = name;
            Gabaryte = gabaryte;
            VirezSq = virez;
        }

        public override string ToString()
        {
            return $"{Gabaryte} [{VirezSq}]";
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Stvorka stvorka)) return false;

            return stvorka.Gabaryte.Equals(this.Gabaryte) && stvorka.VirezSq == VirezSq;
        }

        public void setGabaryte(Gabaryte gabaryte)
        {
            Gabaryte = gabaryte;
        }
    }
}
