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

        public double VirezSq { get; }

        public Stvorka(string name, double square, EGabaryteDirection direction, int delimeter, double virezSq)
        {
            Name = name;

            var dim = (int)(square * 1000000 / delimeter);
            Gabaryte = new Gabaryte(
                direction == EGabaryteDirection.Height ? delimeter : dim, 
                direction == EGabaryteDirection.Width ? delimeter : dim);

            VirezSq = virezSq;
        }

        public Stvorka(string name, Gabaryte gabaryte, double virez)
        {
            Name = name;
            Gabaryte = gabaryte;
            VirezSq = virez;
        }

        public override string ToString()
        {
            return $"{Name} - {Gabaryte} [{VirezSq}]";
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Stvorka stvorka)) return false;

            return stvorka.Gabaryte.Equals(Gabaryte) && stvorka.VirezSq.Equals(VirezSq);
        }

        public void SetGabaryte(Gabaryte gabaryte)
        {
            Gabaryte = gabaryte;
        }
    }
}
