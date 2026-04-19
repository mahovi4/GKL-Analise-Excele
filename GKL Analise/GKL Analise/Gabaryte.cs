using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GKL_Analise
{
    internal class Gabaryte
    {
        public int Height { get; }
        public int Width { get; }

        public Gabaryte(int height, int width) 
        {  
            Height = height; 
            Width = width; 
        }

        public int Square => 
            Height * Width / 1000000;

        public override string ToString()
        {
            return $"{Height}x{Width}";
        }

        public override int GetHashCode()
        {
            return Height.GetHashCode() ^ Width.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if(!(obj is Gabaryte gabaryte)) return false;

            return gabaryte.Height == Height && gabaryte.Width == Width;
        }
    }
}
