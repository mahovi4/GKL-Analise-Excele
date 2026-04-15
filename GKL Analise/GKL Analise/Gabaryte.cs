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
    }
}
