using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GKL_Analise
{
    internal interface IConstruction
    {
        string Name { get; }

        Gabaryte Gabaryte { get; }

        void setGabaryte(Gabaryte gabaryte);
    }
}
