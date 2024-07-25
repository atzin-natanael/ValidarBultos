using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Validar_Bultos
{
    public class GlobalSettings
    {

    }
    public class Folios
    {
        public string Facturas { get; set; }
        public string Bultos { get; set; }

        public List<string> ListaFacturas { get; set; }
        public List<string> Escaneados { get; set; }
        public int Xgrid { get; set; }
        public int Ygrid { get; set; }
    }
}
