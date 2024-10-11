using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Validar_Bultos
{
    public class GlobalSettings
    {
        private static GlobalSettings instance;
        public double BultosTotales {  get; set; }
        public int Orden {  get; set; }
        public string filePath { get; set; }
        public double sumaimportes { get; set; }
        public double sumabultos { get; set; }
        public double FacturasTotales { get; set; }
        public string FacturaActual {  get; set; }
        public bool Cargado {  get; set; }
        private GlobalSettings()
        {
            

        }
        public static GlobalSettings Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new GlobalSettings();
                }
                return instance;
            }
        }
    }
    public class Folios
    {
        public int Orden {  get; set; }
        public bool Completo {  get; set; }
        public bool Incompleto { get; set; }
        public string Facturas { get; set; }
        public string Observacion { get; set; }
        public string Responsable { get; set; }
        public double Bultos { get; set; }
        public string Ciudad { get; set; }
        public string Nombre { get; set; }
        public double Importe { get; set; }
        public string Condiciones { get; set; }
        public int Clave { get; set; }
        public List<string> ListaFacturas { get; set; }
        public List<string> Escaneados { get; set; }
        public int Xgrid { get; set; }
        public int Ygrid { get; set; }
    }
}
