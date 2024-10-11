using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Validar_Bultos
{
    public partial class Observación : Form
    {
        public delegate void EnviarVariableDelegate(string variable, string responsable);
        public event EnviarVariableDelegate EnviarVariableEvent;
        public Observación()
        {
            InitializeComponent();
        }

        private void BtnAgregar_Click(object sender, EventArgs e)
        {
            if (TxtObservacion.Text != string.Empty && TxtVerificado.Text != string.Empty)
            {

                string valor = GlobalSettings.Instance.FacturaActual;
                if (GlobalSettings.Instance.FacturaActual != string.Empty)
                {
                    EnviarVariableEvent(TxtObservacion.Text, TxtVerificado.Text);
                    MessageBox.Show("Agregada correctamente");
                    this.Close();
                }
            }
        }
        public void Revisar(string valor, string responsable, string Factura)
        {
            TxtObservacion.Text = valor;
            TxtVerificado.Text = responsable;
            LbFactura.Text = Factura;
        }
    }
}
