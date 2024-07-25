using ClosedXML.Excel;
using System.Reflection;

namespace Validar_Bultos
{
    public partial class Form1 : Form
    {
        //List<List<string>> Facturas = new List<List<string>>();
        List<Folios> Fact = new();
        public Form1()
        {
            InitializeComponent();
            this.Size = new Size(1366, 768); // Ancho y Alto
            this.MinimumSize = new Size(800, 600); // Tamaño mínimo
            this.MaximumSize = new Size(1920, 1080); // Tamaño máximo
        }

        private void BTN_CARGAR_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Files|*.xlsx";
            openFileDialog1.Title = "Seleccionar un archivo Excel";
            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // Obtener la ruta del archivo seleccionado
                string filePath = openFileDialog1.FileName;

                // Leer el archivo Excel usando ClosedXML
                using (var workbook = new XLWorkbook(filePath))
                {
                    // Obtener la primera hoja de cálculo
                    var worksheet = workbook.Worksheet(1);
                    int rowCount = worksheet.RangeUsed().RowCount();
                    Ruta.Text = worksheet.Cell(1, 1).GetString();
                    Ruta.Visible = true;
                    for (int i = 0; i < rowCount; i++)
                    {
                        List<string> dates = new List<string>();
                        string cellValue1 = worksheet.Cell(i + 3, 1).GetString();
                        string cellValue2 = worksheet.Cell(i + 3, 4).GetString();
                        if (cellValue1 != string.Empty)
                        {
                            string Folio_Mod = cellValue1;
                            if (Folio_Mod[1] == 'F' || Folio_Mod[1] == 'O' || Folio_Mod[1] == 'E' || Folio_Mod[1] == 'M' || Folio_Mod[1] == 'A')
                            {
                                string prefix = Folio_Mod.Substring(0, 2);
                                string suffix = Folio_Mod.Substring(2);
                                int contador = 0;
                                for (int p = 0; p < suffix.Count(); p++)
                                {
                                    if (suffix[p] == '0')
                                    {
                                        contador++;
                                    }
                                    else
                                        break;
                                }
                                suffix = suffix.Substring(contador);
                                Folio_Mod = prefix + suffix;
                            }
                            for (int z = 1; z < int.Parse(cellValue2) + 1; z++)
                            {
                                dates.Add(Folio_Mod + "-" + z);
                            }
                            Folios variables = new Folios
                            {

                                Facturas = Folio_Mod,
                                Bultos = cellValue2,
                                ListaFacturas = dates,
                                Escaneados = new List<string>()
                            };
                            Fact.Add(variables);



                        }
                    }
                    // Mostrar los valores leídos

                }
            }
            for (int i = 0; i < Fact.Count; ++i)
            {
                Grid.Columns.Add(i.ToString(), i.ToString());
                if (i == 9)
                    break;
            }
            decimal cantidad = Fact.Count() / 10.00m;
            decimal filas = Math.Ceiling(cantidad);
            for (int i = 1; i < filas; ++i)
            {
                Grid.Rows.Add();
            }
            Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            int j = 0, k = 0, l = -1;
            for (int i = 0; i < Fact.Count; i++)
            {
                Grid.Rows[k].Height = 50;
                Grid.Rows[k].Cells[j].Value = Fact[i].Facturas;
                Fact[i].Xgrid = k;
                Fact[i].Ygrid = j;
                Grid.Rows[k].Cells[j].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                j++;
                if (j == 10)
                {
                    if (filas - 2 > (l))
                    {
                        k++;
                    }
                    l++;
                    j = 0;
                }
            }
            BTN_CARGAR.Hide();
            label1.Visible = false;
            TxtEscanear.Visible = true;
            TxtEscanear.Focus();
            Txt1.Visible = true;
            BtnAgregar.Visible = true;
            Grid.CurrentCell = null;
            label2.Visible = true;
            label3.Visible = true;
            Bultos.Visible = true;
            Contador.Visible = true;
            FACTURA.Visible = true;
            label4.Visible = true;
            Total.Visible = true;

            //var workbook = new XLWorkbook();
            //var worksheet = workbook.Worksheets.Add("Sample Sheet");

            //// Escribir datos en celdas
            //worksheet.Cell(1, 1).Value = "Hello";
            //worksheet.Cell(1, 2).Value = "World";

            //// Guardar el archivo
            //string filePath = "C:\\Datos_Surtido\\Sample.xlsx";
            //workbook.SaveAs(filePath);
        }

        private void BtnAgregar_Click(object sender, EventArgs e)
        {
            string Folio = "";
            if (TxtEscanear.Text != string.Empty)
            {
                for (int i = 0; i < TxtEscanear.TextLength; i++)
                {
                    if (TxtEscanear.Text[i] == '-')
                    {
                        Folio = TxtEscanear.Text.Substring(0, i);
                        break;
                    }

                }
                bool bandera = false;
                for (int i = 0; i < Fact.Count; ++i)
                {
                    if (Folio == Fact[i].Facturas)
                    {
                        Bultos.Text = Fact[i].ListaFacturas.Count().ToString();
                        FACTURA.Text = Folio;
                        for (int x = 0; x < Fact[i].ListaFacturas.Count; ++x)
                        {
                            if (TxtEscanear.Text == Fact[i].ListaFacturas[x])
                            {
                                bandera = true;
                                if (Fact[i].Escaneados.Count == 0)
                                {
                                    Fact[i].Escaneados.Add(TxtEscanear.Text);
                                }
                                else
                                {
                                    for (int u = 0; u < Fact[i].Escaneados.Count; ++u)
                                    {
                                        if (Fact[i].Escaneados[u] == TxtEscanear.Text)
                                        {
                                            MessageBox.Show("Este Bulto ya se escaneó", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            return;
                                        }
                                    }
                                    Fact[i].Escaneados.Add(TxtEscanear.Text);

                                }
                                Contador.Text = ((Fact[i].ListaFacturas.Count()) - (Fact[i].Escaneados.Count())).ToString();
                            }

                        }
                        if (bandera == false)
                        {
                            MessageBox.Show("Este Bulto no esta registrado en el reporte", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        if (Fact[i].Escaneados.Count == Fact[i].ListaFacturas.Count)
                        {
                            //MessageBox.Show("Encontrado con " + Fact[i].Bultos);
                            DataGridViewCell cell2 = Grid.Rows[Fact[i].Xgrid].Cells[Fact[i].Ygrid];
                            //cell.DefaultCellStyle.BackColor = System.Drawing.Color.LightBlue;
                            cell2.Style.BackColor = System.Drawing.Color.Green;
                        }
                        else
                        {

                            //MessageBox.Show("Encontrado con " + Fact[i].Bultos);
                            DataGridViewCell cell = Grid.Rows[Fact[i].Xgrid].Cells[Fact[i].Ygrid];
                            //cell.DefaultCellStyle.BackColor = System.Drawing.Color.LightBlue;
                            cell.Style.BackColor = System.Drawing.Color.LightBlue;
                        }

                    }
                    //else
                    //{
                    //    MessageBox.Show("Folio no encontrado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    return;

                    //}

                }
                if (bandera == false)
                {
                    MessageBox.Show("Folio no encontrado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                TxtEscanear.Text = "";
                TxtEscanear.Focus();

            }
            else
                MessageBox.Show("Agrega un folio", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void TxtEscanear_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // Evita que se produzca el sonido de Windows al presionar Enter
                BtnAgregar.Focus();
            }
        }
    }
}