using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.CodeDom.Compiler;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using Paragraph = iTextSharp.text.Paragraph;

namespace Validar_Bultos
{
    public partial class Form1 : Form
    {
        //List<List<string>> Facturas = new List<List<string>>();
        List<Folios> Fact = new();
        Observación observacion;
        public Form1()
        {
            InitializeComponent();
            GlobalSettings.Instance.Cargado = false;
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
                    CorteBtn.Visible = true;
                    for (int i = 0; i < rowCount; i++)
                    {
                        string Folio_Fact = worksheet.Cell(i + 3, 1).GetString();
                        if (Folio_Fact != string.Empty)
                        {
                            List<string> dates = new List<string>();
                            string Ciudad = worksheet.Cell(i + 3, 2).GetString();
                            string Nombre = worksheet.Cell(i + 3, 3).GetString();
                            double Bultos = worksheet.Cell(i + 3, 4).GetDouble();
                            double Importe = worksheet.Cell(i + 3, 5).GetDouble();
                            string Condiciones = worksheet.Cell(i + 3, 6).GetString();
                            int Clave = 0;
                            if (worksheet.Cell(i + 3, 8).GetValue<string>() == string.Empty)
                            {
                                MessageBox.Show("REVISA BIEN LOS DATOS DEL CLIENTE", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                this.Close();
                                return;
                            }
                            else
                                Clave = worksheet.Cell(i + 3, 8).GetValue<int>();

                            //else
                            //{
                            //}
                            GlobalSettings.Instance.BultosTotales += (Bultos);
                            string Folio_Mod = Folio_Fact;
                            if (Folio_Mod[1] == 'F' || Folio_Mod[1] == 'O' || Folio_Mod[1] == 'E' || Folio_Mod[1] == 'M' || Folio_Mod[1] == 'A' || Folio_Mod[1] == 'C' || Folio_Mod[1] == 'L')
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
                            if (Folio_Mod[0] == 'A' || Folio_Mod[0] == 'B' || Folio_Mod[0] == 'C' || Folio_Mod[0] == 'D' || Folio_Mod[0] == 'E' || Folio_Mod[0] == 'F' || Folio_Mod[0] == 'G' || Folio_Mod[0] == 'H')
                            {
                                string prefix = Folio_Mod.Substring(0, 1);
                                string suffix = Folio_Mod.Substring(1);
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
                            for (int z = 1; z < (Bultos) + 1; z++)
                            {
                                dates.Add(Folio_Mod + "#" + z);
                            }

                            Folios variables = new Folios
                            {

                                Facturas = Folio_Mod,
                                Nombre = Nombre,
                                Clave = Clave,
                                Ciudad = Ciudad,
                                Importe = Importe,
                                Condiciones = Condiciones,
                                Bultos = Bultos,
                                ListaFacturas = dates,
                                Escaneados = new List<string>()
                            };
                            Fact.Add(variables);



                        }
                    }
                    // Mostrar los valores leídos

                }
            }
            else
                return;
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
            int j = 0, k = 0, l = -1, alt = 0;
            if (Fact.Count > 100)
                alt = 40;
            else
                alt = 50;
            GlobalSettings.Instance.FacturasTotales = Fact.Count;
            TotalF.Text = GlobalSettings.Instance.FacturasTotales.ToString();
            for (int i = 0; i < Fact.Count; i++)
            {
                Grid.Rows[k].Height = alt;
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
            Total.Text = GlobalSettings.Instance.BultosTotales.ToString();
            BTN_CARGAR.Hide();
            label1.Visible = false;
            TxtEscanear.Visible = true;
            TxtEscanear.Focus();
            Txt1.Visible = true;
            BtnAgregar.Visible = true;
            label2.Visible = true;
            label3.Visible = true;
            Bultos.Visible = true;
            Contador.Visible = true;
            FACTURA.Visible = true;
            label4.Visible = true;
            Total.Visible = true;
            TotalF.Visible = true;
            label5.Visible = true;
            Grid.CurrentCell = null;
            GlobalSettings.Instance.Cargado = true;
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
                    if (TxtEscanear.Text[i] == '#')
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
                                    GlobalSettings.Instance.BultosTotales--;
                                    Total.Text = GlobalSettings.Instance.BultosTotales.ToString();
                                }
                                else
                                {
                                    for (int u = 0; u < Fact[i].Escaneados.Count; ++u)
                                    {
                                        if (Fact[i].Escaneados[u] == TxtEscanear.Text)
                                        {
                                            MessageBox.Show("Este Bulto ya se escaneó", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            TxtEscanear.Select(0, TxtEscanear.TextLength);
                                            TxtEscanear.Focus();
                                            return;
                                        }
                                    }
                                    Fact[i].Escaneados.Add(TxtEscanear.Text);
                                    GlobalSettings.Instance.BultosTotales--;
                                    Total.Text = GlobalSettings.Instance.BultosTotales.ToString();

                                }
                                Contador.Text = ((Fact[i].ListaFacturas.Count()) - (Fact[i].Escaneados.Count())).ToString();
                                if (Fact[i].ListaFacturas.Count == Fact[i].Escaneados.Count)
                                {
                                    if (Fact[i].Orden == 0)
                                    {
                                        GlobalSettings.Instance.Orden += 1;
                                        for (int u = 0; u < Fact.Count; ++u)
                                        {
                                            if (Fact[i].Clave == Fact[u].Clave)
                                            {
                                                Fact[u].Orden = GlobalSettings.Instance.Orden;
                                            }
                                        }
                                    }
                                    GlobalSettings.Instance.FacturasTotales--;
                                    TotalF.Text = GlobalSettings.Instance.FacturasTotales.ToString();
                                }


                            }

                        }
                        if (bandera == false)
                        {
                            MessageBox.Show("Este Bulto no esta registrado en el reporte", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            TxtEscanear.Select(0, TxtEscanear.TextLength);
                            TxtEscanear.Focus();
                            return;
                        }
                        if (Fact[i].Escaneados.Count == Fact[i].ListaFacturas.Count)
                        {
                            //MessageBox.Show("Encontrado con " + Fact[i].Bultos);
                            DataGridViewCell cell2 = Grid.Rows[Fact[i].Xgrid].Cells[Fact[i].Ygrid];
                            //cell.DefaultCellStyle.BackColor = System.Drawing.Color.LightBlue;
                            cell2.Style.BackColor = System.Drawing.Color.Green;
                            Fact[i].Completo = true;
                            Fact[i].Incompleto = false;
                        }
                        else
                        {

                            //MessageBox.Show("Encontrado con " + Fact[i].Bultos);
                            DataGridViewCell cell = Grid.Rows[Fact[i].Xgrid].Cells[Fact[i].Ygrid];
                            //cell.DefaultCellStyle.BackColor = System.Drawing.Color.LightBlue;
                            cell.Style.BackColor = System.Drawing.Color.LightBlue;
                            Fact[i].Incompleto = true;
                        }

                        Grid.ClearSelection();

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
            if (e.Control && e.KeyCode == Keys.G)
            {
                Terminar();
            }

        }

        private void Grid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0) // Verificar que no se hizo clic en el encabezado
            {
                string Valor = Grid[e.ColumnIndex, e.RowIndex].Value.ToString();
                for (int i = 0; i < Fact.Count; ++i)
                {
                    if (Valor == Fact[i].Facturas)
                    {
                        FACTURA.Text = Fact[i].Facturas;
                        int cont1 = Fact[i].ListaFacturas.Count;
                        Bultos.Text = cont1.ToString();
                        Contador.Text = ((Fact[i].ListaFacturas.Count()) - (Fact[i].Escaneados.Count())).ToString();
                    }
                }
                if (Grid.CurrentCell.Style.BackColor == System.Drawing.Color.Green)
                {
                    Grid.ClearSelection();
                }
                if (Grid.CurrentCell.Style.BackColor == System.Drawing.Color.LightBlue)
                {
                    Grid.ClearSelection();
                }
            }
        }

        private void Grid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                //Tabla.ClearSelection();
                //Tabla.Rows[e.RowIndex].Selected = true;
                string Factura_tabla = Grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                GlobalSettings.Instance.FacturaActual = Factura_tabla;
                ContextMenuStrip menu = new ContextMenuStrip();

                // Agregar opciones al menú
                ToolStripMenuItem AddObservacionMenuItem = new ToolStripMenuItem("Agregar Observación");
                ToolStripMenuItem verObservacion = new ToolStripMenuItem("Ver Observacion");
                ToolStripMenuItem verClienteMenuItem = new ToolStripMenuItem("Ver Cliente");
                ToolStripMenuItem verBultosMenuItem = new ToolStripMenuItem("Ver Bultos");
                menu.Items.Add(AddObservacionMenuItem);
                menu.Items.Add(verObservacion);
                menu.Items.Add(verClienteMenuItem);
                menu.Items.Add(verBultosMenuItem);
                // Manejar el evento ItemClicked del menú contextual
                menu.ItemClicked += (s, args) =>
                {

                    if (args.ClickedItem == AddObservacionMenuItem)
                    {
                        menu.Close();
                        string obseravacion = "";
                        string responsable = "";
                        string Factura = "";
                        if (observacion == null || observacion.IsDisposed)
                        {
                            for (int i = 0; i < Fact.Count; ++i)
                            {
                                if (Fact[i].Facturas == Factura_tabla)
                                {
                                    obseravacion = Fact[i].Observacion;
                                    responsable = Fact[i].Responsable;
                                    Factura = Fact[i].Facturas;
                                }


                            }
                            // Crear una nueva instancia si no existe o está disposada
                            Observación form2 = new Observación();
                            form2.Revisar(obseravacion, responsable, Factura);
                            form2.EnviarVariableEvent += new Observación.EnviarVariableDelegate(RecibirVariable);
                            form2.Show();
                        }
                    }
                    else if (args.ClickedItem == verObservacion)
                    {
                        // Lógica para la opción "Ver Ubicación"
                        // Aquí puedes ejecutar la acción correspondiente
                        menu.Close();
                        for (int i = 0; i < Fact.Count; ++i)
                        {
                            if (Fact[i].Facturas == Factura_tabla)
                            {
                                if (Fact[i].Observacion != null)
                                    MessageBox.Show(Fact[i].Observacion + "\n" + "Verificado por: " + Fact[i].Responsable.ToString(), "Observación");
                                else
                                    MessageBox.Show("No hay Observación en esta factura", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            }


                        }
                    }
                    else if (args.ClickedItem == verClienteMenuItem)
                    {
                        // Lógica para la opción "Ver Ubicación"
                        // Aquí puedes ejecutar la acción correspondiente
                        menu.Close();
                        for (int i = 0; i < Fact.Count; ++i)
                        {
                            if (Fact[i].Facturas == Factura_tabla)
                            {

                                MessageBox.Show(Fact[i].Nombre + "\n", "Nombre");

                            }


                        }
                    }
                    else if (args.ClickedItem == verBultosMenuItem)
                    {
                        // Lógica para la opción "Ver Ubicación"
                        // Aquí puedes ejecutar la acción correspondiente
                        menu.Close();
                        for (int i = 0; i < Fact.Count; ++i)
                        {
                            if (Fact[i].Facturas == Factura_tabla)
                            {

                                MessageBox.Show(Fact[i].Bultos + "\n", "Bultos");

                            }


                        }
                    }
                };

                menu.Show(Cursor.Position);

            }
        }
        private void RecibirVariable(string variable, string responsable)
        {
            for (int i = 0; i < Fact.Count; ++i)
            {
                if (GlobalSettings.Instance.FacturaActual == Fact[i].Facturas)
                {
                    Fact[i].Observacion = variable;
                    Fact[i].Responsable = responsable;
                }
            }
        }
        private void Grid_CurrentCellChanged(object sender, EventArgs e)
        {
            if (GlobalSettings.Instance.Cargado == true)
            {
                int rowIndex = Grid.CurrentCell.RowIndex;
                int colIndex = Grid.CurrentCell.ColumnIndex;
                string Valor = "";
                if (Grid.Rows[rowIndex].Cells[colIndex].Value != null)
                {
                    Valor = Grid.Rows[rowIndex].Cells[colIndex].Value.ToString();
                    GlobalSettings.Instance.FacturaActual = Valor;
                }
                else
                    return;
                for (int i = 0; i < Fact.Count; ++i)
                {
                    if (Valor == Fact[i].Facturas)
                    {
                        FACTURA.Text = Fact[i].Facturas;
                        int cont1 = Fact[i].ListaFacturas.Count;
                        Bultos.Text = cont1.ToString();
                        Contador.Text = ((Fact[i].ListaFacturas.Count()) - (Fact[i].Escaneados.Count())).ToString();
                    }
                }
                if (Grid.CurrentCell.Style.BackColor == System.Drawing.Color.Green)
                {
                    Grid.ClearSelection();
                }
                if (Grid.CurrentCell.Style.BackColor == System.Drawing.Color.LightBlue)
                {
                    Grid.ClearSelection();
                }

            }
        }

        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.G)
            {
                Terminar();
            }


        }
        public void Corte()
        {
            bool verificar = false;
            var facturasCompletasIsi = Fact.Where(f => f.Facturas.StartsWith("IF") && f.Completo || f.Facturas.StartsWith("IF") && f.Incompleto).OrderByDescending(f => f.Orden).ToList();
            var copia1 = facturasCompletasIsi;
            if (facturasCompletasIsi.Count == 0)
            {
                verificar = true;
            }
            Fact.RemoveAll(f => facturasCompletasIsi.Any(fc => fc.Facturas == f.Facturas));
            var facturasCompletas = Fact.Where(f => f.Completo || f.Incompleto).OrderByDescending(f => f.Orden).ToList();
            if (facturasCompletas.Count == 0 && verificar == true)
            {
                MessageBox.Show("No has escaneado nada", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop); return;
            }
            GlobalSettings.Instance.Cargado = false;
            MessageBox.Show("ASIGNALE UN NOMBRE AL ARCHIVO Y SU UBICACIÓN");
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            saveFileDialog1.InitialDirectory = desktopFolder;
            saveFileDialog1.Title = "ASIGNAR NOMBRE";
            saveFileDialog1.FileName = Ruta.Text + " " + DateTime.Now.ToString("dddd dd-MM-yy");
            saveFileDialog1.Filter = "Archivos de Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*";
            saveFileDialog1.DefaultExt = "xlsx";
            saveFileDialog1.AddExtension = true;
            int fila = 0;
            int fila1 = 0;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // Aquí puedes escribir el código para guardar el archivo en la ubicación seleccionada
                GlobalSettings.Instance.filePath = saveFileDialog1.FileName;
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Reporte Embarques");

                // Escribir datos en celdas
                string imagePath = "C:\\Img\\Cornejo.jpg";
                var picture = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 1)).Scale(0.20);
                worksheet.Row(1).Height = 100;
                worksheet.Cell(4, 1).Value = "Factura";
                worksheet.Cell(2, 2).Value = "Efectivo:";
                worksheet.Cell(3, 2).Value = "F. Devueltas: ";
                worksheet.Cell(4, 2).Value = "Ciudad";
                worksheet.Cell(4, 3).Value = "Bultos";
                worksheet.Cell(4, 4).Value = "Clave";
                worksheet.Cell(4, 5).Value = "Cliente";
                worksheet.Cell(4, 6).Value = "Importe";
                worksheet.Cell(4, 7).Value = "Condiciones";
                worksheet.Cell(4, 8).Value = "EF";
                worksheet.Cell(4, 9).Value = "CR";
                worksheet.Cell(4, 10).Value = "T/D";
                worksheet.Cell(4, 11).Value = "CH";
                worksheet.Cell(4, 12).Value = "Observaciones";
                worksheet.Cell(4, 13).Value = "Firma Cliente";
                worksheet.Cell(2, 6).Value = "Operador:";
                worksheet.Cell(3, 6).Value = "Ruta:";
                worksheet.Cell(2, 12).Value = "Unidad:";
                worksheet.Cell(3, 12).Value = "Fecha:";
                worksheet.Column(1).Width = 7.57;
                worksheet.Column(2).Width = 22;
                worksheet.Column(3).Width = 5;
                worksheet.Column(4).Width = 6.57;
                worksheet.Column(5).Width = 35;
                worksheet.Column(6).Width = 11.14;
                worksheet.Column(7).Width = 12;
                worksheet.Column(8).Width = 3;
                worksheet.Column(9).Width = 3;
                worksheet.Column(10).Width = 3;
                worksheet.Column(11).Width = 3;
                worksheet.Column(12).Width = 13;
                worksheet.Column(13).Width = 25;
                worksheet.Range(1, 1, 1, 13).Merge();
                worksheet.Range(2, 1, 3, 1).Merge();
                worksheet.Range(2, 3, 2, 4).Merge();
                worksheet.Range(3, 3, 3, 4).Merge();
                worksheet.Range(2, 5, 3, 5).Merge();
                worksheet.Range(2, 7, 2, 11).Merge();
                worksheet.Range(3, 7, 3, 11).Merge();
                var rangoDeCeldas = worksheet.Range(2, 6, 3, 6);
                rangoDeCeldas.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                var rangoDeCeldas2 = worksheet.Range(2, 12, 3, 12);
                rangoDeCeldas2.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas2.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                var rangoDeCeldas3 = worksheet.Range(2, 2, 3, 2);
                rangoDeCeldas3.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas3.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Mover la imagen al centro del rango combinado
                picture.MoveTo(worksheet.Cell(1, 2), 200, 40);
                for (int i = 0; i < 13; ++i)
                {
                    var cell3 = worksheet.Cell(4, i + 1);
                    cell3.Style.Fill.BackgroundColor = XLColor.SteelBlue; // Color de fondo
                    cell3.Style.Font.Bold = true; // Fuente en negrita
                    cell3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }
                int celdas = 5;
                // Filtrar las facturas que estén completas (Completo == true)
                for (int i = 0; i < facturasCompletas.Count; ++i)
                {
                    int ckey = facturasCompletas[i].Clave;
                    worksheet.Row(i + celdas).Height = 22;
                    var celdafact = worksheet.Cell(i + celdas, 1);
                    celdafact.Style.Font.Bold = true; // Fuente en negrita
                    var celdabulto = worksheet.Cell(i + celdas, 3);
                    celdabulto.Style.Font.Bold = true; // Fuente en negrita
                    worksheet.Cell(i + celdas, 1).Value = facturasCompletas[i].Facturas;
                    worksheet.Cell(i + celdas, 2).Value = facturasCompletas[i].Ciudad;
                    worksheet.Cell(i + celdas, 3).Value = facturasCompletas[i].Bultos;
                    worksheet.Cell(i + celdas, 4).Value = facturasCompletas[i].Clave;
                    worksheet.Cell(i + celdas, 5).Value = facturasCompletas[i].Nombre;
                    worksheet.Cell(i + celdas, 6).Value = facturasCompletas[i].Importe;
                    worksheet.Cell(i + celdas, 7).Value = facturasCompletas[i].Condiciones;
                    worksheet.Cell(i + celdas, 8).Value = "EF";
                    worksheet.Cell(i + celdas, 9).Value = "CR";
                    worksheet.Cell(i + celdas, 10).Value = "T/D";
                    worksheet.Cell(i + celdas, 11).Value = "CH";
                    worksheet.Cell(i + celdas, 12).Value = "";
                    worksheet.Cell(i + celdas, 13).Value = "Recibí " + facturasCompletas[i].Bultos + " Bultos:";
                    if (i != facturasCompletas.Count - 1)
                    {
                        if (ckey == facturasCompletas[i + 1].Clave)
                        {
                            GlobalSettings.Instance.sumaimportes += facturasCompletas[i].Importe;
                            GlobalSettings.Instance.sumabultos += facturasCompletas[i].Bultos;
                        }
                        else
                        {
                            GlobalSettings.Instance.sumaimportes += facturasCompletas[i].Importe;
                            GlobalSettings.Instance.sumabultos += facturasCompletas[i].Bultos;
                            celdas++;
                            worksheet.Cell(i + celdas, 1).Value = "TOTAL DE " + facturasCompletas[i].Nombre;
                            var cell2 = worksheet.Cell(i + celdas, 1);
                            var cell3 = worksheet.Range(i + celdas, 6, i + celdas, 7);
                            worksheet.Range(i + celdas, 1, i + celdas, 5).Merge();
                            worksheet.Range(i + celdas, 8, i + celdas, 13).Merge();
                            worksheet.Cell(i + celdas, 6).Value = GlobalSettings.Instance.sumaimportes;
                            worksheet.Cell(i + celdas, 7).Value = GlobalSettings.Instance.sumabultos + " BULTOS";
                            cell2.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                            cell2.Style.Font.Bold = true; // Fuente en negrita
                            cell2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell3.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                            cell3.Style.Font.Bold = true; // Fuente en negrita
                            cell3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            GlobalSettings.Instance.sumaimportes = 0;
                            GlobalSettings.Instance.sumabultos = 0;
                        }
                    }
                    else
                    {
                        GlobalSettings.Instance.sumaimportes += facturasCompletas[i].Importe;
                        GlobalSettings.Instance.sumabultos += facturasCompletas[i].Bultos;
                        celdas++;
                        worksheet.Cell(i + celdas, 1).Value = "TOTAL DE " + facturasCompletas[i].Nombre;
                        var cell2 = worksheet.Cell(i + celdas, 1);
                        var cell3 = worksheet.Range(i + celdas, 6, i + celdas, 7);
                        worksheet.Range(i + celdas, 1, i + celdas, 5).Merge();
                        worksheet.Range(i + celdas, 8, i + celdas, 13).Merge();
                        worksheet.Cell(i + celdas, 6).Value = GlobalSettings.Instance.sumaimportes;
                        worksheet.Cell(i + celdas, 7).Value = GlobalSettings.Instance.sumabultos + " BULTOS";
                        cell2.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                        cell2.Style.Font.Bold = true; // Fuente en negrita
                        cell2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell3.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                        cell3.Style.Font.Bold = true; // Fuente en negrita
                        cell3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        GlobalSettings.Instance.sumaimportes = 0;
                        GlobalSettings.Instance.sumabultos = 0;
                    }
                    var range = worksheet.Range("A1:M" + (i + celdas));
                    range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    fila = i + celdas;


                }
                double total = 0;
                double importes = 0;
                for (int l2 = 0; l2 < facturasCompletas.Count; ++l2)
                {
                    total += facturasCompletas[l2].Bultos;
                    importes += facturasCompletas[l2].Importe;

                }
                worksheet.Range(fila + 1, 1, fila + 1, 5).Merge();
                worksheet.Cell(fila + 1, 7).Value = total;
                worksheet.Cell(fila + 1, 6).Value = importes;
                worksheet.Cell(fila + 1, 1).Value = "TOTALES";
                var rangoDeCeldas4 = worksheet.Range(fila + 1, 1, fila + 1, 7);
                rangoDeCeldas4.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas4.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas4.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rangoDeCeldas4.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                rangoDeCeldas4.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                rangoDeCeldas4.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                rangoDeCeldas4.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                // Guardar el archivo

                //workbook.SaveAs(GlobalSettings.Instance.filePath);
                //Process.Start(new ProcessStartInfo(GlobalSettings.Instance.filePath) { UseShellExecute = true });
                iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.LETTER.Rotate());
                string file = "C:\\DatosRutas\\" + Ruta.Text + " " + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".pdf";

                PdfWriter.GetInstance(doc, new FileStream(file, FileMode.Create));
                doc.Open();
                float[] columnWidths = new float[] { 14f, 25f, 40f, 20f, 15f, 35f }; // Asumiendo que la segunda columna tendrá un ancho personalizado
                                                                                     //table.SetWidths(columnWidths);
                                                                                     // Crear una tabla para los datos faltantes
                PdfPTable table = new PdfPTable(6);
                PdfPCell cel = new PdfPCell(new Phrase("Reporte de observaciones y faltantes " + Ruta.Text, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 16f, iTextSharp.text.Font.BOLD)));
                cel.Colspan = 6;
                cel.HorizontalAlignment = 1;
                table.AddCell(cel);
                table.AddCell("Folio");
                table.AddCell("Ciudad");
                table.AddCell("Cliente");
                table.AddCell("Bultos Escaneados");
                table.AddCell("Bultos Recibidos");
                table.AddCell("Observación");
                bool bandera = false;
                for (int i = 0; i < facturasCompletas.Count; ++i)
                {
                    if (facturasCompletas[i].Observacion != null || (facturasCompletas[i].Escaneados.Count() != facturasCompletas[i].ListaFacturas.Count()))
                    {
                        if (facturasCompletas[i].Observacion == null)
                            facturasCompletas[i].Observacion = "";
                        table.AddCell(facturasCompletas[i].Facturas.ToString());
                        table.AddCell(facturasCompletas[i].Ciudad.ToString());
                        table.AddCell(facturasCompletas[i].Nombre.ToString());
                        table.AddCell(facturasCompletas[i].ListaFacturas.Count().ToString());
                        table.AddCell(facturasCompletas[i].Escaneados.Count().ToString());
                        table.AddCell(facturasCompletas[i].Observacion.ToString());
                        bandera = true;
                    }
                }
                if (bandera == false)
                {
                    doc.Add(new Paragraph("Sin observaciones"));
                }
                else
                {
                    table.SetWidths(columnWidths);
                    doc.Add(table);

                }
                doc.Close();
                Process.Start(new ProcessStartInfo(file) { UseShellExecute = true });
                Fact.RemoveAll(f => facturasCompletas.Any(fc => fc.Facturas == f.Facturas));
                Grid.Rows.Clear();
                Grid.Columns.Clear();
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
                int j = 0, k = 0, l = -1, alt = 0;
                if (Fact.Count > 100)
                    alt = 40;
                else
                    alt = 50;
                GlobalSettings.Instance.FacturasTotales = Fact.Count;
                TotalF.Text = GlobalSettings.Instance.FacturasTotales.ToString();
                for (int i = 0; i < Fact.Count; i++)
                {
                    Grid.Rows[k].Height = alt;
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
                Total.Text = GlobalSettings.Instance.BultosTotales.ToString();
                Grid.CurrentCell = null;
                GlobalSettings.Instance.Cargado = true;
                if (facturasCompletasIsi.Count == 0)
                {
                    workbook.SaveAs(GlobalSettings.Instance.filePath);
                    Process.Start(new ProcessStartInfo(GlobalSettings.Instance.filePath) { UseShellExecute = true });
                    return;
                }
                var worksheet2 = workbook.Worksheets.Add("Reporte Embarques Isi");
                string imagePath2 = "C:\\Img\\isi2.png";
                var picture2 = worksheet2.AddPicture(imagePath2).MoveTo(worksheet2.Cell(1, 5)).Scale(0.20);
                worksheet2.Row(1).Height = 100;
                worksheet2.Cell(4, 1).Value = "Factura";
                worksheet2.Cell(2, 2).Value = "Efectivo:";
                worksheet2.Cell(3, 2).Value = "F. Devueltas: ";
                worksheet2.Cell(4, 2).Value = "Ciudad";
                worksheet2.Cell(4, 3).Value = "Bultos";
                worksheet2.Cell(4, 4).Value = "Clave";
                worksheet2.Cell(4, 5).Value = "Cliente";
                worksheet2.Cell(4, 6).Value = "Importe";
                worksheet2.Cell(4, 7).Value = "Condiciones";
                worksheet2.Cell(4, 8).Value = "EF";
                worksheet2.Cell(4, 9).Value = "CR";
                worksheet2.Cell(4, 10).Value = "T/D";
                worksheet2.Cell(4, 11).Value = "CH";
                worksheet2.Cell(4, 12).Value = "Observaciones";
                worksheet2.Cell(4, 13).Value = "Firma Cliente";
                worksheet2.Cell(2, 6).Value = "Operador:";
                worksheet2.Cell(3, 6).Value = "Ruta:";
                worksheet2.Cell(2, 12).Value = "Unidad:";
                worksheet2.Cell(3, 12).Value = "Fecha:";
                worksheet2.Column(1).Width = 7.57;
                worksheet2.Column(2).Width = 22;
                worksheet2.Column(3).Width = 5;
                worksheet2.Column(4).Width = 6.57;
                worksheet2.Column(5).Width = 35;
                worksheet2.Column(6).Width = 11.14;
                worksheet2.Column(7).Width = 12;
                worksheet2.Column(8).Width = 3;
                worksheet2.Column(9).Width = 3;
                worksheet2.Column(10).Width = 3;
                worksheet2.Column(11).Width = 3;
                worksheet2.Column(12).Width = 13;
                worksheet2.Column(13).Width = 25;
                worksheet2.Range(1, 1, 1, 13).Merge();
                worksheet2.Range(2, 1, 3, 1).Merge();
                worksheet2.Range(2, 3, 2, 4).Merge();
                worksheet2.Range(3, 3, 3, 4).Merge();
                worksheet2.Range(2, 5, 3, 5).Merge();
                worksheet2.Range(2, 7, 2, 11).Merge();
                worksheet2.Range(3, 7, 3, 11).Merge();
                var rangoDeCeldas1 = worksheet2.Range(2, 6, 3, 6);
                rangoDeCeldas1.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas1.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                var rangoDeCeldas21 = worksheet2.Range(2, 12, 3, 12);
                rangoDeCeldas21.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas21.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas21.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                var rangoDeCeldas31 = worksheet2.Range(2, 2, 3, 2);
                rangoDeCeldas31.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas31.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas31.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Mover la imagen al centro del rango combinado
                picture2.MoveTo(worksheet2.Cell(1, 5), 200, 40);
                for (int i = 0; i < 13; ++i)
                {
                    var cell3 = worksheet2.Cell(4, i + 1);
                    cell3.Style.Fill.BackgroundColor = XLColor.SteelBlue; // Color de fondo
                    cell3.Style.Font.Bold = true; // Fuente en negrita
                    cell3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }
                int celdas1 = 5;
                // Filtrar las facturas que estén completas (Completo == true)
                for (int i = 0; i < facturasCompletasIsi.Count; ++i)
                {
                    int ckey = facturasCompletasIsi[i].Clave;
                    worksheet2.Row(i + celdas1).Height = 22;
                    var celdafact = worksheet2.Cell(i + celdas1, 1);
                    celdafact.Style.Font.Bold = true; // Fuente en negrita
                    var celdabulto = worksheet2.Cell(i + celdas1, 3);
                    celdabulto.Style.Font.Bold = true; // Fuente en negrita
                    worksheet2.Cell(i + celdas1, 1).Value = facturasCompletasIsi[i].Facturas;
                    worksheet2.Cell(i + celdas1, 2).Value = facturasCompletasIsi[i].Ciudad;
                    worksheet2.Cell(i + celdas1, 3).Value = facturasCompletasIsi[i].Bultos;
                    worksheet2.Cell(i + celdas1, 4).Value = facturasCompletasIsi[i].Clave;
                    worksheet2.Cell(i + celdas1, 5).Value = facturasCompletasIsi[i].Nombre;
                    worksheet2.Cell(i + celdas1, 6).Value = facturasCompletasIsi[i].Importe;
                    worksheet2.Cell(i + celdas1, 7).Value = facturasCompletasIsi[i].Condiciones;
                    worksheet2.Cell(i + celdas1, 8).Value = "EF";
                    worksheet2.Cell(i + celdas1, 9).Value = "CR";
                    worksheet2.Cell(i + celdas1, 10).Value = "T/D";
                    worksheet2.Cell(i + celdas1, 11).Value = "CH";
                    worksheet2.Cell(i + celdas1, 12).Value = "";
                    worksheet2.Cell(i + celdas1, 13).Value = "Recibí " + facturasCompletasIsi[i].Bultos + " Bultos:";
                    if (i != facturasCompletasIsi.Count - 1)
                    {
                        if (ckey == facturasCompletasIsi[i + 1].Clave)
                        {
                            GlobalSettings.Instance.sumaimportes += facturasCompletasIsi[i].Importe;
                            GlobalSettings.Instance.sumabultos += facturasCompletasIsi[i].Bultos;
                        }
                        else
                        {
                            GlobalSettings.Instance.sumaimportes += facturasCompletasIsi[i].Importe;
                            GlobalSettings.Instance.sumabultos += facturasCompletasIsi[i].Bultos;
                            celdas1++;
                            worksheet2.Cell(i + celdas1, 1).Value = "TOTAL DE " + facturasCompletasIsi[i].Nombre;
                            var cell2 = worksheet2.Cell(i + celdas1, 1);
                            var cell3 = worksheet2.Range(i + celdas1, 6, i + celdas1, 7);
                            worksheet2.Range(i + celdas1, 1, i + celdas1, 5).Merge();
                            worksheet2.Range(i + celdas1, 8, i + celdas1, 13).Merge();
                            worksheet2.Cell(i + celdas1, 6).Value = GlobalSettings.Instance.sumaimportes;
                            worksheet2.Cell(i + celdas1, 7).Value = GlobalSettings.Instance.sumabultos + " BULTOS";
                            cell2.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                            cell2.Style.Font.Bold = true; // Fuente en negrita
                            cell2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell3.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                            cell3.Style.Font.Bold = true; // Fuente en negrita
                            cell3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            GlobalSettings.Instance.sumaimportes = 0;
                            GlobalSettings.Instance.sumabultos = 0;
                        }
                    }
                    else
                    {
                        GlobalSettings.Instance.sumaimportes += facturasCompletasIsi[i].Importe;
                        GlobalSettings.Instance.sumabultos += facturasCompletasIsi[i].Bultos;
                        celdas1++;
                        worksheet2.Cell(i + celdas1, 1).Value = "TOTAL DE " + facturasCompletasIsi[i].Nombre;
                        var cell2 = worksheet2.Cell(i + celdas1, 1);
                        var cell3 = worksheet2.Range(i + celdas1, 6, i + celdas1, 7);
                        worksheet2.Range(i + celdas1, 1, i + celdas1, 5).Merge();
                        worksheet2.Range(i + celdas1, 8, i + celdas1, 13).Merge();
                        worksheet2.Cell(i + celdas1, 6).Value = GlobalSettings.Instance.sumaimportes;
                        worksheet2.Cell(i + celdas1, 7).Value = GlobalSettings.Instance.sumabultos + " BULTOS";
                        cell2.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                        cell2.Style.Font.Bold = true; // Fuente en negrita
                        cell2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell3.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                        cell3.Style.Font.Bold = true; // Fuente en negrita
                        cell3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        GlobalSettings.Instance.sumaimportes = 0;
                        GlobalSettings.Instance.sumabultos = 0;
                    }
                    var range = worksheet2.Range("A1:M" + (i + celdas1));
                    range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    fila1 = i + celdas1;
                }
                double total1 = 0;
                double importes1 = 0;
                for (int l2 = 0; l2 < facturasCompletasIsi.Count; ++l2)
                {
                    total1 += facturasCompletasIsi[l2].Bultos;
                    importes1 += facturasCompletasIsi[l2].Importe;

                }
                worksheet2.Range(fila1 + 1, 1, fila1 + 1, 5).Merge();
                worksheet2.Cell(fila1 + 1, 7).Value = total1;
                worksheet2.Cell(fila1 + 1, 6).Value = importes1;
                worksheet2.Cell(fila1 + 1, 1).Value = "TOTALES";
                var rangoDeCeldas41 = worksheet2.Range(fila1 + 1, 1, fila1 + 1, 7);
                rangoDeCeldas41.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas41.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas41.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rangoDeCeldas41.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                rangoDeCeldas41.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                rangoDeCeldas41.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                rangoDeCeldas41.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                // Guardar el archivo
                GlobalSettings.Instance.BultosTotales = 0;
                for (int p =0; p< Fact.Count; p++)
                {
                    GlobalSettings.Instance.BultosTotales += Fact[p].Bultos;
                }
                Total.Text = GlobalSettings.Instance.BultosTotales.ToString();
                Grid.CurrentCell = null;
                GlobalSettings.Instance.Cargado = true;
                workbook.SaveAs(GlobalSettings.Instance.filePath);
                Process.Start(new ProcessStartInfo(GlobalSettings.Instance.filePath) { UseShellExecute = true });

            }
            else
            {
                Fact.AddRange(copia1);
                return;
            }
           

        }

        public void Terminar()
        {
            bool verificar = false;
            var facturasCompletasIsi = Fact.Where(f => f.Facturas.StartsWith("IF")).OrderByDescending(f => f.Orden).ToList();
            var copiaf = facturasCompletasIsi;
            if (facturasCompletasIsi.Count == 0)
            {   
                verificar = true;
            }
            Fact.RemoveAll(f => facturasCompletasIsi.Any(fc => fc.Facturas == f.Facturas));
            var facturasOrdenadas = Fact.OrderByDescending(f => f.Orden).ToList();
            if (facturasOrdenadas.Count == 0 && verificar == true)
            {
                MessageBox.Show("No has escaneado nada", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop); return;
            }
            MessageBox.Show("ASIGNALE UN NOMBRE AL ARCHIVO Y SU UBICACIÓN");
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            saveFileDialog1.InitialDirectory = desktopFolder;
            saveFileDialog1.Title = "ASIGNAR NOMBRE";
            saveFileDialog1.FileName = Ruta.Text + " " + DateTime.Now.ToString("dddd dd-MM-yy");
            saveFileDialog1.Filter = "Archivos de Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*";
            saveFileDialog1.DefaultExt = "xlsx";
            saveFileDialog1.AddExtension = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // Aquí puedes escribir el código para guardar el archivo en la ubicación seleccionada
                GlobalSettings.Instance.filePath = saveFileDialog1.FileName;
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Reporte Embarques");

                // Escribir datos en celdas
                string imagePath = "C:\\Img\\Cornejo.jpg";
                var picture = worksheet.AddPicture(imagePath).MoveTo(worksheet.Cell(1, 1)).Scale(0.20);
                worksheet.Row(1).Height = 100;
                worksheet.Cell(4, 1).Value = "Factura";
                worksheet.Cell(2, 2).Value = "Efectivo:";
                worksheet.Cell(3, 2).Value = "F. Devueltas: ";
                worksheet.Cell(4, 2).Value = "Ciudad";
                worksheet.Cell(4, 3).Value = "Bultos";
                worksheet.Cell(4, 4).Value = "Clave";
                worksheet.Cell(4, 5).Value = "Cliente";
                worksheet.Cell(4, 6).Value = "Importe";
                worksheet.Cell(4, 7).Value = "Condiciones";
                worksheet.Cell(4, 8).Value = "EF";
                worksheet.Cell(4, 9).Value = "CR";
                worksheet.Cell(4, 10).Value = "T/D";
                worksheet.Cell(4, 11).Value = "CH";
                worksheet.Cell(4, 12).Value = "Observaciones";
                worksheet.Cell(4, 13).Value = "Firma Cliente";
                worksheet.Cell(2, 6).Value = "Operador:";
                worksheet.Cell(3, 6).Value = "Ruta:";
                worksheet.Cell(2, 12).Value = "Unidad:";
                worksheet.Cell(3, 12).Value = "Fecha:";
                worksheet.Column(1).Width = 7.57;
                worksheet.Column(2).Width = 22;
                worksheet.Column(3).Width = 5;
                worksheet.Column(4).Width = 6.57;
                worksheet.Column(5).Width = 35;
                worksheet.Column(6).Width = 11.14;
                worksheet.Column(7).Width = 12;
                worksheet.Column(8).Width = 3;
                worksheet.Column(9).Width = 3;
                worksheet.Column(10).Width = 3;
                worksheet.Column(11).Width = 3;
                worksheet.Column(12).Width = 13;
                worksheet.Column(13).Width = 25;
                worksheet.Range(1, 1, 1, 13).Merge();
                worksheet.Range(2, 1, 3, 1).Merge();
                worksheet.Range(2, 3, 2, 4).Merge();
                worksheet.Range(3, 3, 3, 4).Merge();
                worksheet.Range(2, 5, 3, 5).Merge();
                worksheet.Range(2, 7, 2, 11).Merge();
                worksheet.Range(3, 7, 3, 11).Merge();
                var rangoDeCeldas = worksheet.Range(2, 6, 3, 6);
                rangoDeCeldas.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                var rangoDeCeldas2 = worksheet.Range(2, 12, 3, 12);
                rangoDeCeldas2.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas2.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                var rangoDeCeldas3 = worksheet.Range(2, 2, 3, 2);
                rangoDeCeldas3.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas3.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Mover la imagen al centro del rango combinado
                picture.MoveTo(worksheet.Cell(1, 2), 200, 40);
                for (int i = 0; i < 13; ++i)
                {
                    var cell3 = worksheet.Cell(4, i + 1);
                    cell3.Style.Fill.BackgroundColor = XLColor.SteelBlue; // Color de fondo
                    cell3.Style.Font.Bold = true; // Fuente en negrita
                    cell3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }
                int fila = 0;
                int fila1 = 0;
                int celdas = 5;
                for (int i = 0; i < Fact.Count; ++i)
                {
                    int ckey = facturasOrdenadas[i].Clave;
                    worksheet.Row(i + celdas).Height = 22;
                    var celdafact = worksheet.Cell(i + celdas, 1);
                    celdafact.Style.Font.Bold = true; // Fuente en negrita
                    var celdabulto = worksheet.Cell(i + celdas, 3);
                    celdabulto.Style.Font.Bold = true; // Fuente en negrita
                    worksheet.Cell(i + celdas, 1).Value = facturasOrdenadas[i].Facturas;
                    worksheet.Cell(i + celdas, 2).Value = facturasOrdenadas[i].Ciudad;
                    worksheet.Cell(i + celdas, 3).Value = facturasOrdenadas[i].Bultos;
                    worksheet.Cell(i + celdas, 4).Value = facturasOrdenadas[i].Clave;
                    worksheet.Cell(i + celdas, 5).Value = facturasOrdenadas[i].Nombre;
                    worksheet.Cell(i + celdas, 6).Value = facturasOrdenadas[i].Importe;
                    worksheet.Cell(i + celdas, 7).Value = facturasOrdenadas[i].Condiciones;
                    worksheet.Cell(i + celdas, 8).Value = "EF";
                    worksheet.Cell(i + celdas, 9).Value = "CR";
                    worksheet.Cell(i + celdas, 10).Value = "T/D";
                    worksheet.Cell(i + celdas, 11).Value = "CH";
                    worksheet.Cell(i + celdas, 12).Value = "";
                    worksheet.Cell(i + celdas, 13).Value = "Recibí " + facturasOrdenadas[i].Bultos + " Bultos:";
                    if (i != Fact.Count - 1)
                    {
                        if (ckey == facturasOrdenadas[i + 1].Clave)
                        {
                            GlobalSettings.Instance.sumaimportes += facturasOrdenadas[i].Importe;
                            GlobalSettings.Instance.sumabultos += facturasOrdenadas[i].Bultos;
                        }
                        else
                        {
                            GlobalSettings.Instance.sumaimportes += facturasOrdenadas[i].Importe;
                            GlobalSettings.Instance.sumabultos += facturasOrdenadas[i].Bultos;
                            celdas++;
                            worksheet.Cell(i + celdas, 1).Value = "TOTAL DE " + facturasOrdenadas[i].Nombre;
                            var cell2 = worksheet.Cell(i + celdas, 1);
                            var cell3 = worksheet.Range(i + celdas, 6, i + celdas, 7);
                            worksheet.Range(i + celdas, 1, i + celdas, 5).Merge();
                            worksheet.Range(i + celdas, 8, i + celdas, 13).Merge();
                            worksheet.Cell(i + celdas, 6).Value = GlobalSettings.Instance.sumaimportes;
                            worksheet.Cell(i + celdas, 7).Value = GlobalSettings.Instance.sumabultos + " BULTOS";
                            cell2.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                            cell2.Style.Font.Bold = true; // Fuente en negrita
                            cell2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell3.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                            cell3.Style.Font.Bold = true; // Fuente en negrita
                            cell3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            GlobalSettings.Instance.sumaimportes = 0;
                            GlobalSettings.Instance.sumabultos = 0;
                        }
                    }
                    else
                    {
                        GlobalSettings.Instance.sumaimportes += facturasOrdenadas[i].Importe;
                        GlobalSettings.Instance.sumabultos += facturasOrdenadas[i].Bultos;
                        celdas++;
                        worksheet.Cell(i + celdas, 1).Value = "TOTAL DE " + facturasOrdenadas[i].Nombre;
                        var cell2 = worksheet.Cell(i + celdas, 1);
                        var cell3 = worksheet.Range(i + celdas, 6, i + celdas, 7);
                        worksheet.Range(i + celdas, 1, i + celdas, 5).Merge();
                        worksheet.Range(i + celdas, 8, i + celdas, 13).Merge();
                        worksheet.Cell(i + celdas, 6).Value = GlobalSettings.Instance.sumaimportes;
                        worksheet.Cell(i + celdas, 7).Value = GlobalSettings.Instance.sumabultos + " BULTOS";
                        cell2.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                        cell2.Style.Font.Bold = true; // Fuente en negrita
                        cell2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell3.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                        cell3.Style.Font.Bold = true; // Fuente en negrita
                        cell3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        GlobalSettings.Instance.sumaimportes = 0;
                        GlobalSettings.Instance.sumabultos = 0;
                    }
                    var range = worksheet.Range("A1:M" + (i + celdas));
                    range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    fila = i + celdas;

                }
                double total = 0;
                double importes = 0;
                for (int l = 0; l < facturasOrdenadas.Count; ++l)
                {
                    total += facturasOrdenadas[l].Bultos;
                    importes += facturasOrdenadas[l].Importe;

                }
                worksheet.Range(fila + 1, 1, fila + 1, 5).Merge();
                worksheet.Cell(fila + 1, 7).Value = total;
                worksheet.Cell(fila + 1, 6).Value = importes;
                worksheet.Cell(fila + 1, 1).Value = "TOTALES";
                var rangoDeCeldas4 = worksheet.Range(fila + 1, 1, fila + 1, 7);
                rangoDeCeldas4.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas4.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas4.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rangoDeCeldas4.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                rangoDeCeldas4.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                rangoDeCeldas4.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                rangoDeCeldas4.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                // Guardar el archivo
                //ISI
                var worksheet2 = workbook.Worksheets.Add("Reporte Embarques Isi");
                string imagePath2 = "C:\\Img\\isi2.png";
                var picture2 = worksheet2.AddPicture(imagePath2).MoveTo(worksheet2.Cell(1,5)).Scale(0.20);
                worksheet2.Row(1).Height = 100;
                worksheet2.Cell(4, 1).Value = "Factura";
                worksheet2.Cell(2, 2).Value = "Efectivo:";
                worksheet2.Cell(3, 2).Value = "F. Devueltas: ";
                worksheet2.Cell(4, 2).Value = "Ciudad";
                worksheet2.Cell(4, 3).Value = "Bultos";
                worksheet2.Cell(4, 4).Value = "Clave";
                worksheet2.Cell(4, 5).Value = "Cliente";
                worksheet2.Cell(4, 6).Value = "Importe";
                worksheet2.Cell(4, 7).Value = "Condiciones";
                worksheet2.Cell(4, 8).Value = "EF";
                worksheet2.Cell(4, 9).Value = "CR";
                worksheet2.Cell(4, 10).Value = "T/D";
                worksheet2.Cell(4, 11).Value = "CH";
                worksheet2.Cell(4, 12).Value = "Observaciones";
                worksheet2.Cell(4, 13).Value = "Firma Cliente";
                worksheet2.Cell(2, 6).Value = "Operador:";
                worksheet2.Cell(3, 6).Value = "Ruta:";
                worksheet2.Cell(2, 12).Value = "Unidad:";
                worksheet2.Cell(3, 12).Value = "Fecha:";
                worksheet2.Column(1).Width = 7.57;
                worksheet2.Column(2).Width = 22;
                worksheet2.Column(3).Width = 5;
                worksheet2.Column(4).Width = 6.57;
                worksheet2.Column(5).Width = 35;
                worksheet2.Column(6).Width = 11.14;
                worksheet2.Column(7).Width = 12;
                worksheet2.Column(8).Width = 3;
                worksheet2.Column(9).Width = 3;
                worksheet2.Column(10).Width = 3;
                worksheet2.Column(11).Width = 3;
                worksheet2.Column(12).Width = 13;
                worksheet2.Column(13).Width = 25;
                worksheet2.Range(1, 1, 1, 13).Merge();
                worksheet2.Range(2, 1, 3, 1).Merge();
                worksheet2.Range(2, 3, 2, 4).Merge();
                worksheet2.Range(3, 3, 3, 4).Merge();
                worksheet2.Range(2, 5, 3, 5).Merge();
                worksheet2.Range(2, 7, 2, 11).Merge();
                worksheet2.Range(3, 7, 3, 11).Merge();
                var rangoDeCeldas1 = worksheet2.Range(2, 6, 3, 6);
                rangoDeCeldas1.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas1.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                var rangoDeCeldas21 = worksheet2.Range(2, 12, 3, 12);
                rangoDeCeldas21.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas21.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas21.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                var rangoDeCeldas31 = worksheet2.Range(2, 2, 3, 2);
                rangoDeCeldas31.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas31.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas31.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Mover la imagen al centro del rango combinado
                picture2.MoveTo(worksheet2.Cell(1, 5), 200, 40);
                for (int i = 0; i < 13; ++i)
                {
                    var cell3 = worksheet2.Cell(4, i + 1);
                    cell3.Style.Fill.BackgroundColor = XLColor.SteelBlue; // Color de fondo
                    cell3.Style.Font.Bold = true; // Fuente en negrita
                    cell3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }
                int celdas1 = 5;
                // Filtrar las facturas que estén completas (Completo == true)
                for (int i = 0; i < facturasCompletasIsi.Count; ++i)
                {
                    int ckey = facturasCompletasIsi[i].Clave;
                    worksheet2.Row(i + celdas1).Height = 22;
                    var celdafact = worksheet2.Cell(i + celdas1, 1);
                    celdafact.Style.Font.Bold = true; // Fuente en negrita
                    var celdabulto = worksheet2.Cell(i + celdas1, 3);
                    celdabulto.Style.Font.Bold = true; // Fuente en negrita
                    worksheet2.Cell(i + celdas1, 1).Value = facturasCompletasIsi[i].Facturas;
                    worksheet2.Cell(i + celdas1, 2).Value = facturasCompletasIsi[i].Ciudad;
                    worksheet2.Cell(i + celdas1, 3).Value = facturasCompletasIsi[i].Bultos;
                    worksheet2.Cell(i + celdas1, 4).Value = facturasCompletasIsi[i].Clave;
                    worksheet2.Cell(i + celdas1, 5).Value = facturasCompletasIsi[i].Nombre;
                    worksheet2.Cell(i + celdas1, 6).Value = facturasCompletasIsi[i].Importe;
                    worksheet2.Cell(i + celdas1, 7).Value = facturasCompletasIsi[i].Condiciones;
                    worksheet2.Cell(i + celdas1, 8).Value = "EF";
                    worksheet2.Cell(i + celdas1, 9).Value = "CR";
                    worksheet2.Cell(i + celdas1, 10).Value = "T/D";
                    worksheet2.Cell(i + celdas1, 11).Value = "CH";
                    worksheet2.Cell(i + celdas1, 12).Value = "";
                    worksheet2.Cell(i + celdas1, 13).Value = "Recibí " + facturasCompletasIsi[i].Bultos + " Bultos:";
                    if (i != facturasCompletasIsi.Count - 1)
                    {
                        if (ckey == facturasCompletasIsi[i + 1].Clave)
                        {
                            GlobalSettings.Instance.sumaimportes += facturasCompletasIsi[i].Importe;
                            GlobalSettings.Instance.sumabultos += facturasCompletasIsi[i].Bultos;
                        }
                        else
                        {
                            GlobalSettings.Instance.sumaimportes += facturasCompletasIsi[i].Importe;
                            GlobalSettings.Instance.sumabultos += facturasCompletasIsi[i].Bultos;
                            celdas1++;
                            worksheet2.Cell(i + celdas1, 1).Value = "TOTAL DE " + facturasCompletasIsi[i].Nombre;
                            var cell2 = worksheet2.Cell(i + celdas1, 1);
                            var cell3 = worksheet2.Range(i + celdas1, 6, i + celdas1, 7);
                            worksheet2.Range(i + celdas1, 1, i + celdas1, 5).Merge();
                            worksheet2.Range(i + celdas1, 8, i + celdas1, 13).Merge();
                            worksheet2.Cell(i + celdas1, 6).Value = GlobalSettings.Instance.sumaimportes;
                            worksheet2.Cell(i + celdas1, 7).Value = GlobalSettings.Instance.sumabultos + " BULTOS";
                            cell2.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                            cell2.Style.Font.Bold = true; // Fuente en negrita
                            cell2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell3.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                            cell3.Style.Font.Bold = true; // Fuente en negrita
                            cell3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            GlobalSettings.Instance.sumaimportes = 0;
                            GlobalSettings.Instance.sumabultos = 0;
                        }
                    }
                    else
                    {
                        GlobalSettings.Instance.sumaimportes += facturasCompletasIsi[i].Importe;
                        GlobalSettings.Instance.sumabultos += facturasCompletasIsi[i].Bultos;
                        celdas1++;
                        worksheet2.Cell(i + celdas1, 1).Value = "TOTAL DE " + facturasCompletasIsi[i].Nombre;
                        var cell2 = worksheet2.Cell(i + celdas1, 1);
                        var cell3 = worksheet2.Range(i + celdas1, 6, i + celdas1, 7);
                        worksheet2.Range(i + celdas1, 1, i + celdas1, 5).Merge();
                        worksheet2.Range(i + celdas1, 8, i + celdas1, 13).Merge();
                        worksheet2.Cell(i + celdas1, 6).Value = GlobalSettings.Instance.sumaimportes;
                        worksheet2.Cell(i + celdas1, 7).Value = GlobalSettings.Instance.sumabultos + " BULTOS";
                        cell2.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                        cell2.Style.Font.Bold = true; // Fuente en negrita
                        cell2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell3.Style.Fill.BackgroundColor = XLColor.LightBlue; // Color de fondo
                        cell3.Style.Font.Bold = true; // Fuente en negrita
                        cell3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        GlobalSettings.Instance.sumaimportes = 0;
                        GlobalSettings.Instance.sumabultos = 0;
                    }
                    var range = worksheet2.Range("A1:M" + (i + celdas1));
                    range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    fila1 = i + celdas1;
                }
                double total1 = 0;
                double importes1 = 0;
                for (int l2 = 0; l2 < facturasCompletasIsi.Count; ++l2)
                {
                    total1 += facturasCompletasIsi[l2].Bultos;
                    importes1 += facturasCompletasIsi[l2].Importe;

                }
                worksheet2.Range(fila1 + 1, 1, fila1 + 1, 5).Merge();
                worksheet2.Cell(fila1 + 1, 7).Value = total1;
                worksheet2.Cell(fila1 + 1, 6).Value = importes1;
                worksheet2.Cell(fila1 + 1, 1).Value = "TOTALES";
                var rangoDeCeldas41 = worksheet2.Range(fila1 + 1, 1, fila1 + 1, 7);
                rangoDeCeldas41.Style.Fill.BackgroundColor = XLColor.SteelBlue;
                rangoDeCeldas41.Style.Font.Bold = true; // Fuente en negrita
                rangoDeCeldas41.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rangoDeCeldas41.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                rangoDeCeldas41.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                rangoDeCeldas41.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                rangoDeCeldas41.Style.Border.RightBorder = XLBorderStyleValues.Thin;

                workbook.SaveAs(GlobalSettings.Instance.filePath);
                Process.Start(new ProcessStartInfo(GlobalSettings.Instance.filePath) { UseShellExecute = true });
                iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.LETTER.Rotate());
                string file = "C:\\DatosRutas\\" + Ruta.Text + " " + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".pdf";

                PdfWriter.GetInstance(doc, new FileStream(file, FileMode.Create));
                doc.Open();
                float[] columnWidths = new float[] { 14f, 25f, 40f, 20f, 15f, 35f }; // Asumiendo que la segunda columna tendrá un ancho personalizado
                                                                                     //table.SetWidths(columnWidths);
                                                                                     // Crear una tabla para los datos faltantes
                PdfPTable table = new PdfPTable(6);
                PdfPCell cel = new PdfPCell(new Phrase("Reporte de observaciones y faltantes " + Ruta.Text, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 16f, iTextSharp.text.Font.BOLD)));
                cel.Colspan = 6;
                cel.HorizontalAlignment = 1;
                table.AddCell(cel);
                table.AddCell("Folio");
                table.AddCell("Ciudad");
                table.AddCell("Cliente");
                table.AddCell("Bultos Escaneados");
                table.AddCell("Bultos Recibidos");
                table.AddCell("Observación");
                bool bandera = false;
                for (int i = 0; i < Fact.Count; ++i)
                {
                    if (Fact[i].Observacion != null || (Fact[i].Escaneados.Count() != Fact[i].ListaFacturas.Count()))
                    {
                        if (Fact[i].Observacion == null)
                            Fact[i].Observacion = "";
                        table.AddCell(Fact[i].Facturas.ToString());
                        table.AddCell(Fact[i].Ciudad.ToString());
                        table.AddCell(Fact[i].Nombre.ToString());
                        table.AddCell(Fact[i].ListaFacturas.Count().ToString());
                        table.AddCell(Fact[i].Escaneados.Count().ToString());
                        table.AddCell(Fact[i].Observacion.ToString());
                        bandera = true;
                    }
                }
                if (bandera == false)
                {
                    doc.Add(new Paragraph("Sin observaciones"));
                }
                else
                {
                    table.SetWidths(columnWidths);
                    doc.Add(table);

                }
                doc.Close();
                Process.Start(new ProcessStartInfo(file) { UseShellExecute = true });
              //  this.Close();
            }
            else
            {
                Fact.AddRange(copiaf);
                return;
            }

        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.G)
            {
                Terminar();
            }

            if (panel1.Focused || panel1.ClientRectangle.Contains(panel1.PointToClient(Cursor.Position)))
            {
                // Comprobar si la tecla Ctrl + G fue presionada
                if (e.Control && e.KeyCode == Keys.G)
                {
                    Terminar(); // Llamada al método cuando se presionan Ctrl + G
                }

            }
        }

        private void CorteBtn_Click(object sender, EventArgs e)
        {
            Corte();
        }
    }
}