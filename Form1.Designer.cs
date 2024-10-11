namespace Validar_Bultos
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle2 = new DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            label1 = new Label();
            panel1 = new Panel();
            CorteBtn = new Button();
            TotalF = new Label();
            label5 = new Label();
            Total = new Label();
            label4 = new Label();
            FACTURA = new Label();
            Bultos = new Label();
            label3 = new Label();
            Contador = new Label();
            label2 = new Label();
            BtnAgregar = new Button();
            Txt1 = new Label();
            TxtEscanear = new TextBox();
            Ruta = new Label();
            BTN_CARGAR = new Button();
            Grid = new DataGridView();
            openFileDialog1 = new OpenFileDialog();
            saveFileDialog1 = new SaveFileDialog();
            panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)Grid).BeginInit();
            SuspendLayout();
            // 
            // label1
            // 
            label1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom;
            label1.AutoSize = true;
            label1.BackColor = Color.Transparent;
            label1.FlatStyle = FlatStyle.Flat;
            label1.Location = new Point(590, 9);
            label1.Name = "label1";
            label1.Size = new Size(186, 23);
            label1.TabIndex = 0;
            label1.Text = "Validar Facturas";
            label1.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // panel1
            // 
            panel1.BackColor = Color.SteelBlue;
            panel1.Controls.Add(CorteBtn);
            panel1.Controls.Add(TotalF);
            panel1.Controls.Add(label5);
            panel1.Controls.Add(Total);
            panel1.Controls.Add(label4);
            panel1.Controls.Add(FACTURA);
            panel1.Controls.Add(Bultos);
            panel1.Controls.Add(label3);
            panel1.Controls.Add(Contador);
            panel1.Controls.Add(label2);
            panel1.Controls.Add(BtnAgregar);
            panel1.Controls.Add(Txt1);
            panel1.Controls.Add(TxtEscanear);
            panel1.Controls.Add(Ruta);
            panel1.Controls.Add(BTN_CARGAR);
            panel1.Controls.Add(label1);
            panel1.Dock = DockStyle.Top;
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(1350, 96);
            panel1.TabIndex = 1;
            // 
            // CorteBtn
            // 
            CorteBtn.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            CorteBtn.BackColor = SystemColors.ActiveCaptionText;
            CorteBtn.FlatStyle = FlatStyle.Flat;
            CorteBtn.ForeColor = Color.White;
            CorteBtn.Location = new Point(1261, 1);
            CorteBtn.Name = "CorteBtn";
            CorteBtn.Size = new Size(89, 36);
            CorteBtn.TabIndex = 14;
            CorteBtn.Text = "Corte";
            CorteBtn.UseVisualStyleBackColor = false;
            CorteBtn.Visible = false;
            CorteBtn.Click += CorteBtn_Click;
            // 
            // TotalF
            // 
            TotalF.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            TotalF.AutoSize = true;
            TotalF.BackColor = Color.Transparent;
            TotalF.FlatStyle = FlatStyle.Flat;
            TotalF.Font = new Font("Verdana", 21.75F, FontStyle.Bold, GraphicsUnit.Point);
            TotalF.Location = new Point(1174, 24);
            TotalF.Name = "TotalF";
            TotalF.Size = new Size(37, 35);
            TotalF.TabIndex = 13;
            TotalF.Text = "B";
            TotalF.TextAlign = ContentAlignment.MiddleCenter;
            TotalF.Visible = false;
            // 
            // label5
            // 
            label5.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            label5.AutoSize = true;
            label5.BackColor = Color.Transparent;
            label5.FlatStyle = FlatStyle.Flat;
            label5.Font = new Font("Verdana", 15.75F, FontStyle.Bold, GraphicsUnit.Point);
            label5.Location = new Point(932, 34);
            label5.Name = "label5";
            label5.Size = new Size(245, 25);
            label5.TabIndex = 12;
            label5.Text = "Facturas Restantes:";
            label5.TextAlign = ContentAlignment.MiddleCenter;
            label5.Visible = false;
            // 
            // Total
            // 
            Total.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            Total.AutoSize = true;
            Total.BackColor = Color.Transparent;
            Total.FlatStyle = FlatStyle.Flat;
            Total.Font = new Font("Verdana", 21.75F, FontStyle.Bold, GraphicsUnit.Point);
            Total.Location = new Point(1156, 59);
            Total.Name = "Total";
            Total.Size = new Size(37, 35);
            Total.TabIndex = 11;
            Total.Text = "B";
            Total.TextAlign = ContentAlignment.MiddleCenter;
            Total.Visible = false;
            // 
            // label4
            // 
            label4.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            label4.AutoSize = true;
            label4.BackColor = Color.Transparent;
            label4.FlatStyle = FlatStyle.Flat;
            label4.Font = new Font("Verdana", 15.75F, FontStyle.Bold, GraphicsUnit.Point);
            label4.Location = new Point(932, 65);
            label4.Name = "label4";
            label4.Size = new Size(218, 25);
            label4.TabIndex = 10;
            label4.Text = "Bultos Restantes:";
            label4.TextAlign = ContentAlignment.MiddleCenter;
            label4.Visible = false;
            // 
            // FACTURA
            // 
            FACTURA.Anchor = AnchorStyles.Top | AnchorStyles.Bottom;
            FACTURA.AutoSize = true;
            FACTURA.Font = new Font("Verdana", 24F, FontStyle.Bold, GraphicsUnit.Point);
            FACTURA.Location = new Point(590, 1);
            FACTURA.Name = "FACTURA";
            FACTURA.Size = new Size(0, 38);
            FACTURA.TabIndex = 9;
            FACTURA.Visible = false;
            // 
            // Bultos
            // 
            Bultos.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            Bultos.AutoSize = true;
            Bultos.BackColor = Color.Transparent;
            Bultos.FlatStyle = FlatStyle.Flat;
            Bultos.Font = new Font("Verdana", 21.75F, FontStyle.Bold, GraphicsUnit.Point);
            Bultos.Location = new Point(12, 41);
            Bultos.Name = "Bultos";
            Bultos.Size = new Size(0, 35);
            Bultos.TabIndex = 8;
            Bultos.TextAlign = ContentAlignment.MiddleCenter;
            Bultos.Visible = false;
            // 
            // label3
            // 
            label3.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            label3.AutoSize = true;
            label3.BackColor = Color.Transparent;
            label3.FlatStyle = FlatStyle.Flat;
            label3.Font = new Font("Verdana", 21.75F, FontStyle.Bold, GraphicsUnit.Point);
            label3.Location = new Point(12, -3);
            label3.Name = "label3";
            label3.Size = new Size(118, 35);
            label3.TabIndex = 7;
            label3.Text = "Bultos";
            label3.TextAlign = ContentAlignment.MiddleCenter;
            label3.Visible = false;
            // 
            // Contador
            // 
            Contador.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            Contador.AutoSize = true;
            Contador.BackColor = Color.Transparent;
            Contador.FlatStyle = FlatStyle.Flat;
            Contador.Font = new Font("Verdana", 21.75F, FontStyle.Bold, GraphicsUnit.Point);
            Contador.Location = new Point(251, 41);
            Contador.Name = "Contador";
            Contador.Size = new Size(0, 35);
            Contador.TabIndex = 6;
            Contador.TextAlign = ContentAlignment.MiddleCenter;
            Contador.Visible = false;
            // 
            // label2
            // 
            label2.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            label2.AutoSize = true;
            label2.BackColor = Color.Transparent;
            label2.FlatStyle = FlatStyle.Flat;
            label2.Font = new Font("Verdana", 21.75F, FontStyle.Bold, GraphicsUnit.Point);
            label2.Location = new Point(251, 1);
            label2.Name = "label2";
            label2.Size = new Size(176, 35);
            label2.TabIndex = 5;
            label2.Text = "Restantes";
            label2.TextAlign = ContentAlignment.MiddleCenter;
            label2.Visible = false;
            // 
            // BtnAgregar
            // 
            BtnAgregar.Anchor = AnchorStyles.Top;
            BtnAgregar.BackColor = SystemColors.ActiveCaptionText;
            BtnAgregar.FlatStyle = FlatStyle.Flat;
            BtnAgregar.ForeColor = Color.White;
            BtnAgregar.Location = new Point(802, 54);
            BtnAgregar.Name = "BtnAgregar";
            BtnAgregar.Size = new Size(124, 36);
            BtnAgregar.TabIndex = 4;
            BtnAgregar.Text = "Agregar";
            BtnAgregar.UseVisualStyleBackColor = false;
            BtnAgregar.Visible = false;
            BtnAgregar.Click += BtnAgregar_Click;
            // 
            // Txt1
            // 
            Txt1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom;
            Txt1.AutoSize = true;
            Txt1.BackColor = Color.Transparent;
            Txt1.FlatStyle = FlatStyle.Flat;
            Txt1.Location = new Point(469, 61);
            Txt1.Name = "Txt1";
            Txt1.Size = new Size(115, 23);
            Txt1.TabIndex = 3;
            Txt1.Text = "Escanear:";
            Txt1.TextAlign = ContentAlignment.MiddleCenter;
            Txt1.Visible = false;
            // 
            // TxtEscanear
            // 
            TxtEscanear.Anchor = AnchorStyles.Top | AnchorStyles.Bottom;
            TxtEscanear.BackColor = Color.Ivory;
            TxtEscanear.BorderStyle = BorderStyle.FixedSingle;
            TxtEscanear.CharacterCasing = CharacterCasing.Upper;
            TxtEscanear.ForeColor = Color.Black;
            TxtEscanear.Location = new Point(590, 59);
            TxtEscanear.Name = "TxtEscanear";
            TxtEscanear.Size = new Size(177, 31);
            TxtEscanear.TabIndex = 2;
            TxtEscanear.Visible = false;
            TxtEscanear.KeyDown += TxtEscanear_KeyDown;
            // 
            // Ruta
            // 
            Ruta.Anchor = AnchorStyles.Top | AnchorStyles.Bottom;
            Ruta.AutoSize = true;
            Ruta.Font = new Font("Verdana", 20.25F, FontStyle.Bold, GraphicsUnit.Point);
            Ruta.ForeColor = Color.Red;
            Ruta.Location = new Point(940, -1);
            Ruta.Name = "Ruta";
            Ruta.Size = new Size(106, 32);
            Ruta.TabIndex = 1;
            Ruta.Text = "label2";
            Ruta.Visible = false;
            // 
            // BTN_CARGAR
            // 
            BTN_CARGAR.Anchor = AnchorStyles.Top | AnchorStyles.Bottom;
            BTN_CARGAR.FlatStyle = FlatStyle.Flat;
            BTN_CARGAR.Location = new Point(590, 54);
            BTN_CARGAR.Name = "BTN_CARGAR";
            BTN_CARGAR.Size = new Size(177, 36);
            BTN_CARGAR.TabIndex = 0;
            BTN_CARGAR.Text = "Cargar Ruta";
            BTN_CARGAR.UseVisualStyleBackColor = true;
            BTN_CARGAR.Click += BTN_CARGAR_Click;
            // 
            // Grid
            // 
            Grid.BackgroundColor = Color.SteelBlue;
            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = SystemColors.ButtonShadow;
            dataGridViewCellStyle1.Font = new Font("Verdana", 14.25F, FontStyle.Bold, GraphicsUnit.Point);
            dataGridViewCellStyle1.ForeColor = SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.True;
            Grid.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            Grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            Grid.ColumnHeadersVisible = false;
            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = SystemColors.Window;
            dataGridViewCellStyle2.Font = new Font("Verdana", 14.25F, FontStyle.Bold, GraphicsUnit.Point);
            dataGridViewCellStyle2.ForeColor = SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = SystemColors.AppWorkspace;
            dataGridViewCellStyle2.SelectionForeColor = Color.Black;
            dataGridViewCellStyle2.WrapMode = DataGridViewTriState.False;
            Grid.DefaultCellStyle = dataGridViewCellStyle2;
            Grid.Dock = DockStyle.Fill;
            Grid.EditMode = DataGridViewEditMode.EditProgrammatically;
            Grid.Location = new Point(0, 96);
            Grid.MultiSelect = false;
            Grid.Name = "Grid";
            Grid.RowHeadersVisible = false;
            Grid.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            Grid.RowTemplate.Height = 25;
            Grid.Size = new Size(1350, 633);
            Grid.TabIndex = 2;
            Grid.CellContentClick += Grid_CellContentClick;
            Grid.CellMouseDown += Grid_CellMouseDown;
            Grid.CurrentCellChanged += Grid_CurrentCellChanged;
            Grid.KeyDown += Grid_KeyDown;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(13F, 23F);
            AutoScaleMode = AutoScaleMode.Font;
            BackgroundImage = (Image)resources.GetObject("$this.BackgroundImage");
            BackgroundImageLayout = ImageLayout.None;
            ClientSize = new Size(1350, 729);
            Controls.Add(Grid);
            Controls.Add(panel1);
            Cursor = Cursors.Hand;
            Font = new Font("Verdana", 14.25F, FontStyle.Bold, GraphicsUnit.Point);
            Icon = (Icon)resources.GetObject("$this.Icon");
            ImeMode = ImeMode.Off;
            Margin = new Padding(5);
            Name = "Form1";
            ShowIcon = false;
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Validar Facturas";
            WindowState = FormWindowState.Maximized;
            KeyDown += Form1_KeyDown;
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)Grid).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private Label label1;
        private Panel panel1;
        private Button BTN_CARGAR;
        private OpenFileDialog openFileDialog1;
        private Label Ruta;
        private TextBox TxtEscanear;
        private Label Txt1;
        private Button BtnAgregar;
        private Label label2;
        private Label Bultos;
        private Label label3;
        private Label Contador;
        private Label FACTURA;
        private Label Total;
        private Label label4;
        public DataGridView Grid;
        private Label TotalF;
        private Label label5;
        private SaveFileDialog saveFileDialog1;
        private Button CorteBtn;
    }
}