namespace Validar_Bultos
{
    partial class Observación
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            TxtObservacion = new TextBox();
            TxtVerificado = new TextBox();
            label1 = new Label();
            BtnAgregar = new Button();
            LbFactura = new Label();
            SuspendLayout();
            // 
            // TxtObservacion
            // 
            TxtObservacion.Location = new Point(21, 40);
            TxtObservacion.Multiline = true;
            TxtObservacion.Name = "TxtObservacion";
            TxtObservacion.Size = new Size(303, 130);
            TxtObservacion.TabIndex = 1;
            // 
            // TxtVerificado
            // 
            TxtVerificado.Location = new Point(129, 190);
            TxtVerificado.Name = "TxtVerificado";
            TxtVerificado.Size = new Size(195, 23);
            TxtVerificado.TabIndex = 2;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point);
            label1.Location = new Point(0, 188);
            label1.Name = "label1";
            label1.Size = new Size(123, 21);
            label1.TabIndex = 3;
            label1.Text = "Verificado por:";
            // 
            // BtnAgregar
            // 
            BtnAgregar.Anchor = AnchorStyles.Top;
            BtnAgregar.BackColor = SystemColors.ActiveCaptionText;
            BtnAgregar.Cursor = Cursors.Hand;
            BtnAgregar.FlatStyle = FlatStyle.Flat;
            BtnAgregar.ForeColor = Color.White;
            BtnAgregar.Location = new Point(105, 239);
            BtnAgregar.Name = "BtnAgregar";
            BtnAgregar.Size = new Size(124, 36);
            BtnAgregar.TabIndex = 5;
            BtnAgregar.Text = "Agregar";
            BtnAgregar.UseVisualStyleBackColor = false;
            BtnAgregar.Click += BtnAgregar_Click;
            // 
            // LbFactura
            // 
            LbFactura.Anchor = AnchorStyles.Top;
            LbFactura.AutoSize = true;
            LbFactura.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point);
            LbFactura.Location = new Point(105, 9);
            LbFactura.Name = "LbFactura";
            LbFactura.Size = new Size(18, 21);
            LbFactura.TabIndex = 6;
            LbFactura.Text = "F";
            // 
            // Observación
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.SlateGray;
            ClientSize = new Size(349, 287);
            Controls.Add(LbFactura);
            Controls.Add(BtnAgregar);
            Controls.Add(label1);
            Controls.Add(TxtVerificado);
            Controls.Add(TxtObservacion);
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            Name = "Observación";
            Text = "Observación";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private TextBox TxtObservacion;
        private TextBox TxtVerificado;
        private Label label1;
        private Button BtnAgregar;
        private Label LbFactura;
    }
}