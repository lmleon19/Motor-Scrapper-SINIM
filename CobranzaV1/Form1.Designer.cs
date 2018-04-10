namespace CobranzaV1
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador requerida.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén utilizando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben eliminar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.BtnProcesar = new System.Windows.Forms.Button();
            this.progressText = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxANO = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // BtnProcesar
            // 
            this.BtnProcesar.BackColor = System.Drawing.Color.Gray;
            this.BtnProcesar.Font = new System.Drawing.Font("Cambria", 9.75F);
            this.BtnProcesar.ForeColor = System.Drawing.Color.White;
            this.BtnProcesar.Location = new System.Drawing.Point(49, 217);
            this.BtnProcesar.Name = "BtnProcesar";
            this.BtnProcesar.Size = new System.Drawing.Size(182, 54);
            this.BtnProcesar.TabIndex = 2;
            this.BtnProcesar.Text = "Clasificador Presupuestario";
            this.BtnProcesar.UseVisualStyleBackColor = false;
            this.BtnProcesar.Click += new System.EventHandler(this.button2_Click);
            // 
            // progressText
            // 
            this.progressText.AutoSize = true;
            this.progressText.BackColor = System.Drawing.Color.Transparent;
            this.progressText.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.progressText.Location = new System.Drawing.Point(12, 371);
            this.progressText.Name = "progressText";
            this.progressText.Size = new System.Drawing.Size(0, 15);
            this.progressText.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Cambria", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(131, 48);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(226, 28);
            this.label3.TabIndex = 9;
            this.label3.Text = "Motor de Búsqueda";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.DarkGray;
            this.button1.Font = new System.Drawing.Font("Cambria", 9.75F);
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(268, 217);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(198, 54);
            this.button1.TabIndex = 2;
            this.button1.Text = "Datos Municipales";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(179, 93);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(122, 13);
            this.label1.TabIndex = 11;
            this.label1.Text = "http://datos.sinim.gov.cl";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(78, 308);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(357, 13);
            this.label2.TabIndex = 12;
            this.label2.Text = "*Descarga y transformacion de los datos por municipio y año seleccionado";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // textBoxANO
            // 
            this.textBoxANO.Location = new System.Drawing.Point(234, 154);
            this.textBoxANO.Name = "textBoxANO";
            this.textBoxANO.Size = new System.Drawing.Size(100, 20);
            this.textBoxANO.TabIndex = 13;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(123, 157);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(105, 13);
            this.label4.TabIndex = 14;
            this.label4.Text = "AÑO (2008-2016)";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(196, 353);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(76, 13);
            this.label5.TabIndex = 15;
            this.label5.Text = "CONTADOR";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(521, 396);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBoxANO);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.progressText);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.BtnProcesar);
            this.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.ShowIcon = false;
            this.Text = "Motor de Búsqueda SINIM";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button BtnProcesar;
        private System.Windows.Forms.Label progressText;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxANO;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
    }
}

