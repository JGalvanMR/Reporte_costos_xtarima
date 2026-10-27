namespace Reporte_costos_xtarima
{
    partial class Reporte_costos_xtarima
    {
        /// <summary>
        /// Variable del diseñador requerida.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén utilizando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben eliminar; false en caso contrario, false.</param>
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Reporte_costos_xtarima));
            this.DTPFinfecha = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.DTPInifecha = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.lbtil = new System.Windows.Forms.Label();
            this.lbcab = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btnRepo = new System.Windows.Forms.PictureBox();
            this.btnsalir = new System.Windows.Forms.PictureBox();
            this.BtnSemanal = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label4 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnRepo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnsalir)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.BtnSemanal)).BeginInit();
            this.SuspendLayout();
            // 
            // DTPFinfecha
            // 
            this.DTPFinfecha.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DTPFinfecha.Location = new System.Drawing.Point(160, 138);
            this.DTPFinfecha.Name = "DTPFinfecha";
            this.DTPFinfecha.Size = new System.Drawing.Size(109, 24);
            this.DTPFinfecha.TabIndex = 203;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(113, 140);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(30, 20);
            this.label3.TabIndex = 205;
            this.label3.Text = "Al:";
            // 
            // DTPInifecha
            // 
            this.DTPInifecha.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DTPInifecha.Location = new System.Drawing.Point(160, 97);
            this.DTPInifecha.Name = "DTPInifecha";
            this.DTPInifecha.Size = new System.Drawing.Size(109, 24);
            this.DTPInifecha.TabIndex = 202;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(113, 100);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 20);
            this.label2.TabIndex = 204;
            this.label2.Text = "Del:";
            // 
            // lbtil
            // 
            this.lbtil.AutoSize = true;
            this.lbtil.BackColor = System.Drawing.Color.Transparent;
            this.lbtil.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbtil.ForeColor = System.Drawing.Color.White;
            this.lbtil.Location = new System.Drawing.Point(137, 40);
            this.lbtil.Name = "lbtil";
            this.lbtil.Size = new System.Drawing.Size(274, 20);
            this.lbtil.TabIndex = 200;
            this.lbtil.Text = "REPORTE DE COSTOS POR TARIMA";
            // 
            // lbcab
            // 
            this.lbcab.AutoSize = true;
            this.lbcab.BackColor = System.Drawing.Color.Transparent;
            this.lbcab.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbcab.ForeColor = System.Drawing.Color.White;
            this.lbcab.Location = new System.Drawing.Point(127, 14);
            this.lbcab.Name = "lbcab";
            this.lbcab.Size = new System.Drawing.Size(295, 20);
            this.lbcab.TabIndex = 199;
            this.lbcab.Text = "SISTEMA INTEGRAL DE PRODUCCIÓN";
            // 
            // pictureBox2
            // 
            this.pictureBox2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBox2.Image = global::Reporte_costos_xtarima.Properties.Resources.resumen;
            this.pictureBox2.Location = new System.Drawing.Point(404, 88);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(77, 72);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 328;
            this.pictureBox2.TabStop = false;
            this.toolTip1.SetToolTip(this.pictureBox2, "Genera reporte de fletes");
            this.pictureBox2.Click += new System.EventHandler(this.pictureBox2_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Reporte_costos_xtarima.Properties.Resources.minimizar;
            this.pictureBox1.Location = new System.Drawing.Point(449, 14);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(32, 32);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 327;
            this.pictureBox1.TabStop = false;
            this.toolTip1.SetToolTip(this.pictureBox1, "Minimizar");
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // btnRepo
            // 
            this.btnRepo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnRepo.Image = global::Reporte_costos_xtarima.Properties.Resources.txt;
            this.btnRepo.Location = new System.Drawing.Point(307, 88);
            this.btnRepo.Name = "btnRepo";
            this.btnRepo.Size = new System.Drawing.Size(77, 72);
            this.btnRepo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.btnRepo.TabIndex = 215;
            this.btnRepo.TabStop = false;
            this.toolTip1.SetToolTip(this.btnRepo, "Genera reporte");
            this.btnRepo.Click += new System.EventHandler(this.btnRepo_Click);
            // 
            // btnsalir
            // 
            this.btnsalir.Image = global::Reporte_costos_xtarima.Properties.Resources.salir;
            this.btnsalir.Location = new System.Drawing.Point(491, 14);
            this.btnsalir.Name = "btnsalir";
            this.btnsalir.Size = new System.Drawing.Size(32, 32);
            this.btnsalir.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.btnsalir.TabIndex = 201;
            this.btnsalir.TabStop = false;
            this.toolTip1.SetToolTip(this.btnsalir, "Cerrar");
            this.btnsalir.Click += new System.EventHandler(this.btnsalir_Click);
            // 
            // BtnSemanal
            // 
            this.BtnSemanal.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BtnSemanal.Image = global::Reporte_costos_xtarima.Properties.Resources.txt;
            this.BtnSemanal.Location = new System.Drawing.Point(491, 196);
            this.BtnSemanal.Name = "BtnSemanal";
            this.BtnSemanal.Size = new System.Drawing.Size(77, 72);
            this.BtnSemanal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.BtnSemanal.TabIndex = 330;
            this.BtnSemanal.TabStop = false;
            this.toolTip1.SetToolTip(this.BtnSemanal, "Genera reporte Semanal");
            this.BtnSemanal.Click += new System.EventHandler(this.BtnSemanal_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Lucida Sans Unicode", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(0, 235);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(403, 16);
            this.label1.TabIndex = 216;
            this.label1.Text = "* Primero debe generarse el reporte para poder registrar los costos";
            // 
            // progressBar1
            // 
            this.progressBar1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar1.Location = new System.Drawing.Point(0, 274);
            this.progressBar1.Maximum = 200;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(580, 19);
            this.progressBar1.Step = 1;
            this.progressBar1.TabIndex = 326;
            this.progressBar1.Visible = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Font = new System.Drawing.Font("Lucida Sans Unicode", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(186, 252);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(163, 17);
            this.label4.TabIndex = 329;
            this.label4.Text = "Generando reporte . . .";
            this.label4.Visible = false;
            // 
            // Reporte_costos_xtarima
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(580, 293);
            this.ControlBox = false;
            this.Controls.Add(this.BtnSemanal);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnRepo);
            this.Controls.Add(this.DTPFinfecha);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.DTPInifecha);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnsalir);
            this.Controls.Add(this.lbtil);
            this.Controls.Add(this.lbcab);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(425, 110);
            this.Name = "Reporte_costos_xtarima";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "4.19 Reporte de costos por tarima";
            this.Load += new System.EventHandler(this.Reporte_costos_xtarima_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnRepo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnsalir)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.BtnSemanal)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker DTPFinfecha;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker DTPInifecha;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox btnsalir;
        private System.Windows.Forms.Label lbtil;
        private System.Windows.Forms.Label lbcab;
        private System.Windows.Forms.PictureBox btnRepo;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.PictureBox BtnSemanal;
    }
}

