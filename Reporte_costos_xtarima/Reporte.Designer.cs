namespace Reporte_costos_xtarima
{
    partial class Reporte
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Reporte));
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.producto = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.prod_clave = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ucajas = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.utarimas = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ccajas = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ctarimas = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cpucajas = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cputari = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cputarim = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Moneda = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnExcel = new System.Windows.Forms.PictureBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btngenera = new System.Windows.Forms.PictureBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.DGV2 = new System.Windows.Forms.DataGridView();
            this.fcn_folio = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fcn_fecha = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.prov_clave = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.prov_nombre = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fcn_transporte = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fcn_monto_transporte = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pdn_folio = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.um_clave = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LblSem = new System.Windows.Forms.Label();
            this.LblTC = new System.Windows.Forms.Label();
            this.lblUltimoCalculo = new System.Windows.Forms.Label();
            this.DgDet = new System.Windows.Forms.DataGridView();
            this.BtnDet = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnExcel)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btngenera)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DgDet)).BeginInit();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(12, 49);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 20);
            this.label2.TabIndex = 205;
            this.label2.Text = "Del:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(261, 49);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(30, 20);
            this.label1.TabIndex = 206;
            this.label1.Text = "Al:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(378, 8);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(398, 20);
            this.label3.TabIndex = 207;
            this.label3.Text = "Costos promedio de cajas y tarimas por producto";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(477, 32);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(96, 20);
            this.label4.TabIndex = 208;
            this.label4.Text = "Costo flete";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(686, 49);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(145, 20);
            this.label5.TabIndex = 209;
            this.label5.Text = "Moneda Nacional";
            this.label5.Visible = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.producto,
            this.prod_clave,
            this.ucajas,
            this.utarimas,
            this.ccajas,
            this.ctarimas,
            this.cpucajas,
            this.cputari,
            this.cputarim,
            this.Moneda});
            this.dataGridView1.Location = new System.Drawing.Point(12, 92);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(1130, 417);
            this.dataGridView1.TabIndex = 210;
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            this.dataGridView1.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dataGridView1_CellPainting);
            // 
            // producto
            // 
            this.producto.HeaderText = "Producto";
            this.producto.Name = "producto";
            this.producto.ReadOnly = true;
            // 
            // prod_clave
            // 
            this.prod_clave.HeaderText = "prod_clave";
            this.prod_clave.Name = "prod_clave";
            this.prod_clave.ReadOnly = true;
            // 
            // ucajas
            // 
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.ucajas.DefaultCellStyle = dataGridViewCellStyle1;
            this.ucajas.HeaderText = "Cajas";
            this.ucajas.Name = "ucajas";
            this.ucajas.ReadOnly = true;
            // 
            // utarimas
            // 
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.utarimas.DefaultCellStyle = dataGridViewCellStyle2;
            this.utarimas.HeaderText = "Tarimas";
            this.utarimas.Name = "utarimas";
            this.utarimas.ReadOnly = true;
            // 
            // ccajas
            // 
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.ccajas.DefaultCellStyle = dataGridViewCellStyle3;
            this.ccajas.HeaderText = "Cajas";
            this.ccajas.Name = "ccajas";
            this.ccajas.ReadOnly = true;
            // 
            // ctarimas
            // 
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.ctarimas.DefaultCellStyle = dataGridViewCellStyle4;
            this.ctarimas.HeaderText = "tarimas";
            this.ctarimas.Name = "ctarimas";
            this.ctarimas.ReadOnly = true;
            // 
            // cpucajas
            // 
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.cpucajas.DefaultCellStyle = dataGridViewCellStyle5;
            this.cpucajas.HeaderText = "Cajas";
            this.cpucajas.Name = "cpucajas";
            this.cpucajas.ReadOnly = true;
            // 
            // cputari
            // 
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.cputari.DefaultCellStyle = dataGridViewCellStyle6;
            this.cputari.HeaderText = "Tarimas";
            this.cputari.Name = "cputari";
            this.cputari.ReadOnly = true;
            // 
            // cputarim
            // 
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.cputarim.DefaultCellStyle = dataGridViewCellStyle7;
            this.cputarim.HeaderText = "Tarimas";
            this.cputarim.Name = "cputarim";
            this.cputarim.ReadOnly = true;
            // 
            // Moneda
            // 
            this.Moneda.HeaderText = "Moneda";
            this.Moneda.Name = "Moneda";
            this.Moneda.ReadOnly = true;
            // 
            // btnExcel
            // 
            this.btnExcel.Image = global::Reporte_costos_xtarima.Properties.Resources.ExcelMini;
            this.btnExcel.Location = new System.Drawing.Point(1092, 8);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(50, 44);
            this.btnExcel.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.btnExcel.TabIndex = 216;
            this.btnExcel.TabStop = false;
            this.toolTip1.SetToolTip(this.btnExcel, "Exportar reporte a excel");
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // btngenera
            // 
            this.btngenera.Image = global::Reporte_costos_xtarima.Properties.Resources.Guardar;
            this.btngenera.Location = new System.Drawing.Point(1036, 8);
            this.btngenera.Name = "btngenera";
            this.btngenera.Size = new System.Drawing.Size(50, 44);
            this.btngenera.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.btngenera.TabIndex = 293;
            this.btngenera.TabStop = false;
            this.toolTip1.SetToolTip(this.btngenera, "Registra");
            this.btngenera.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar1.Location = new System.Drawing.Point(0, 524);
            this.progressBar1.Maximum = 200;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(1154, 19);
            this.progressBar1.Step = 1;
            this.progressBar1.TabIndex = 291;
            this.progressBar1.Visible = false;
            // 
            // DGV2
            // 
            this.DGV2.AllowUserToAddRows = false;
            this.DGV2.AllowUserToDeleteRows = false;
            this.DGV2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.fcn_folio,
            this.fcn_fecha,
            this.prov_clave,
            this.prov_nombre,
            this.fcn_transporte,
            this.fcn_monto_transporte,
            this.pdn_folio,
            this.um_clave});
            this.DGV2.Location = new System.Drawing.Point(16, 135);
            this.DGV2.Name = "DGV2";
            this.DGV2.ReadOnly = true;
            this.DGV2.RowHeadersVisible = false;
            this.DGV2.Size = new System.Drawing.Size(1130, 417);
            this.DGV2.TabIndex = 294;
            this.DGV2.Visible = false;
            // 
            // fcn_folio
            // 
            this.fcn_folio.HeaderText = "Factura";
            this.fcn_folio.Name = "fcn_folio";
            this.fcn_folio.ReadOnly = true;
            // 
            // fcn_fecha
            // 
            this.fcn_fecha.HeaderText = "Fecha";
            this.fcn_fecha.Name = "fcn_fecha";
            this.fcn_fecha.ReadOnly = true;
            // 
            // prov_clave
            // 
            this.prov_clave.HeaderText = "Clave";
            this.prov_clave.Name = "prov_clave";
            this.prov_clave.ReadOnly = true;
            // 
            // prov_nombre
            // 
            this.prov_nombre.HeaderText = "Nombre";
            this.prov_nombre.Name = "prov_nombre";
            this.prov_nombre.ReadOnly = true;
            // 
            // fcn_transporte
            // 
            this.fcn_transporte.HeaderText = "Talón de embarque";
            this.fcn_transporte.Name = "fcn_transporte";
            this.fcn_transporte.ReadOnly = true;
            // 
            // fcn_monto_transporte
            // 
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.fcn_monto_transporte.DefaultCellStyle = dataGridViewCellStyle8;
            this.fcn_monto_transporte.HeaderText = "Importe flete";
            this.fcn_monto_transporte.Name = "fcn_monto_transporte";
            this.fcn_monto_transporte.ReadOnly = true;
            // 
            // pdn_folio
            // 
            this.pdn_folio.HeaderText = "Orden de venta";
            this.pdn_folio.Name = "pdn_folio";
            this.pdn_folio.ReadOnly = true;
            // 
            // um_clave
            // 
            this.um_clave.HeaderText = "Moneda";
            this.um_clave.Name = "um_clave";
            this.um_clave.ReadOnly = true;
            // 
            // LblSem
            // 
            this.LblSem.AutoSize = true;
            this.LblSem.BackColor = System.Drawing.Color.Transparent;
            this.LblSem.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblSem.ForeColor = System.Drawing.Color.White;
            this.LblSem.Location = new System.Drawing.Point(12, 6);
            this.LblSem.Name = "LblSem";
            this.LblSem.Size = new System.Drawing.Size(71, 20);
            this.LblSem.TabIndex = 295;
            this.LblSem.Text = "Semana";
            // 
            // LblTC
            // 
            this.LblTC.AutoSize = true;
            this.LblTC.BackColor = System.Drawing.Color.Transparent;
            this.LblTC.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblTC.ForeColor = System.Drawing.Color.White;
            this.LblTC.Location = new System.Drawing.Point(12, 29);
            this.LblTC.Name = "LblTC";
            this.LblTC.Size = new System.Drawing.Size(114, 20);
            this.LblTC.TabIndex = 296;
            this.LblTC.Text = "Tipo Cambio:";
            // 
            // lblUltimoCalculo
            // 
            this.lblUltimoCalculo.AutoSize = true;
            this.lblUltimoCalculo.BackColor = System.Drawing.Color.Black;
            this.lblUltimoCalculo.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUltimoCalculo.ForeColor = System.Drawing.Color.Chartreuse;
            this.lblUltimoCalculo.Location = new System.Drawing.Point(12, 82);
            this.lblUltimoCalculo.Name = "lblUltimoCalculo";
            this.lblUltimoCalculo.Size = new System.Drawing.Size(19, 20);
            this.lblUltimoCalculo.TabIndex = 297;
            this.lblUltimoCalculo.Text = "-";
            // 
            // DgDet
            // 
            this.DgDet.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgDet.Location = new System.Drawing.Point(902, 82);
            this.DgDet.Name = "DgDet";
            this.DgDet.Size = new System.Drawing.Size(240, 18);
            this.DgDet.TabIndex = 298;
            this.DgDet.Visible = false;
            // 
            // BtnDet
            // 
            this.BtnDet.Location = new System.Drawing.Point(902, 63);
            this.BtnDet.Name = "BtnDet";
            this.BtnDet.Size = new System.Drawing.Size(75, 23);
            this.BtnDet.TabIndex = 299;
            this.BtnDet.Text = "Detalle";
            this.BtnDet.UseVisualStyleBackColor = true;
            this.BtnDet.Visible = false;
            this.BtnDet.Click += new System.EventHandler(this.BtnDet_Click);
            // 
            // Reporte
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(1154, 543);
            this.Controls.Add(this.BtnDet);
            this.Controls.Add(this.DgDet);
            this.Controls.Add(this.lblUltimoCalculo);
            this.Controls.Add(this.LblTC);
            this.Controls.Add(this.LblSem);
            this.Controls.Add(this.DGV2);
            this.Controls.Add(this.btngenera);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Reporte";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Reporte";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Reporte_FormClosing);
            this.Load += new System.EventHandler(this.Reporte_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnExcel)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btngenera)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DgDet)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.PictureBox btnExcel;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.PictureBox btngenera;
        private System.Windows.Forms.DataGridView DGV2;
        private System.Windows.Forms.DataGridViewTextBoxColumn fcn_folio;
        private System.Windows.Forms.DataGridViewTextBoxColumn fcn_fecha;
        private System.Windows.Forms.DataGridViewTextBoxColumn prov_clave;
        private System.Windows.Forms.DataGridViewTextBoxColumn prov_nombre;
        private System.Windows.Forms.DataGridViewTextBoxColumn fcn_transporte;
        private System.Windows.Forms.DataGridViewTextBoxColumn fcn_monto_transporte;
        private System.Windows.Forms.DataGridViewTextBoxColumn pdn_folio;
        private System.Windows.Forms.DataGridViewTextBoxColumn um_clave;
        private System.Windows.Forms.Label LblSem;
        private System.Windows.Forms.Label LblTC;
        private System.Windows.Forms.DataGridViewTextBoxColumn producto;
        private System.Windows.Forms.DataGridViewTextBoxColumn prod_clave;
        private System.Windows.Forms.DataGridViewTextBoxColumn ucajas;
        private System.Windows.Forms.DataGridViewTextBoxColumn utarimas;
        private System.Windows.Forms.DataGridViewTextBoxColumn ccajas;
        private System.Windows.Forms.DataGridViewTextBoxColumn ctarimas;
        private System.Windows.Forms.DataGridViewTextBoxColumn cpucajas;
        private System.Windows.Forms.DataGridViewTextBoxColumn cputari;
        private System.Windows.Forms.DataGridViewTextBoxColumn cputarim;
        private System.Windows.Forms.DataGridViewTextBoxColumn Moneda;
        private System.Windows.Forms.Label lblUltimoCalculo;
        private System.Windows.Forms.DataGridView DgDet;
        private System.Windows.Forms.Button BtnDet;
    }
}