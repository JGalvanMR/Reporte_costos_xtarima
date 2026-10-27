namespace Reporte_costos_xtarima
{
    partial class Producto_Factura
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Producto_Factura));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.fac = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fec = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ped = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cnte_clave = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.prov_clave = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.car_porte = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mon = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.elabo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tip_fle = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mon_fle = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pago_gab = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.iva_fle = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.num_cajas = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lbtpro = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.fac,
            this.fec,
            this.ped,
            this.cnte_clave,
            this.prov_clave,
            this.car_porte,
            this.mon,
            this.elabo,
            this.tip_fle,
            this.mon_fle,
            this.pago_gab,
            this.iva_fle,
            this.num_cajas});
            this.dataGridView1.Location = new System.Drawing.Point(12, 47);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.Size = new System.Drawing.Size(1125, 296);
            this.dataGridView1.TabIndex = 0;
            // 
            // fac
            // 
            this.fac.DataPropertyName = "folio";
            this.fac.HeaderText = "Factura";
            this.fac.Name = "fac";
            this.fac.ReadOnly = true;
            // 
            // fec
            // 
            this.fec.DataPropertyName = "fecha";
            this.fec.HeaderText = "Fecha";
            this.fec.Name = "fec";
            this.fec.ReadOnly = true;
            // 
            // ped
            // 
            this.ped.DataPropertyName = "pdn_folio";
            this.ped.HeaderText = "Pedido";
            this.ped.Name = "ped";
            this.ped.ReadOnly = true;
            // 
            // cnte_clave
            // 
            this.cnte_clave.DataPropertyName = "cnte_clave";
            this.cnte_clave.HeaderText = "Cliente";
            this.cnte_clave.Name = "cnte_clave";
            this.cnte_clave.ReadOnly = true;
            // 
            // prov_clave
            // 
            this.prov_clave.DataPropertyName = "proveedor";
            this.prov_clave.HeaderText = "Transportista";
            this.prov_clave.Name = "prov_clave";
            this.prov_clave.ReadOnly = true;
            // 
            // car_porte
            // 
            this.car_porte.DataPropertyName = "flete";
            this.car_porte.HeaderText = "Carta porte";
            this.car_porte.Name = "car_porte";
            this.car_porte.ReadOnly = true;
            // 
            // mon
            // 
            this.mon.DataPropertyName = "monto1";
            this.mon.HeaderText = "Monto";
            this.mon.Name = "mon";
            this.mon.ReadOnly = true;
            // 
            // elabo
            // 
            this.elabo.DataPropertyName = "fcn_elaboro";
            this.elabo.HeaderText = "Elaboro";
            this.elabo.Name = "elabo";
            this.elabo.ReadOnly = true;
            // 
            // tip_fle
            // 
            this.tip_fle.DataPropertyName = "fcn_tipoflete";
            this.tip_fle.HeaderText = "Tipo flete";
            this.tip_fle.Name = "tip_fle";
            this.tip_fle.ReadOnly = true;
            // 
            // mon_fle
            // 
            this.mon_fle.DataPropertyName = "fcn_monedaflete";
            this.mon_fle.HeaderText = "Moneda flete";
            this.mon_fle.Name = "mon_fle";
            this.mon_fle.ReadOnly = true;
            // 
            // pago_gab
            // 
            this.pago_gab.DataPropertyName = "fcn_pagofletegab";
            this.pago_gab.HeaderText = "Pago GAB";
            this.pago_gab.Name = "pago_gab";
            this.pago_gab.ReadOnly = true;
            // 
            // iva_fle
            // 
            this.iva_fle.DataPropertyName = "fcn_ivaflete";
            this.iva_fle.HeaderText = "IVA flete";
            this.iva_fle.Name = "iva_fle";
            this.iva_fle.ReadOnly = true;
            // 
            // num_cajas
            // 
            this.num_cajas.DataPropertyName = "unicajas";
            this.num_cajas.HeaderText = "# Cajas";
            this.num_cajas.Name = "num_cajas";
            this.num_cajas.ReadOnly = true;
            // 
            // lbtpro
            // 
            this.lbtpro.AutoSize = true;
            this.lbtpro.BackColor = System.Drawing.Color.Transparent;
            this.lbtpro.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbtpro.ForeColor = System.Drawing.Color.White;
            this.lbtpro.Location = new System.Drawing.Point(12, 9);
            this.lbtpro.Name = "lbtpro";
            this.lbtpro.Size = new System.Drawing.Size(0, 20);
            this.lbtpro.TabIndex = 201;
            // 
            // Producto_Factura
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(1150, 355);
            this.Controls.Add(this.lbtpro);
            this.Controls.Add(this.dataGridView1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Producto_Factura";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Facturas-Productos";
            this.Load += new System.EventHandler(this.Producto_Factura_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label lbtpro;
        private System.Windows.Forms.DataGridViewTextBoxColumn fac;
        private System.Windows.Forms.DataGridViewTextBoxColumn fec;
        private System.Windows.Forms.DataGridViewTextBoxColumn ped;
        private System.Windows.Forms.DataGridViewTextBoxColumn cnte_clave;
        private System.Windows.Forms.DataGridViewTextBoxColumn prov_clave;
        private System.Windows.Forms.DataGridViewTextBoxColumn car_porte;
        private System.Windows.Forms.DataGridViewTextBoxColumn mon;
        private System.Windows.Forms.DataGridViewTextBoxColumn elabo;
        private System.Windows.Forms.DataGridViewTextBoxColumn tip_fle;
        private System.Windows.Forms.DataGridViewTextBoxColumn mon_fle;
        private System.Windows.Forms.DataGridViewTextBoxColumn pago_gab;
        private System.Windows.Forms.DataGridViewTextBoxColumn iva_fle;
        private System.Windows.Forms.DataGridViewTextBoxColumn num_cajas;
    }
}