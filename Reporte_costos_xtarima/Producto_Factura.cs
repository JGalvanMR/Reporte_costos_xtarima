using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Reporte_costos_xtarima
{
    public partial class Producto_Factura : Form
    {
        public Producto_Factura()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
        }
        public DataTable tabla = new DataTable();
        public DataTable detalle = new DataTable();
        private void Producto_Factura_Load(object sender, EventArgs e)
        {
            lbtpro.Text = "Producto: " + Convert.ToString(tabla.Rows[0]["producto"].ToString()) + " - " + Convert.ToString(tabla.Rows[0]["nproducto"].ToString());
            foreach (DataRow row in tabla.Rows)
                dataGridView1.Rows.Add(row["folio"].ToString().Trim(), row["fecha"].ToString().Trim(), row["pdn_folio"].ToString().Trim(), row["cnte_clave"].ToString().Trim(), row["proveedor"].ToString().Trim(), row["flete"].ToString().Trim(), row["monto1"].ToString().Trim(), row["fcn_elaboro"].ToString().Trim(), row["fcn_tipoflete"].ToString().Trim(), row["fcn_monedaflete"].ToString().Trim(), row["fcn_pagofletegab"].ToString().Trim(), row["fcn_ivaflete"].ToString().Trim(), row["unicajas"].ToString().Trim());

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }
    }  
}
