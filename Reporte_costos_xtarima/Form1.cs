using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Globalization;

namespace Reporte_costos_xtarima
{
    public partial class Reporte_costos_xtarima : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        DataSet des = new DataSet("tm1");
        DataSet ds = new DataSet();
        DataTable tmr1 = new DataTable("tmr1");
        DataTable tipo_cambio = new DataTable("tipo_cambio");
        DataTable flete = new DataTable("flete");
        DataTable productos = new DataTable("productos");
        DataTable detalle_factura = new DataTable("detalle_factura");
        SqlDataAdapter da;
        SqlCommand cmnd1 = new SqlCommand();
        SqlDataReader reader1;

        public static string prov_ini, prov_fin, query, FI, FF, fle, fle2, letrero, pro, pro2, factura, moneda;
        public static double tot = 0, mon = 1, t = 0, porceiva = 0;
        public static double caj = 0, tar = 0, dincaj = 0, dintar = 0, t1 = 0, t2 = 0, t3 = 0, t4 = 0;
        public static DateTime date;
        public static bool find = false;
        public static int reporte = 0;
        public static string tipo_reporte = "";

        public Reporte_costos_xtarima()
        {
            InitializeComponent();

            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
        }

        private void Reporte_costos_xtarima_Load(object sender, EventArgs e)
        {
            Maketmr1();
            var inv_pt = "C:\\Reportes\\Reportecostoxtarima.xls";
            if (!File.Exists(inv_pt))
            {
                var fileStream = File.Create(inv_pt);
            }
            int var_mesActual = DateTime.Now.Month; // obtengo el mes actual
            int var_anio = DateTime.Now.Year; // obtengo el año actual
            int var_mesSiguiente = DateTime.Now.Month + 1; // obtengo el mes siguiente
            int var_aniosiguiente = DateTime.Now.Year;
            DTPInifecha.Value = Convert.ToDateTime("01/" + var_mesActual + "/" + var_anio);// pongo el 1 porque siempre es el primer día obvio
            if (var_mesSiguiente > 12)
            {
                var_mesSiguiente = 1;
                var_aniosiguiente = var_aniosiguiente + 1;
            }
            DTPFinfecha.Value = Convert.ToDateTime("01/" + var_mesSiguiente + "/" + var_aniosiguiente).AddDays(-1); //resto un día al mes y con esto obtengo el ultimo día
            thisConnection.Open();
            query = "select * from tb_cat_producto order by prod_nombre";
            da = new SqlDataAdapter(query, thisConnection);
            da.Fill(ds, "productos");
            productos = ds.Tables["productos"];
            thisConnection.Close();
        }

        //crea datatable para el reporte
        private void Maketmr1()
        {
            tmr1.Columns.Add("flete");
            tmr1.Columns.Add("folio");
            tmr1.Columns.Add("fecha", typeof(DateTime));
            tmr1.Columns.Add("tipo");
            tmr1.Columns.Add("lugar");
            tmr1.Columns.Add("proveedor");
            tmr1.Columns.Add("nombrep");
            tmr1.Columns.Add("producto");
            tmr1.Columns.Add("nproducto");
            tmr1.Columns.Add("CVEL");
            tmr1.Columns.Add("NOMLINEA");
            des.Tables.Add(tmr1);

            DataColumn columna2 = new DataColumn();
            columna2.ColumnName = "unicajas";
            columna2.DataType = System.Type.GetType("System.Double");
            columna2.DefaultValue = 0;
            tmr1.Columns.Add(columna2);

            DataColumn columna3 = new DataColumn();
            columna3.ColumnName = "monto1";
            columna3.DataType = System.Type.GetType("System.Double");
            columna3.DefaultValue = 0;
            tmr1.Columns.Add(columna3);

            DataColumn columna4 = new DataColumn();
            columna4.ColumnName = "monto2";
            columna4.DataType = System.Type.GetType("System.Double");
            columna4.DefaultValue = 0;
            tmr1.Columns.Add(columna4);

            DataColumn columna5 = new DataColumn();
            columna5.ColumnName = "impuesto";
            columna5.DataType = System.Type.GetType("System.Double");
            columna5.DefaultValue = 0;
            tmr1.Columns.Add(columna5);

            DataColumn columna6 = new DataColumn();
            columna6.ColumnName = "unitarim";
            columna6.DataType = System.Type.GetType("System.Double");
            columna6.DefaultValue = 0;
            tmr1.Columns.Add(columna6);

            DataColumn columna7 = new DataColumn();
            columna7.ColumnName = "cajtot";
            columna7.DataType = System.Type.GetType("System.Double");
            columna7.DefaultValue = 0;
            tmr1.Columns.Add(columna7);

            DataColumn columna8 = new DataColumn();
            columna8.ColumnName = "tartot";
            columna8.DataType = System.Type.GetType("System.Double");
            columna8.DefaultValue = 0;
            tmr1.Columns.Add(columna8);

            DataColumn columna9 = new DataColumn();
            columna9.ColumnName = "tartot2";
            columna9.DataType = System.Type.GetType("System.Double");
            columna9.DefaultValue = 0;
            tmr1.Columns.Add(columna9);

            DataColumn columna10 = new DataColumn();
            columna10.ColumnName = "cosxcaja";
            columna10.DataType = System.Type.GetType("System.Double");
            columna10.DefaultValue = 0;
            tmr1.Columns.Add(columna10);

            DataColumn columna11 = new DataColumn();
            columna11.ColumnName = "cosxtari";
            columna11.DataType = System.Type.GetType("System.Double");
            columna11.DefaultValue = 0;
            tmr1.Columns.Add(columna11);

            DataColumn columna12 = new DataColumn();
            columna12.ColumnName = "totcajp";
            columna12.DataType = System.Type.GetType("System.Double");
            columna12.DefaultValue = 0;
            tmr1.Columns.Add(columna12);

            DataColumn columna13 = new DataColumn();
            columna13.ColumnName = "tottarp";
            columna13.DataType = System.Type.GetType("System.Double");
            columna13.DefaultValue = 0;
            tmr1.Columns.Add(columna13);

            DataColumn columna14 = new DataColumn();
            columna14.ColumnName = "coscajp";
            columna14.DataType = System.Type.GetType("System.Double");
            columna14.DefaultValue = 0;
            tmr1.Columns.Add(columna14);

            DataColumn columna15 = new DataColumn();
            columna15.ColumnName = "costarp";
            columna15.DataType = System.Type.GetType("System.Double");
            columna15.DefaultValue = 0;
            tmr1.Columns.Add(columna15);
        }

        private void btnRepo_Click(object sender, EventArgs e)
        {
            reporte = 1;
            date = DTPInifecha.Value;
            TimeSpan dias = DTPFinfecha.Value.Subtract(DTPInifecha.Value);
            //int dias = DTPFinfecha.Value.Day - DTPInifecha.Value.Day;

            if (dias.Days != 31 && dias.Days != 30 && dias.Days != 28 && dias.Days != 29 && dias.Days != 27)
            //if (dias != 31 && dias != 30 && dias != 28 && dias != 29 && dias != 27)
            {
                MessageBox.Show("EL periodo a generarse debe ser de un mes", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                tipo_reporte = "M";
                Genera("M"); // Tipo Mensual
            }
        }

        private void btnsalir_Click(object sender, EventArgs e)
        {
            try
            {
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT TOP 1 inicio_sesion, usu_login FROM tb_cat_historial_dia where nombre_maquina = '" + Environment.MachineName + "' and sistema = 'SIPGAB' ORDER BY inicio_sesion desc";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    Utilerias.Class1.Inicio_sesion = reader1.GetSqlDateTime(0).Value;
                    Utilerias.Class1.Usu_login = reader1.GetSqlString(1).ToString();
                    Utilerias.Class1.Nombre_equipo = Environment.MachineName;
                }
                reader1.Close();

                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "update tb_cat_historial_dia set formulario = ' ' where nombre_maquina ='" + Utilerias.Class1.Nombre_equipo + "' and usu_login = '" + Utilerias.Class1.Usu_login + "' and inicio_sesion = '" + Utilerias.Class1.Inicio_sesion.ToString("s") + "' and sistema = 'SIPGAB'";
                reader1 = cmnd1.ExecuteReader();
                reader1.Close();
                thisConnection.Close();
                Application.Exit();
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString());
                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            catch (Exception ex1)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString());
                MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            label4.Visible = true;
            label4.Update();
            reporte = 2;
            DateTime fecha1 = Convert.ToDateTime(DTPInifecha.Text);
            DateTime fecha2 = Convert.ToDateTime(DTPFinfecha.Text);
            int fecha = DateTime.Compare(fecha2, fecha1);

            if (fecha < 0)
            {
                MessageBox.Show("Rango de fechas invalido", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            FI = DTPInifecha.Text;
            FF = DTPFinfecha.Text;

            query = "select a.fcn_folio, a.fcn_fecha, a.prov_clave, b.prov_nombre, a.fcn_transporte, a.fcn_monto_transporte, a.pdn_folio, " +
                    "a.um_clave from tb_mstr_facturas_nal a, tb_cat_proveedor b where (a.fcn_fecha between '" + DTPInifecha.Text + "' and " +
                    "'" + DTPFinfecha.Text + "') and b.prov_clave = a.prov_clave and a.fcn_transporte <> '' and a.fcn_pagofletegab ='N' and " +
                    "a.fcn_estatus <> 'C' and a.fcn_pedorigen = 0 order by a.fcn_fecha, a.fcn_folio, a.um_clave";
            thisConnection.Open();
            da = new SqlDataAdapter(query, thisConnection);
            da.Fill(des, "flete");
            flete = des.Tables["flete"];
            thisConnection.Close();

            Reporte repo = new Reporte();
            repo.tem = flete;
            label4.Visible = false;
            repo.ShowDialog();
            flete.Rows.Clear();
        }

        private void BtnSemanal_Click(object sender, EventArgs e)
        {
            reporte = 1;
            date = DTPInifecha.Value;
            TimeSpan dias = DTPFinfecha.Value.Subtract(DTPInifecha.Value);
            //int dias = DTPFinfecha.Value.Day - DTPInifecha.Value.Day;

            //if (dias.Days != 31 && dias.Days != 30 && dias.Days != 28 && dias.Days != 29 && dias.Days != 27)
            //if (dias != 31 && dias != 30 && dias != 28 && dias != 29 && dias != 27)
            if (DTPInifecha.Value.DayOfWeek.ToString().ToUpper() != "MONDAY")
            {
                MessageBox.Show("La Fecha Inicial Debe de Empezar en Lunes", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dias.Days != 6)
            {
                MessageBox.Show("EL periodo a generarse debe ser de una Semana Terminando en Domingo", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                thisConnection.Open();
                string Cadena = "SELECT semana FROM tb_cat_semanas " +
                             " Where fecha1 = '" + DTPInifecha.Value.ToShortDateString() + "'order by ano,semana";
                SqlCommand cmd = new SqlCommand(Cadena, thisConnection);
                string NSem = Convert.ToString(cmd.ExecuteScalar()).Trim();
                Cadena = "Select NoSem, Año, Lin_Clave, Prod_Clave FROM Tb_Costos_DetProd Where Año = '" + DTPInifecha.Value.Year.ToString() + "' AND NoSem = '" + NSem.ToString() + "'";
                SqlDataAdapter da = new SqlDataAdapter(Cadena, thisConnection);
                DataSet ds = new DataSet();
                da.Fill(ds, "Costos");
                DataTable Costos = ds.Tables["Costos"];
                //if (Costos.Rows.Count == 0)
                //{
                //    MessageBox.Show("Error: No se ha generado la Informacion de Ventas, Generar las Ventas de Esta Semana " + NSem.ToString(), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                //    thisConnection.Close();
                //    return;
                //}
                thisConnection.Close();
                tipo_reporte = "S";
                Genera("S"); // Tipo Semanal
            }
        }

        private void Genera(string Tipo)
        {
            try
            {
                if (DTPInifecha.Value < DateTime.Parse("01/09/2012"))
                    porceiva = 1.11;
                else
                    porceiva = 1.12;

                thisConnection.Open();
                FI = DTPInifecha.Value.ToShortDateString();
                FF = DTPFinfecha.Value.ToShortDateString();

                query = "select * from tb_cat_tipocambio where mes = '" + DTPInifecha.Value.Month.ToString() + "' and año ='" + DTPInifecha.Value.Year.ToString() + "' order by dia desc";
                da = new SqlDataAdapter(query, thisConnection);
                da.Fill(des, "tipo_cambio");
                tipo_cambio = des.Tables["tipo_cambio"];


                //query = "SELECT tmfn.fcn_fecha as fecha, tmfn.fcn_transporte as flete, Tmfn.fcn_folio as folio, tmfn.fcn_monto_transporte as monto1, " +
                //        "tmfn.prov_clave as proveedor, tmfn.fcn_tipo as tipo, tmfn.fcn_lugar as lugar, tdf.PROD_CLAVE as producto, " +
                //        "tdf.FCN_NUM_UNIDADES as unicajas, tcpr.prod_nombre as nproducto, tcpr.prod_num_tarimas as unitarim, " +
                //        "tcprov.prov_nombre as nombrep, tcpr.lin_clave as CVEL, tcl.lin_nombre as NOMLINEA, tmfn.fcn_monedaflete " +
                //        "FROM tb_mstr_facturas_nal tmfn, tb_det_facturas tdf, tb_cat_producto tcpr, tb_cat_proveedor tcprov, tb_cat_linea tcl " +
                //        "WHERE (tmfn.fcn_fecha BETWEEN '" + FI + "' AND '" + FF + "') " +
                //        "AND (tmfn.fcn_transporte <='zzzzzz' AND tmfn.fcn_transporte >='000000') AND (tmfn.prov_clave BETWEEN '000000' AND 'zzzzzz') " +
                //        "AND tmfn.fcn_estatus <> 'C' AND tmfn.fcn_lugar <> 'EXP' AND tmfn.fcn_transporte <> '' AND tdf.fcn_folio= tmfn.fcn_folio " +
                //        "AND tdf.fcn_tipo = tmfn.fcn_lugar AND tcpr.prod_clave = tdf.prod_clave AND tcprov.prov_clave = tmfn.prov_clave " +
                //        "AND tmfn.fcn_lugar <> 'FC' AND tcl.lin_clave = tdf.lin_clave and tmfn.fcn_lugar <> 'FC' and tmfn.fcn_pagofletegab <> 'N' " +
                //    //"ORDER BY tmfn.fcn_fecha, tmfn.fcn_transporte, tmfn.fcn_folio, tdf.prod_clave";
                //    //"ORDER BY tdf.prod_clave, tcpr.prod_nombre, tcpr.lin_clave, tcl.lin_nombre";
                //        "ORDER BY tmfn.fcn_folio, tmfn.fcn_fecha, tdf.prod_clave";
                //"ORDER BY tmfn.fcn_fecha, tmfn.fcn_folio";

                /*query = "SELECT tmfn.fcn_fecha as fecha, tmfn.fcn_transporte as flete, Tmfn.fcn_folio as folio, tmfn.fcn_monto_transporte as monto1, " +
                        "tmfn.prov_clave as proveedor, tmfn.fcn_tipo as tipo, tmfn.fcn_lugar as lugar, tdf.PROD_CLAVE as producto, " +
                        "tdf.FCN_NUM_UNIDADES as unicajas, tcpr.prod_nombre as nproducto, tcpr.prod_num_tarimas as unitarim, " +
                        "tcprov.prov_nombre as nombrep, tcpr.lin_clave as CVEL, tcl.lin_nombre as NOMLINEA, tmfn.fcn_monedaflete, tmfn.fcn_ivaflete " +
                        "FROM tb_mstr_facturas_nal tmfn, tb_det_facturas tdf, tb_cat_producto tcpr, tb_cat_proveedor tcprov, tb_cat_linea tcl " +
                        "WHERE (tmfn.fcn_fecha BETWEEN '" + FI + "' AND '" + FF + "') " +
                        "AND (tmfn.fcn_transporte <='zzzzzz' AND tmfn.fcn_transporte >='000000') AND (tmfn.prov_clave BETWEEN '000000' AND 'zzzzzz') " +
                        "AND tmfn.fcn_estatus <> 'C' AND tmfn.fcn_transporte <> '' AND tdf.fcn_folio= tmfn.fcn_folio " +
                        "AND tdf.fcn_tipo = tmfn.fcn_lugar AND tcpr.prod_clave = tdf.prod_clave AND tcprov.prov_clave = tmfn.prov_clave " +
                        "AND tmfn.fcn_lugar <> 'FC' AND tcl.lin_clave = tdf.lin_clave and tmfn.fcn_lugar <> 'FC' and tmfn.fcn_pagofletegab <> 'N' " +                        
                        "ORDER BY tmfn.fcn_monedaflete, tmfn.fcn_folio, tmfn.fcn_fecha, tdf.prod_clave";*/

                /*query = "SELECT tmfn.fcn_fecha as fecha, tmfn.fcn_transporte as flete, Tmfn.fcn_folio as folio, tmfn.fcn_monto_transporte as monto1, tmfn.prov_clave as proveedor, " +
                        "tmfn.fcn_tipo as tipo, tmfn.fcn_lugar as lugar, tdf.PROD_CLAVE as producto, tdf.FCN_NUM_UNIDADES as unicajas, tmfn.fcn_monedaflete, tmfn.fcn_ivaflete " +
                        "FROM tb_mstr_facturas_nal tmfn, tb_det_facturas tdf " +
                        "WHERE (tmfn.fcn_fecha BETWEEN '01/12/2015' AND '31/12/2015') AND " +
                        "(tmfn.fcn_transporte <='zzzzzz' AND tmfn.fcn_transporte >='000000') AND (tmfn.prov_clave BETWEEN '000000' AND 'zzzzzz') AND tmfn.fcn_estatus <> 'C' AND " +
                        "tmfn.fcn_transporte <> '' AND tdf.fcn_folio= tmfn.fcn_folio AND tdf.fcn_tipo = tmfn.fcn_lugar " +
                        "and tmfn.fcn_pagofletegab = 'S' " +
                        "ORDER BY tmfn.fcn_monedaflete, tmfn.fcn_folio, tmfn.fcn_fecha, tdf.prod_clave";*/
                /*query = "SELECT tmfn.fcn_fecha as fecha, tmfn.fcn_transporte as flete, Tmfn.fcn_folio as folio, tmfn.fcn_monto_transporte as monto1, tmfn.prov_clave as proveedor, " +
                        "tmfn.fcn_tipo as tipo, tmfn.fcn_lugar as lugar, tdf.PROD_CLAVE as producto, tdf.FCN_NUM_UNIDADES as unicajas, tmfn.fcn_monedaflete, tmfn.fcn_ivaflete, " +
                        "tcpr.prod_nombre as nproducto, tcpr.prod_num_tarimas as unitarim FROM tb_mstr_facturas_nal tmfn, tb_det_facturas tdf, tb_cat_producto tcpr " +
                        "WHERE (tmfn.fcn_fecha BETWEEN '" + FI + "' AND '" + FF + "') AND " +
                        "tmfn.fcn_estatus <> 'C' AND " +
                        "tdf.fcn_folio= tmfn.fcn_folio AND tdf.fcn_tipo = tmfn.fcn_lugar " +
                        "AND tcpr.prod_clave = tdf.prod_clave  and tmfn.fcn_pagofletegab = 'S' ORDER BY tmfn.fcn_monedaflete, tmfn.fcn_folio, tmfn.fcn_fecha, tdf.prod_clave";*/
                /*query = "SELECT tmfn.fcn_fecha as fecha, tmfn.fcn_transporte as flete, Tmfn.fcn_folio as folio, tmfn.fcn_monto_transporte as monto1, tmfn.prov_clave as proveedor, " +
                        "tmfn.fcn_tipo as tipo, tmfn.fcn_lugar as lugar, tdf.PROD_CLAVE as producto, tdf.FCN_NUM_UNIDADES as unicajas, tcpr.prod_nombre as nproducto, tcpr.prod_num_tarimas as unitarim, " +
                        "tmfn.fcn_monedaflete, tmfn.fcn_ivaflete FROM tb_mstr_facturas_nal tmfn, tb_det_facturas tdf, tb_cat_producto tcpr " +
                        "WHERE (tmfn.fcn_fecha BETWEEN '" + FI + "' AND '" + FF + "') AND tmfn.fcn_estatus <> 'C' AND " +
                        "tdf.fcn_folio= tmfn.fcn_folio AND tdf.fcn_tipo = tmfn.fcn_lugar AND tcpr.prod_clave = tdf.prod_clave " +
                        "AND tmfn.fcn_lugar <> 'FC' AND tmfn.fcn_pagofletegab = 'S' ORDER BY tmfn.fcn_monedaflete, tmfn.fcn_folio, tmfn.fcn_fecha, tdf.prod_clave";*/
                /*query = "SELECT tmfn.fcn_fecha as fecha, tmfn.fcn_transporte as flete, Tmfn.fcn_folio as folio, tmfn.fcn_monto_transporte as monto1, tmfn.prov_clave as proveedor, " +
                        "tmfn.fcn_tipo as tipo, tmfn.fcn_lugar as lugar, tdf.PROD_CLAVE as producto, tdf.FCN_NUM_UNIDADES as unicajas, tcpr.prod_nombre as nproducto, tcpr.prod_num_tarimas as unitarim, " +
                        "tcpr.lin_clave as CVEL, tmfn.fcn_monedaflete, tmfn.fcn_ivaflete FROM tb_mstr_facturas_nal tmfn, tb_det_facturas tdf, " +
                        "tb_cat_producto tcpr WHERE (tmfn.fcn_fecha BETWEEN '" + FI + "' AND '" + FF + "') AND " +
                        "(tmfn.fcn_transporte <='zzzzzz' AND tmfn.fcn_transporte >='000000') AND (tmfn.prov_clave BETWEEN '000000' AND 'zzzzzz') AND tmfn.fcn_estatus <> 'C' AND " +
                        "tmfn.fcn_transporte <> '' AND tdf.fcn_folio= tmfn.fcn_folio AND tdf.fcn_tipo = tmfn.fcn_lugar AND tcpr.prod_clave = tdf.prod_clave " +
                        "AND tmfn.fcn_lugar <> 'FC' and tmfn.fcn_pagofletegab <> 'N' ORDER BY tmfn.fcn_monedaflete, tmfn.fcn_folio, tmfn.fcn_fecha, tdf.prod_clave";*/
                /*query = "SELECT tmfn.fcn_fecha as fecha, tmfn.fcn_transporte as flete, Tmfn.fcn_folio as folio, tmfn.fcn_monto_transporte as monto1, tmfn.prov_clave as proveedor, " +
                        "tmfn.fcn_tipo as tipo, tmfn.fcn_lugar as lugar, tdf.PROD_CLAVE as producto, tdf.FCN_NUM_UNIDADES as unicajas, tcpr.prod_nombre as nproducto, " +
                        "tcpr.prod_num_tarimas as unitarim, tcpr.lin_clave as CVEL, tmfn.fcn_monedaflete, tmfn.fcn_ivaflete, tmfn.pdn_folio, tmfn.cnte_clave, tmfn.fcn_elaboro, " +
                        "tmfn.fcn_tipoflete, tmfn.fcn_pagofletegab FROM tb_mstr_facturas_nal tmfn, tb_det_facturas tdf, tb_cat_producto tcpr " +
                        "WHERE (tmfn.fcn_fecha BETWEEN '" + FI + "' AND '" + FF + "') AND (tmfn.fcn_transporte <='zzzzzz' AND tmfn.fcn_transporte >='000000') " +
                        "AND (tmfn.prov_clave BETWEEN '000000' AND 'zzzzzz') AND tmfn.fcn_estatus <> 'C' AND tmfn.fcn_transporte <> '' AND tdf.fcn_folio= tmfn.fcn_folio " +
                        "AND tdf.fcn_tipo = tmfn.fcn_lugar AND tcpr.prod_clave = tdf.prod_clave AND tmfn.fcn_lugar <> 'FC' " +
                        "and tmfn.fcn_pagofletegab <> 'N' ORDER BY tmfn.fcn_monedaflete, tmfn.fcn_folio, tmfn.fcn_fecha, tdf.prod_clave";*/
                // SE QUITO ESTE CODIGO PORQUE YA INCLUYE A CANCUN PARA LOS CALCULOS  RCC 22/ABR/19
                //query = "SELECT tmfn.fcn_fecha as fecha, tmfn.fcn_transporte as flete, Tmfn.fcn_folio as folio, tmfn.fcn_monto_transporte as monto1, tmfn.prov_clave as proveedor, " +
                //        "tmfn.fcn_tipo as tipo, tmfn.fcn_lugar as lugar, tdf.PROD_CLAVE as producto, tdf.FCN_NUM_UNIDADES as unicajas, tcpr.prod_nombre as nproducto, " +
                //        "tcpr.prod_num_tarimas as unitarim, tcpr.lin_clave as CVEL, tmfn.fcn_lugar, tmfn.fcn_ivaflete, tmfn.pdn_folio, tmfn.cnte_clave, tmfn.fcn_elaboro, " +
                //        "tmfn.fcn_tipoflete, tmfn.fcn_pagofletegab, tmfn.fcn_monedaflete, tmfn.fcn_pedorigen FROM tb_mstr_facturas_nal tmfn, tb_det_facturas tdf, tb_cat_producto tcpr " +
                //        "WHERE (tmfn.fcn_fecha BETWEEN '" + FI + "' AND '" + FF + "') AND (tmfn.fcn_transporte <='zzzzzz' AND tmfn.fcn_transporte >='000000') " +
                //        "AND (tmfn.prov_clave BETWEEN '000000' AND 'zzzzzz') AND tmfn.fcn_estatus <> 'C' AND tmfn.fcn_transporte <> '' AND tdf.fcn_folio= tmfn.fcn_folio " +
                //        "AND tdf.fcn_tipo = tmfn.fcn_lugar AND tcpr.prod_clave = tdf.prod_clave AND tmfn.fcn_lugar <> 'FC' " +
                //        "and tmfn.fcn_pagofletegab <> 'N' " +
                //        //"ORDER BY tmfn.fcn_monedaflete, TMFN.um_clave, tmfn.fcn_folio, tmfn.fcn_fecha, tdf.prod_clave";
                //        "ORDER BY tmfn.fcn_tipo, tmfn.fcn_folio, tmfn.fcn_fecha, tdf.prod_clave";
                query = "SELECT tmfn.fcn_fecha as fecha, tmfn.fcn_transporte as flete, Tmfn.fcn_folio as folio, tmfn.fcn_monto_transporte as monto1, tmfn.prov_clave as proveedor, " +
                        "tmfn.fcn_tipo as tipo, tmfn.fcn_lugar as lugar, tdf.PROD_CLAVE as producto, tdf.FCN_NUM_UNIDADES as unicajas, tcpr.prod_nombre as nproducto, " +
                        "tcpr.prod_num_tarimas as unitarim, tcpr.lin_clave as CVEL, tmfn.fcn_lugar, tmfn.fcn_ivaflete, tmfn.pdn_folio, tmfn.cnte_clave, tmfn.fcn_elaboro, " +
                        "tmfn.fcn_tipoflete, tmfn.fcn_pagofletegab, tmfn.fcn_monedaflete, tmfn.fcn_pedorigen FROM tb_mstr_facturas_nal tmfn, tb_det_facturas tdf, tb_cat_producto tcpr " +
                        "WHERE (tmfn.fcn_fecha BETWEEN '" + FI + "' AND '" + FF + "') AND (tmfn.fcn_transporte <='zzzzzz' AND tmfn.fcn_transporte >='000000') " +
                        "AND (tmfn.prov_clave BETWEEN '000000' AND 'zzzzzz') AND tmfn.fcn_estatus <> 'C' AND tmfn.fcn_transporte <> '' AND tdf.fcn_folio= tmfn.fcn_folio " +
                        "AND tdf.fcn_tipo = tmfn.fcn_lugar AND tcpr.prod_clave = tdf.prod_clave " +
                        //"and tmfn.fcn_pagofletegab <> 'N' " +
                        //"/*and tmfn.fcn_transporte in ('TBC', 'TBA','CCP-013560','CTWA-10516','CCP-013791','CCP-013910','CCP-0115','CCP-014233','CCP-014356','CCP-014504','CCP-014683','CCP-29')*/ " +
                        //"and tmfn.fcn_transporte in ('SE58248') " +
                        //"and tdf.prod_clave = '16TOOCA121' " +
                        //" and tdf.prod_clave = '16TOWHTO64'" +
                        //"ORDER BY tmfn.fcn_monedaflete, TMFN.um_clave, tmfn.fcn_folio, tmfn.fcn_fecha, tdf.prod_clave"; and tdf.prod_clave = '16TOWHTO64' and tmfn.fcn_transporte = 'CCP-0115'
                        "ORDER BY tmfn.fcn_tipo, tmfn.fcn_folio, tmfn.fcn_fecha, tdf.prod_clave";
                letrero = "Moneda Nacional";
                da = new SqlDataAdapter(query, thisConnection);
                da.Fill(des, "tmr1");
                tmr1 = des.Tables["tmr1"];
                query = "SELECT fcn_transporte, sum(fcn_monto_transporte) AS importe FROM tb_mstr_facturas_nal " +
                        "WHERE fcn_fecha BETWEEN '" + FI + "' AND '" + FF + "' AND (fcn_transporte <='zzzzzz' " +
                        "AND fcn_transporte >='000000') AND (prov_clave BETWEEN '000000' AND 'zzzzzz') " +
                        //"AND fcn_estatus <> 'C' AND fcn_transporte <> '' and fcn_pagofletegab <> 'N' " +
                        "AND fcn_estatus <> 'C' AND fcn_transporte <> '' " +
                        "GROUP BY fcn_transporte ORDER BY fcn_transporte  ";
                da = new SqlDataAdapter(query, thisConnection);
                des = new DataSet();
                da.Fill(des, "CtaPor");
                DataTable CtaPor = des.Tables["CtaPor"];
                thisConnection.Close();

                if (tmr1.Rows.Count == 0)
                {
                    MessageBox.Show("No se encontro información dentro del periodo de fechas seleccionado", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }


                //foreach (DataRow row in tmr1.Rows)
                //{
                //    foreach (DataRow row1 in productos.Select("prod_clave = '" + Convert.ToString(row["prod_clave"].ToString().Trim()) + "'"))
                //    {
                //        row["nproducto"] = Convert.ToString(row1["prod_nombre"].ToString().Trim());
                //        row["
                //    }
                //}

                foreach (DataRow row in tmr1.Rows)
                {
                    find = false;
                    if (Convert.ToString(row["tipo"].ToString().Trim()) == "EXP")
                        row["tipo"] = "EXP";

                    else
                        row["tipo"] = "NAL";
                    //if (Convert.ToString(row["fcn_monedaflete"].ToString().Trim()) == "USD")
                    if (Convert.ToString(row["fcn_monedaflete"].ToString().Trim()) == "USD")
                    {
                        foreach (DataRow row2 in CtaPor.Select("fcn_transporte = '" + row["flete"].ToString() + "'")) //<+ "' and flete = '" + row["flete"].ToString() + "'"))
                        {
                            //row["monto1"] = Convert.ToDouble(row["monto1"]) + Convert.ToDouble(row2["importe"].ToString());
                            row["monto1"] = Convert.ToDouble(row2["importe"].ToString());
                        }
                    }
                    if (Convert.ToString(row["fcn_monedaflete"].ToString().Trim()) == "")
                    {
                        //MessageBox.Show("El flete con la factura: " + Convert.ToString(row["folio"].ToString().Trim()) + " no tiene el tipo de moneda del flete\r\nSe va a tomar la de default si es nacional pesos, si es exportación usd", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);                            
                        if (Convert.ToString(row["tipo"].ToString().Trim()) == "EXP")
                            row["fcn_monedaflete"] = "USD";

                        else
                            row["fcn_monedaflete"] = "PESOS";


                        if (Convert.ToString(row["fcn_monedaflete"].ToString().Trim()) == "USD")
                        {
                            //foreach (DataRow row2 in tmr1.Select("pdn_folio = '" + row["fcn_pedorigen"].ToString() + "' and flete = '" + row["flete"].ToString() + "'"))
                            foreach (DataRow row2 in tmr1.Select("flete = '" + row["flete"].ToString() + "'")) //<+ "' and flete = '" + row["flete"].ToString() + "'"))
                            {
                                if (Convert.ToDouble(row2["monto1"].ToString()) != 0)
                                {
                                    row["monto1"] = Convert.ToDouble(row["monto1"]) + Convert.ToDouble(row2["monto1"].ToString());
                                    //break;
                                }
                            }
                            foreach (DataRow row1 in tipo_cambio.Select("fecha = '" + Convert.ToString(row["fecha"].ToString().Trim()) + "'"))
                            {
                                row["monto1"] = Convert.ToDouble(row["monto1"].ToString()) * Convert.ToDouble(row1["valor"].ToString().Trim());
                                find = true;
                            }
                            if (find == false)
                                MessageBox.Show("No se ha capturado el tipo de cambio del día " + Convert.ToString(row["fecha"].ToString().Substring(0, 10)));
                        }
                    }

                    if ((Convert.ToString(row["fcn_monedaflete"].ToString().Trim()) == "PESOS") && (Convert.ToString(row["tipo"].ToString().Trim()) == "EXP"))
                    {
                        //MessageBox.Show(row["folio"].ToString());
                        //foreach (DataRow row2 in tmr1.Select("flete = '" + row["flete"].ToString() + "' and fcn_monedaflete = '" + row["fcn_monedaflete"].ToString().Trim() + "' and tipo = '" + row["tipo"].ToString().Trim() + "'"))
                        foreach (DataRow row2 in tmr1.Select("pdn_folio = '" + row["fcn_pedorigen"].ToString() + "' and flete = '" + row["flete"].ToString() + "'"))
                        {
                            if (Convert.ToDouble(row2["monto1"].ToString()) != 0)
                            {
                                row["monto1"] = Convert.ToDouble(row2["monto1"].ToString());
                                break;
                            }
                        }
                        foreach (DataRow row1 in tipo_cambio.Select("fecha = '" + Convert.ToString(row["fecha"].ToString().Trim()) + "'"))
                        {
                            row["monto1"] = Convert.ToDouble(row["monto1"].ToString()) / Convert.ToDouble(row1["valor"].ToString().Trim());
                            find = true;
                        }
                    }
                    //if ((Convert.ToString(row["fcn_monedaflete"].ToString().Trim()) == "") && (factura != Convert.ToString(row["folio"].ToString().Trim())))
                    //{
                    //    MessageBox.Show("El flete con la factura: " + Convert.ToString(row["folio"].ToString().Trim()) + " no tiene el tipo de moneda del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //    factura = Convert.ToString(row["folio"].ToString().Trim());
                    //}
                    //if (Convert.ToString(row["fcn_monedaflete"].ToString().Trim()) == "")
                    //{
                    //    //if (Convert.ToString(row["tipo"].ToString().Trim()) != "EXP")
                    //    //    row["fcn_monedaflete"] = "PESOS";

                    //    //if (Convert.ToString(row["tipo"].ToString().Trim()) == "EUROS")
                    //    //    row["fcn_monedaflete"] = "USD";
                    //    if (Convert.ToString(row["fcn_lugar"].ToString().Trim()) != "EXP")
                    //        row["fcn_monedaflete"] = "PESOS";

                    //    if (Convert.ToString(row["fcn_lugar"].ToString().Trim()) == "EUROS")
                    //        row["fcn_monedaflete"] = "USD";
                    //}
                }


                DataView dv = new DataView();
                dv = tmr1.DefaultView;
                //dv.Sort = "fcn_monedaflete, flete";
                dv.Sort = "tipo, flete";
                tmr1 = dv.ToTable();

                //foreach (DataRow row in tmr1.Rows)
                //{
                //    find = false;
                //    if ((Convert.ToString(row["fcn_monedaflete"].ToString().Trim()) == "USD") && (Convert.ToString(row["fcn_lugar"].ToString().Trim()) != "EXP"))
                //    //if (Convert.ToString(row["fcn_monedaflete"].ToString().Trim()) == "EUROS")
                //    {
                //        foreach (DataRow row1 in tipo_cambio.Select("fecha = '" + Convert.ToString(row["fecha"].ToString().Trim()) + "'"))
                //        {
                //            row["monto1"] = Convert.ToDouble(row["monto1"].ToString()) * Convert.ToDouble(row1["valor"].ToString().Trim());
                //            find = true;
                //        }
                //        if (find == false)
                //            MessageBox.Show("No se ha capturado el tipo de cambio del día " + Convert.ToString(row["fecha"].ToString().Substring(0, 10)));
                //    }
                //    if ((Convert.ToString(row["fcn_monedaflete"].ToString().Trim()) == "") && (factura != Convert.ToString(row["folio"].ToString().Trim())))
                //    {
                //        MessageBox.Show("El flete con la factura: " + Convert.ToString(row["folio"].ToString().Trim()) + " no tiene el tipo de moneda del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //        factura = Convert.ToString(row["folio"].ToString().Trim());
                //    }
                //    if (Convert.ToString(row["fcn_monedaflete"].ToString().Trim()) == "")
                //    {
                //        if (Convert.ToString(row["tipo"].ToString().Trim()) != "EXP")
                //            row["fcn_monedaflete"] = "PESOS";

                //        if (Convert.ToString(row["tipo"].ToString().Trim()) == "EUROS")
                //            row["fcn_monedaflete"] = "USD";
                //        //if (Convert.ToString(row["fcn_lugar"].ToString().Trim()) != "EXP")
                //        //    row["fcn_monedaflete"] = "PESOS";

                //        //if (Convert.ToString(row["fcn_lugar"].ToString().Trim()) == "EUROS")
                //        //    row["fcn_monedaflete"] = "USD";
                //    }
                //    if (Convert.ToString(row["fcn_LUGAR"].ToString().Trim()) != "EXP")
                //        row["fcn_lugar"] = "NAL";
                //}


                detalle_factura = tmr1.Copy();

                tot = 0; mon = 0;
                progressBar1.Maximum = tmr1.Rows.Count;
                progressBar1.Visible = true;
                string mtip = "";
                for (int i = 0; i < tmr1.Rows.Count; i++)
                //foreach(DataRow dr in tmr1.Rows)
                {
                    //if (fle2 == "30393")
                    //    MessageBox.Show("30393");
                    int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                    progressBar1.CreateGraphics().DrawString("Generando reporte . . .  " + percent.ToString() + "%",
                        new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 60, progressBar1.Height / 2 - 7));
                    fle = tmr1.Rows[i]["flete"].ToString().Trim();
                    if (i == 0)
                    {
                        tot = tot + Convert.ToDouble(tmr1.Rows[i]["unicajas"].ToString().Trim());

                        if (Convert.ToDecimal(tmr1.Rows[i]["monto1"].ToString().Trim()) > 0)
                            mon = Convert.ToDouble(tmr1.Rows[i]["monto1"].ToString().Trim());
                        else
                            tmr1.Rows[i]["monto1"] = Convert.ToDouble(tmr1.Rows[i]["monto1"].ToString().Trim());
                    }
                    //if (i != 0)
                    else
                    {
                        //if (fle != tmr1.Rows[i - 1]["flete"].ToString().Trim())
                        if (fle == tmr1.Rows[i - 1]["flete"].ToString().Trim())
                        {
                            tot = tot + Convert.ToDouble(tmr1.Rows[i]["unicajas"].ToString().Trim());

                            if (Convert.ToDecimal(tmr1.Rows[i]["monto1"].ToString().Trim()) > 0)
                                mon = Convert.ToDouble(tmr1.Rows[i]["monto1"].ToString().Trim());
                            else
                                tmr1.Rows[i]["monto1"] = Convert.ToDouble(tmr1.Rows[i]["monto1"].ToString().Trim());
                        }
                        else
                        {
                            fle2 = tmr1.Rows[i - 1]["flete"].ToString().Trim();
                            fle = tmr1.Rows[i]["flete"].ToString().Trim();
                            t = Convert.ToDouble(tmr1.Rows[i]["unicajas"].ToString().Trim());
                            //if (fle2 == "30393")
                            //    MessageBox.Show("30393");
                            mtip = tmr1.Rows[i - 1]["tipo"].ToString().Trim();
                            foreach (DataRow row in tmr1.Select("tipo = '" + tmr1.Rows[i - 1]["tipo"].ToString().Trim() + "' and flete = '" + fle2 + "'"))
                            {
                                row["cajtot"] = tot.ToString("###,###,###,##0.0000");
                                row["monto1"] = mon.ToString("###,###,##0.0000");
                                if (Convert.ToString(row["fcn_ivaflete"].ToString().Trim()) != "N")
                                    row["monto2"] = (Convert.ToDouble(Convert.ToDouble(row["monto1"].ToString()) / porceiva)).ToString("###,###,##0.0000");
                                else
                                    row["monto2"] = (Convert.ToDouble(Convert.ToDouble(row["monto1"].ToString())).ToString("###,###,##0.0000"));
                            }
                            tot = t;
                            mon = Convert.ToDouble(tmr1.Rows[i]["monto1"].ToString().Trim());
                        }
                    }
                    //else
                    //{
                    //    tot = tot + Convert.ToDouble(tmr1.Rows[i]["unicajas"].ToString().Trim());

                    //    if (Convert.ToDecimal(tmr1.Rows[i]["monto1"].ToString().Trim()) > 0)
                    //        mon = Convert.ToDouble(tmr1.Rows[i]["monto1"].ToString().Trim());
                    //}
                    progressBar1.PerformStep();
                }//for                    
                //foreach (DataRow row in tmr1.Select("tipo = '" + mtip.ToString().Trim() + "' and flete = '" + fle2 + "'"))
                //{
                //    row["cajtot"] = tot.ToString("###,###,###,##0.0000");
                //    row["monto1"] = mon.ToString("###,###,##0.0000");
                //    if (Convert.ToString(row["fcn_ivaflete"].ToString().Trim()) != "N")
                //        row["monto2"] = (Convert.ToDouble(Convert.ToDouble(row["monto1"].ToString()) / porceiva)).ToString("###,###,##0.0000");
                //    else
                //        row["monto2"] = (Convert.ToDouble(Convert.ToDouble(row["monto1"].ToString())).ToString("###,###,##0.0000"));
                //}

                fle2 = fle;
                fle = tmr1.Rows[tmr1.Rows.Count - 1]["flete"].ToString().Trim();
                foreach (DataRow row in tmr1.Select("tipo = '" + tmr1.Rows[tmr1.Rows.Count - 1]["tipo"].ToString().Trim() + "' and flete = '" + fle2 + "'"))
                {
                    row["cajtot"] = tot.ToString("###,###,###,##0.00");
                    row["monto1"] = mon.ToString("###,###,##0.00");
                    if (Convert.ToString(row["fcn_ivaflete"].ToString().Trim()) != "N")
                        row["monto2"] = (Convert.ToDouble(Convert.ToDouble(row["monto1"].ToString()) / porceiva)).ToString("###,###,##0.0000");
                    else
                        row["monto2"] = (Convert.ToDouble(Convert.ToDouble(row["monto1"].ToString())).ToString("###,###,##0.0000"));
                    //row["monto2"] = (Convert.ToDouble(Convert.ToDouble(row["monto1"].ToString()) / porceiva)).ToString("###,###,##0.00");
                }


                //dv = tmr1.DefaultView;
                //dv.Sort = "fcn_monedaflete,nproducto";
                //tmr1 = dv.ToTable();

                //for (int i = 0; i < tmr1.Rows.Count; i++)
                progressBar1.Value = 0;
                progressBar1.Maximum = tmr1.Rows.Count;
                progressBar1.Update();
                foreach (DataRow regis in tmr1.Rows)
                {
                    int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                    progressBar1.CreateGraphics().DrawString("Generando reporte . . .  " + percent.ToString() + "%",
                        new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 60, progressBar1.Height / 2 - 7));
                    regis["cosxcaja"] = (Convert.ToDouble((Convert.ToDouble(regis["monto2"].ToString()) * Convert.ToDouble(regis["unicajas"].ToString())) / Convert.ToDouble(regis["cajtot"].ToString()))).ToString("###,###,##0.0000");
                    regis["impuesto"] = (Convert.ToDouble(Convert.ToDouble(regis["monto1"].ToString()) - Convert.ToDouble(regis["monto2"].ToString()))).ToString("###,###,##0.0000");

                    if (Convert.ToDouble(regis["unitarim"].ToString()) > 0)
                        regis["tartot"] = (Convert.ToDouble(Convert.ToDouble(regis["unicajas"].ToString()) / Convert.ToDouble(regis["unitarim"].ToString()))).ToString("###,###,##0.0000");
                    else
                        regis["tartot"] = 0;
                    progressBar1.PerformStep();
                }//for

                progressBar1.Value = 0;
                progressBar1.Maximum = tmr1.Rows.Count;
                progressBar1.Update();
                tot = 0;
                for (int i = 0; i < tmr1.Rows.Count; i++)
                {
                    int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                    progressBar1.CreateGraphics().DrawString("Generando reporte . . .  " + percent.ToString() + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black,
                        new PointF(progressBar1.Width / 2 - 60, progressBar1.Height / 2 - 7));
                    fle = tmr1.Rows[i]["flete"].ToString().Trim();
                    if (i != 0)
                    {
                        if (fle != tmr1.Rows[i - 1]["flete"].ToString().Trim())
                        {
                            fle2 = tmr1.Rows[i - 1]["flete"].ToString().Trim();
                            fle = tmr1.Rows[i]["flete"].ToString().Trim();
                            t = Convert.ToDouble(tmr1.Rows[i]["tartot"].ToString().Trim());

                            foreach (DataRow row in tmr1.Select("tipo = '" + tmr1.Rows[i - 1]["tipo"].ToString().Trim() + "' and flete = '" + fle2 + "'"))
                            {
                                row["tartot2"] = tot.ToString("###,###,##0.0000");
                            }
                            tot = t;
                        }
                        else
                        {
                            tot = tot + Convert.ToDouble(tmr1.Rows[i]["tartot"].ToString());
                        }
                    }
                    else
                        tot = tot + Convert.ToDouble(tmr1.Rows[i]["tartot"].ToString());
                    progressBar1.PerformStep();
                }//for

                fle2 = fle;
                fle = tmr1.Rows[tmr1.Rows.Count - 1]["flete"].ToString().Trim();
                foreach (DataRow row1 in tmr1.Select("tipo = '" + tmr1.Rows[tmr1.Rows.Count - 1]["tipo"].ToString().Trim() + "' and flete = '" + fle2 + "'"))
                {
                    row1["tartot2"] = tot.ToString("###,###,##0.0000");
                }

                progressBar1.Value = 0;
                progressBar1.Maximum = tmr1.Rows.Count;
                progressBar1.Update();

                //dv = tmr1.DefaultView;
                //dv.Sort = "fcn_monedaflete,nproducto";
                //tmr1 = dv.ToTable();

                //for (int i = 0; i < tmr1.Rows.Count; i++)
                foreach (DataRow registro in tmr1.Rows)
                {
                    int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                    progressBar1.CreateGraphics().DrawString("Generando reporte . . .  " + percent.ToString() + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 60, progressBar1.Height / 2 - 7));
                    if (Convert.ToDouble(registro["tartot2"].ToString()) > 0)//tmr1.Rows[i]["tartot2"].ToString()) > 0)
                    {
                        double res = Convert.ToDouble((Convert.ToDouble(registro["monto2"].ToString()) * Convert.ToDouble(registro["tartot"].ToString())) / Convert.ToDouble(registro["tartot2"].ToString()));

                        if (res != 0)
                            registro["cosxtari"] = res.ToString("###,###,###,##0.0000");
                    }
                    progressBar1.PerformStep();
                }

                caj = 0; tar = 0; dincaj = 0; dintar = 0;
                //DataView dv = new DataView();
                dv = tmr1.DefaultView;
                //dv.Sort = "fcn_monedaflete, producto";
                dv.Sort = "tipo, producto";
                tmr1 = dv.ToTable();

                progressBar1.Value = 0;
                progressBar1.Maximum = tmr1.Rows.Count;
                progressBar1.Update();
                for (int i = 0; i < tmr1.Rows.Count; i++)
                {
                    int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                    progressBar1.CreateGraphics().DrawString("Generando reporte . . .  " + percent.ToString() + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 60, progressBar1.Height / 2 - 7));
                    pro = tmr1.Rows[i]["producto"].ToString().Trim();
                    if (i != 0)
                    {
                        if (pro != tmr1.Rows[i - 1]["producto"].ToString().Trim())
                        {
                            pro2 = tmr1.Rows[i - 1]["producto"].ToString();
                            pro = tmr1.Rows[i]["producto"].ToString().Trim();
                            t1 = Convert.ToDouble(tmr1.Rows[i]["unicajas"].ToString());
                            t2 = Convert.ToDouble(tmr1.Rows[i]["tartot"].ToString());
                            t3 = Convert.ToDouble(tmr1.Rows[i]["cosxcaja"].ToString());
                            t4 = Convert.ToDouble(tmr1.Rows[i]["cosxtari"].ToString());

                            foreach (DataRow row in tmr1.Select("tipo = '" + tmr1.Rows[i - 1]["tipo"].ToString().Trim() + "' and producto = '" + pro2 + "'"))
                            {
                                row["totcajp"] = caj.ToString("###,###,##0.0000");
                                row["tottarp"] = tar.ToString("###,###,##0.0000");
                                row["coscajp"] = dincaj.ToString("###,###,##0.0000");
                                row["costarp"] = dintar.ToString("###,###,##0.0000");
                            }

                            caj = t1;
                            tar = t2;
                            dincaj = t3;
                            dintar = t4;
                        }
                        else
                        {
                            caj = caj + Convert.ToDouble(tmr1.Rows[i]["unicajas"].ToString());
                            tar = tar + Convert.ToDouble(tmr1.Rows[i]["tartot"].ToString());
                            dincaj = dincaj + Convert.ToDouble(tmr1.Rows[i]["cosxcaja"].ToString());
                            dintar = dintar + Convert.ToDouble(tmr1.Rows[i]["cosxtari"].ToString());
                        }
                    }
                    else
                    {
                        caj = caj + Convert.ToDouble(tmr1.Rows[i]["unicajas"].ToString());
                        tar = tar + Convert.ToDouble(tmr1.Rows[i]["tartot"].ToString());
                        dincaj = dincaj + Convert.ToDouble(tmr1.Rows[i]["cosxcaja"].ToString());
                        dintar = dintar + Convert.ToDouble(tmr1.Rows[i]["cosxtari"].ToString());
                    }
                    progressBar1.PerformStep();
                }//for

                pro2 = pro;
                pro = tmr1.Rows[tmr1.Rows.Count - 1]["producto"].ToString().Trim();
                foreach (DataRow row in tmr1.Select("tipo = '" + tmr1.Rows[tmr1.Rows.Count - 1]["tipo"].ToString().Trim() + "' and producto = '" + pro2 + "'"))
                {
                    row["totcajp"] = caj.ToString("###,###,##0.0000");
                    row["tottarp"] = tar.ToString("###,###,##0.0000");
                    row["coscajp"] = dincaj.ToString("###,###,##0.0000");
                    row["costarp"] = dintar.ToString("###,###,##0.0000");
                }

                DataTable temporal = new DataTable();
                temporal = tmr1.Clone();
                //double unicajas = 0, tartot = 0, cosxcaja = 0, cosxtari = 0;
                //for (int i = 0; i < tmr1.Rows.Count; i++)
                //{
                //    pro = tmr1.Rows[i]["producto"].ToString().Trim();
                //    if (i != 0)
                //    {                            
                //        if (pro != tmr1.Rows[i - 1]["producto"].ToString().Trim())
                //        {
                //            unicajas = 0; tartot = 0; cosxcaja = 0; cosxtari = 0;
                //            foreach (DataRow reg in tmr1.Select("producto = '" + tmr1.Rows[i - 1]["producto"].ToString().Trim() + "'"))
                //            {
                //                unicajas = unicajas + Convert.ToDouble(reg["unicajas"].ToString());
                //                tartot = tartot + Convert.ToDouble(reg["tartot"].ToString());
                //                cosxcaja = cosxcaja + Convert.ToDouble(reg["cosxcaja"].ToString());
                //                cosxtari = cosxtari + Convert.ToDouble(reg["cosxtari"].ToString());
                //            }

                //            DataRow fila = temporal.NewRow();
                //            fila["producto"] = tmr1.Rows[i -1]["producto"].ToString();
                //            fila["nproducto"] = tmr1.Rows[i - 1]["nproducto"].ToString();
                //            fila["unicajas"] = unicajas.ToString("###,###,##0.00");
                //            fila["tartot"] = tartot.ToString("###,###,##0.00");
                //            fila["cosxcaja"] = cosxcaja.ToString("###,###,##0.00");
                //            fila["cosxtari"] = cosxtari.ToString("###,###,##0.00");
                //            fila["coscajp"] = Convert.ToDouble(tmr1.Rows[i - 1]["coscajp"].ToString());
                //            fila["totcajp"] = Convert.ToDouble(tmr1.Rows[i - 1]["totcajp"].ToString());
                //            fila["costarp"] = Convert.ToDouble(tmr1.Rows[i - 1]["costarp"].ToString());
                //            fila["tottarp"] = Convert.ToDouble(tmr1.Rows[i - 1]["tottarp"].ToString());
                //            fila["CVEL"] = tmr1.Rows[i - 1]["CVEL"].ToString();
                //            fila["NOMLINEA"] = tmr1.Rows[i - 1]["NOMLINEA"].ToString();
                //            temporal.Rows.Add(fila);
                //        }
                //    }
                //}

                progressBar1.Value = 0;
                progressBar1.Maximum = tmr1.Rows.Count;
                progressBar1.Update();
                /*
                double cajas = Convert.ToDouble(Convert.ToDouble(tem.Rows[i]["coscajp"].ToString()) / Convert.ToDouble(tem.Rows[i]["totcajp"].ToString()));
                double tar = Convert.ToDouble(Convert.ToDouble(tem.Rows[i]["costarp"].ToString()) / Convert.ToDouble(tem.Rows[i]["totcajp"].ToString()));
                double tari = Convert.ToDouble(Convert.ToDouble(tem.Rows[i]["costarp"].ToString()) / Convert.ToDouble(tem.Rows[i]["tottarp"].ToString()));
                 * */
                for (int i = 0; i < tmr1.Rows.Count; i++)
                {
                    int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                    progressBar1.CreateGraphics().DrawString("Generando reporte . . .  " + percent.ToString() + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 60, progressBar1.Height / 2 - 7));
                    bool find = false;
                    pro = tmr1.Rows[i]["producto"].ToString().Trim();
                    //if (tmr1.Rows[i]["fcn_monedaflete"].ToString().Trim() == "USD")
                    //    MessageBox.Show("LELGO A USD");
                    foreach (DataRow reg in temporal.Select("tipo = '" + tmr1.Rows[i]["tipo"].ToString().Trim() + "' and  producto ='" + pro + "'"))
                    {
                        reg["unicajas"] = Convert.ToDouble(reg["unicajas"].ToString()) + Convert.ToDouble(tmr1.Rows[i]["unicajas"].ToString());
                        reg["tartot"] = Convert.ToDouble(reg["tartot"].ToString()) + Convert.ToDouble(tmr1.Rows[i]["tartot"].ToString());
                        reg["cosxcaja"] = Convert.ToDouble(reg["cosxcaja"].ToString()) + Convert.ToDouble(tmr1.Rows[i]["cosxcaja"].ToString());
                        reg["cosxtari"] = Convert.ToDouble(reg["cosxtari"].ToString()) + Convert.ToDouble(tmr1.Rows[i]["cosxtari"].ToString());
                        /*reg["cajas"] = Convert.ToDouble(reg["cajas"].ToString()) +  (Convert.ToDouble(reg["coscajp"].ToString()) / Convert.ToDouble(tmr1.Rows[i]["totcajp"].ToString()));
                        reg["tar"] = Convert.ToDouble(reg["tar"].ToString()) + (Convert.ToDouble(reg["costarp"].ToString()) / Convert.ToDouble(tmr1.Rows[i]["totcajp"].ToString()));
                        reg["tari"] = Convert.ToDouble(reg["tari"].ToString()) + (Convert.ToDouble(reg["costarp"].ToString()) / Convert.ToDouble(tmr1.Rows[i]["tottarp"].ToString()));*/
                        find = true;
                    }
                    if (find == false)
                    {
                        DataRow fila = temporal.NewRow();
                        fila["producto"] = tmr1.Rows[i]["producto"].ToString();
                        fila["nproducto"] = tmr1.Rows[i]["nproducto"].ToString();
                        fila["unicajas"] = Convert.ToDouble(Convert.ToDouble(tmr1.Rows[i]["unicajas"].ToString()));
                        fila["tartot"] = Convert.ToDouble(Convert.ToDouble(tmr1.Rows[i]["tartot"].ToString()));
                        fila["cosxcaja"] = Convert.ToDouble(Convert.ToDouble(tmr1.Rows[i]["cosxcaja"].ToString()));
                        fila["cosxtari"] = Convert.ToDouble(Convert.ToDouble(tmr1.Rows[i]["cosxtari"].ToString()));
                        fila["coscajp"] = Convert.ToDouble(Convert.ToDouble(tmr1.Rows[i]["coscajp"].ToString()));
                        fila["totcajp"] = Convert.ToDouble(Convert.ToDouble(tmr1.Rows[i]["totcajp"].ToString()));
                        fila["costarp"] = Convert.ToDouble(Convert.ToDouble(tmr1.Rows[i]["costarp"].ToString()));
                        fila["tottarp"] = Convert.ToDouble(Convert.ToDouble(tmr1.Rows[i]["tottarp"].ToString()));
                        fila["CVEL"] = tmr1.Rows[i]["CVEL"].ToString();
                        fila["NOMLINEA"] = tmr1.Rows[i]["NOMLINEA"].ToString();
                        fila["fcn_monedaflete"] = tmr1.Rows[i]["fcn_monedaflete"].ToString();
                        fila["tipo"] = tmr1.Rows[i]["tipo"].ToString();
                        /*fila["cajas"] = Convert.ToDouble(tmr1.Rows[i]["coscajp"].ToString()) / Convert.ToDouble(tmr1.Rows[i]["totcajp"].ToString());
                        fila["tar"] = Convert.ToDouble(tmr1.Rows[i]["costarp"].ToString()) / Convert.ToDouble(tmr1.Rows[i]["totcajp"].ToString());
                        fila["tari"] = Convert.ToDouble(tmr1.Rows[i]["costarp"].ToString()) / Convert.ToDouble(tmr1.Rows[i]["tottarp"].ToString());*/
                        temporal.Rows.Add(fila);
                    }
                    progressBar1.PerformStep();
                }

                progressBar1.Visible = false;
                progressBar1.Value = 0;
                progressBar1.Update();
                string NSem = "0";
                string TipoC = "0";
                if (Tipo == "S")
                {
                    thisConnection.Open();
                    string Cadena = "SELECT semana FROM tb_cat_semanas " +
                             " Where fecha1 = '" + DTPInifecha.Value.ToShortDateString() + "'order by ano,semana";
                    SqlCommand cmd = new SqlCommand(Cadena, thisConnection);
                    NSem = Convert.ToString(cmd.ExecuteScalar()).Trim();
                    NSem = (NSem.ToString().Trim() == "") ? "0" : NSem;
                    Cadena = "SELECT valor FROM tb_cat_tipocambio WHERE fecha = '" + DTPFinfecha.Value.ToShortDateString() + "'";
                    cmd = new SqlCommand(Cadena, thisConnection);
                    TipoC = Convert.ToString(cmd.ExecuteScalar()).Trim();
                    TipoC = (TipoC.Trim() == "") ? "0" : TipoC;
                    thisConnection.Close();
                }
                dv = temporal.DefaultView;
                //dv.Sort = "fcn_monedaflete, nproducto";
                dv.Sort = "tipo, nproducto";
                temporal = dv.ToTable();
                Reporte repo = new Reporte();
                repo.tem = temporal;
                repo.NoSem = Convert.ToInt32(NSem);
                repo.TipoCam = Convert.ToDecimal(TipoC);
                repo.detalle_factura = detalle_factura.Copy();
                repo.detalle_flete = tmr1.Copy();
                //repo.tem = tmr1;
                repo.ShowDialog();
                prov_ini = ""; prov_fin = ""; query = ""; FI = ""; FF = ""; fle = ""; fle2 = ""; letrero = ""; pro = ""; pro2 = ""; factura = "";
                tot = 0; mon = 1; t = 0; porceiva = 0;
                caj = 0; tar = 0; dincaj = 0; dintar = 0; t1 = 0; t2 = 0; t3 = 0; t4 = 0;
                //find = true;
                reporte = 0;
                des.Clear();
                temporal.Rows.Clear();
                tmr1.Rows.Clear();
            }

            catch (Exception ex1)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("ricardo.cortes@mrlucky.com.mx", "ricardo.cortes", "rcedillo", ex1.ToString());
                MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
    }

}
