using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Globalization;
using System.Net.Mail;
using System.Net;

namespace Reporte_costos_xtarima
{
    public partial class Reporte : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlCommand cmnd1 = new SqlCommand();
        SqlDataReader reader1;
        public DataTable tem;
        public DataTable yacasi = new DataTable("yacasi");
        public DataTable detalle_factura = new DataTable("detalle_factura");
        public DataTable productos = new DataTable("productos");
        public DataTable detalle_flete = new DataTable("detalle_flete");
        DataTable temporal = new DataTable();
        DataTable linea = new DataTable("linea");
        DataSet ds = new DataSet("yacasi");
        SqlDataAdapter da;
        public int count = 0;
        public static double totalunicajas = 0, totalunitarimas = 0, totalcostoscajas = 0, totalcostostarimas = 0;
        public int NoSem = 0, Año = 0;
        public decimal TipoCam = 0;
        string tipomoneda = "", prod = "";

        string tipo_reporte = "";
        string fech_ulti = "";
        public void makeyacasi()
        {
            yacasi.Columns.Add("producto");
            yacasi.Columns.Add("nproducto");
            yacasi.Columns.Add("linea");
            yacasi.Columns.Add("nomlinea");
            yacasi.Columns.Add("movimiento");
            ds.Tables.Add(yacasi);

            DataColumn columna2 = new DataColumn();
            columna2.ColumnName = "costo";
            columna2.DataType = System.Type.GetType("System.Decimal");
            columna2.DefaultValue = 0;
            yacasi.Columns.Add(columna2);
        }

        public Reporte()
        {
            InitializeComponent();

            tipo_reporte = Reporte_costos_xtarima.tipo_reporte;
            lblUltimoCalculo.Text = ultima_fecha_calculada();

            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);

            
        }

        private void Reporte_Load(object sender, EventArgs e)
        {            
            if (Reporte_costos_xtarima.reporte == 1)
            {
                productos = detalle_factura.Clone();
                LblSem.Text = (NoSem > 0) ? "Semana: " + NoSem.ToString() : "";
                LblTC.Text = (TipoCam > 0) ? "Tipo Cambio: " + TipoCam.ToString("##.0000") : "";
                DGV2.Visible = false;
                dataGridView1.Visible = true;
                label2.Text = "Del " + Reporte_costos_xtarima.FI;
                label1.Text = "Al " + Reporte_costos_xtarima.FF;
                Año = Convert.ToDateTime(Reporte_costos_xtarima.FI).Year;

                totalunicajas = 0; totalunitarimas = 0; totalcostoscajas = 0; totalcostostarimas = 0;

                dataGridView1.Rows.Insert(0, "", "", "Unidades", "", "Costos", "", "", "Costo Promedio Unitario");

                for (int i = 0; i < tem.Rows.Count; i++)
                {
                    count = dataGridView1.Rows.Count;
                    //tipomoneda = Convert.ToString(tem.Rows[i]["fcn_monedaflete"].ToString().Trim());
                    tipomoneda = Convert.ToString(tem.Rows[i]["tipo"].ToString().Trim());
                    double cajas = Convert.ToDouble(Convert.ToDouble(tem.Rows[i]["coscajp"].ToString()) / Convert.ToDouble(tem.Rows[i]["totcajp"].ToString()));
                    double tar = Convert.ToDouble(Convert.ToDouble(tem.Rows[i]["costarp"].ToString()) / Convert.ToDouble(tem.Rows[i]["totcajp"].ToString()));
                    double tari = Convert.ToDouble(Convert.ToDouble(tem.Rows[i]["costarp"].ToString()) / Convert.ToDouble(tem.Rows[i]["tottarp"].ToString()));                    
                    if (i == 0)
                        dataGridView1.Rows.Insert(count, tem.Rows[i]["nproducto"].ToString(), tem.Rows[i]["producto"].ToString(), tem.Rows[i]["unicajas"].ToString(), tem.Rows[i]["tartot"].ToString(), tem.Rows[i]["cosxcaja"].ToString(), tem.Rows[i]["cosxtari"].ToString(), cajas.ToString("###,###,###,##0.00"), tar.ToString("###,###,###,##0.0000"), tari.ToString("###,###,###,##0.00"),tipomoneda);
                    else
                    {
                        //cuando cambia la moneda
                        //if (tipomoneda.Trim() != Convert.ToString(tem.Rows[i - 1]["fcn_monedaflete"].ToString().Trim()))
                        if (tipomoneda.Trim() != Convert.ToString(tem.Rows[i - 1]["tipo"].ToString().Trim()))
                        {
                            totalunicajas = 0; totalunitarimas = 0; totalcostoscajas = 0; totalcostostarimas = 0;
//                            foreach (DataRow row in tem.Select("fcn_monedaflete = '" + Convert.ToString(tem.Rows[i - 1]["fcn_monedaflete"].ToString().Trim()) + "'"))
                            foreach (DataRow row in tem.Select("tipo = '" + Convert.ToString(tem.Rows[i - 1]["tipo"].ToString().Trim()) + "'"))
                            {
                                totalunicajas = totalunicajas + Convert.ToDouble(row["unicajas"].ToString());
                                totalunitarimas = totalunitarimas + Convert.ToDouble(row["tartot"].ToString());
                                totalcostoscajas = totalcostoscajas + Convert.ToDouble(row["cosxcaja"].ToString());
                                totalcostostarimas = totalcostostarimas + Convert.ToDouble(row["cosxtari"].ToString());
                            }
                            //dataGridView1.Rows.Insert(count, "TOTALES " + Convert.ToString(tem.Rows[i - 1]["fcn_monedaflete"].ToString().Trim()), "", totalunicajas.ToString("###,###,###,##0"), totalunitarimas.ToString("###,###,###,##0.00"), totalcostoscajas.ToString("###,###,###,##0"), totalcostostarimas.ToString("###,###,###,##0.00"));
                            dataGridView1.Rows.Insert(count, "TOTALES " + Convert.ToString(tem.Rows[i - 1]["tipo"].ToString().Trim()), "", totalunicajas.ToString("###,###,###,##0"), totalunitarimas.ToString("###,###,###,##0.00"), totalcostoscajas.ToString("###,###,###,##0.0000"), totalcostostarimas.ToString("###,###,###,##0.00"));
                            dataGridView1.Rows[count].DefaultCellStyle.BackColor = Color.Yellow;

                            count = dataGridView1.Rows.Count;
                            dataGridView1.Rows.Insert(count, tem.Rows[i]["nproducto"].ToString(), tem.Rows[i]["producto"].ToString(), tem.Rows[i]["unicajas"].ToString(), tem.Rows[i]["tartot"].ToString(), tem.Rows[i]["cosxcaja"].ToString(), tem.Rows[i]["cosxtari"].ToString(), cajas.ToString("###,###,###,##0.00"), tar.ToString("###,###,###,##0.0000"), tari.ToString("###,###,###,##0.00"),tipomoneda);
                        }
                        else
                            dataGridView1.Rows.Insert(count, tem.Rows[i]["nproducto"].ToString(), tem.Rows[i]["producto"].ToString(), tem.Rows[i]["unicajas"].ToString(), tem.Rows[i]["tartot"].ToString(), tem.Rows[i]["cosxcaja"].ToString(), tem.Rows[i]["cosxtari"].ToString(), cajas.ToString("###,###,###,##0.00"), tar.ToString("###,###,###,##0.0000"), tari.ToString("###,###,###,##0.00"),tipomoneda);

                        if (i == tem.Rows.Count - 1)
                        {
                            count = dataGridView1.Rows.Count;
                            totalunicajas = 0; totalunitarimas = 0; totalcostoscajas = 0; totalcostostarimas = 0;
                            //foreach (DataRow row in tem.Select("fcn_monedaflete = '" + Convert.ToString(tem.Rows[i]["fcn_monedaflete"].ToString().Trim()) + "'"))
                            foreach (DataRow row in tem.Select("tipo = '" + Convert.ToString(tem.Rows[i]["tipo"].ToString().Trim()) + "'"))
                            {
                                totalunicajas = totalunicajas + Convert.ToDouble(row["unicajas"].ToString());
                                totalunitarimas = totalunitarimas + Convert.ToDouble(row["tartot"].ToString());
                                totalcostoscajas = totalcostoscajas + Convert.ToDouble(row["cosxcaja"].ToString());
                                totalcostostarimas = totalcostostarimas + Convert.ToDouble(row["cosxtari"].ToString());
                            }
                            //dataGridView1.Rows.Insert(count, "TOTALES " + Convert.ToString(tem.Rows[i]["fcn_monedaflete"].ToString().Trim()), "", totalunicajas.ToString("###,###,###,##0"), totalunitarimas.ToString("###,###,###,##0.00"), totalcostoscajas.ToString("###,###,###,##0"), totalcostostarimas.ToString("###,###,###,##0.00"));
                            dataGridView1.Rows.Insert(count, "TOTALES " + Convert.ToString(tem.Rows[i]["tipo"].ToString().Trim()), "", totalunicajas.ToString("###,###,###,##0"), totalunitarimas.ToString("###,###,###,##0.00"), totalcostoscajas.ToString("###,###,###,##0.0000"), totalcostostarimas.ToString("###,###,###,##0.00"));
                            dataGridView1.Rows[count].DefaultCellStyle.BackColor = Color.Yellow;
                        }
                    }
                }


                /*foreach (DataRow fila in tem.Rows)
                {
                    count = dataGridView1.Rows.Count;
                    double cajas = Convert.ToDouble(Convert.ToDouble(fila["coscajp"].ToString()) / Convert.ToDouble(fila["totcajp"].ToString()));
                    double tar = Convert.ToDouble(Convert.ToDouble(fila["costarp"].ToString()) / Convert.ToDouble(fila["totcajp"].ToString()));
                    double tari = Convert.ToDouble(Convert.ToDouble(fila["costarp"].ToString()) / Convert.ToDouble(fila["tottarp"].ToString()));
                    dataGridView1.Rows.Insert(count, fila["nproducto"].ToString(), fila["producto"].ToString(), fila["unicajas"].ToString(), fila["tartot"].ToString(), fila["cosxcaja"].ToString(), fila["cosxtari"].ToString(), cajas.ToString("###,###,###,##0"), tar.ToString("###,###,###,##0.00"), tari.ToString("###,###,###,##0.00"));

                    totalunicajas = totalunicajas + Convert.ToDouble(fila["unicajas"].ToString());
                    totalunitarimas = totalunitarimas + Convert.ToDouble(fila["tartot"].ToString());
                    totalcostoscajas = totalcostoscajas + Convert.ToDouble(fila["cosxcaja"].ToString());
                    totalcostostarimas = totalcostostarimas + Convert.ToDouble(fila["cosxtari"].ToString());
                }
                count = dataGridView1.Rows.Count;
                dataGridView1.Rows.Insert(count, "TOTALES POR REPORTE", "", totalunicajas.ToString("###,###,###,##0"), totalunitarimas.ToString("###,###,###,##0.00"), totalcostoscajas.ToString("###,###,###,##0"), totalcostostarimas.ToString("###,###,###,##0.00"));*/
                if (TipoCam > 0)
                {
                    String Mone = "EXP";
                    dataGridView1.Columns.Add("ImpFlete", "Imp Flete");
                    int i = 0;
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        decimal Tar = (dataGridView1.Rows[i].Cells["ctarimas"].Value.ToString().Trim() == "") ? 0 : Convert.ToDecimal(dataGridView1.Rows[i].Cells["ctarimas"].Value);
                        if (dataGridView1.Rows[i].Cells["producto"].Value.ToString().Trim() == "TOTALES EXP")
                            Mone = "NAL";
                        dataGridView1.Rows[i].Cells["ImpFlete"].Value = (Tar * ((Mone == "EXP") ? TipoCam : 1)).ToString("#,###.#0");
                        i++;
                    }
                }

                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }
            if (Reporte_costos_xtarima.reporte == 2)
            {
                btngenera.Visible = false;
                DGV2.Visible = true;
                dataGridView1.Visible = false;
                label4.Visible = false;
                label4.Text = "";
                label3.Text = "Reporte de fletes";
                label2.Text = "Del " + Reporte_costos_xtarima.FI;
                label1.Text = "Al " + Reporte_costos_xtarima.FF;
                label5.Visible = false;
                count = 0;
                foreach (DataRow row in tem.Rows)
                {
                    count = DGV2.Rows.Count;
                    string fech = Convert.ToString(row["fcn_fecha"].ToString()).Substring(0, 10);
                    decimal imp_fle = Convert.ToDecimal(row["fcn_monto_transporte"].ToString());
                    DGV2.Rows.Insert(count, row["fcn_folio"].ToString().Trim(), fech, row["prov_clave"].ToString().Trim(), row["prov_nombre"].ToString().Trim(), row["fcn_transporte"].ToString().Trim(), imp_fle.ToString("###,###,##0.00"), row["pdn_folio"].ToString().Trim(), row["um_clave"].ToString().Trim());
                }
                DGV2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }

            Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "I", "4.19", "", "INGRESO A RECALCULO DE FLETES " + LblTC.Text, "SIPGAB");

            manda_correo_empaques_fletes();
        }

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            //comprobar que la fila y columna son las adecuadas
            if (e.RowIndex == 0 && (e.ColumnIndex == 2) || e.RowIndex == 0 && (e.ColumnIndex == 4) || e.RowIndex == 0 && (e.ColumnIndex == 6) || e.RowIndex == 0 && (e.ColumnIndex == 7))           
            {
                Color clrFondoCelda;
                Color clrTextoCelda;
                // en función de si la celda está o no seleccionada
                // establecer los colores
                if ((e.State & DataGridViewElementStates.Selected) == DataGridViewElementStates.Selected)
                {
                    clrFondoCelda = SystemColors.Highlight;
                    clrTextoCelda = SystemColors.Window;
                }
                else
                {
                    clrFondoCelda = SystemColors.Window;
                    clrTextoCelda = SystemColors.WindowText;
                }

                // rellenar el rectángulo de la celda con el color correspondiente
                e.Graphics.FillRectangle(new SolidBrush(clrFondoCelda), e.CellBounds);

                // dibujar solamente la línea vertical de la celda
                e.Graphics.DrawLine(new Pen(SystemColors.ActiveBorder),
                    new Point(e.CellBounds.X, e.CellBounds.Y + e.CellBounds.Height - 1),
                    new Point(e.CellBounds.X + e.CellBounds.Width, e.CellBounds.Y + e.CellBounds.Height - 1));

                // si la celda tiene valor
                if (e.Value != null)
                {
                    // dibujar el texto
                    TextRenderer.DrawText(e.Graphics, e.FormattedValue.ToString(), e.CellStyle.Font, e.CellBounds, clrTextoCelda, TextFormatFlags.Left);
                }

                e.Handled = true;
            }
        }

        //botón exportar reporte a excel
        private void btnExcel_Click(object sender, EventArgs e)
        {
            if(Reporte_costos_xtarima.reporte == 1)
                ReporteExcel(dataGridView1);
            if (Reporte_costos_xtarima.reporte == 2)
                ReporteExcel(DGV2);
        }

        //exporta datatable a excel
        void ReporteExcel(DataGridView MiGrid)
        {            
            Excel.Range r;

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Add();
                /*.Open(ruta, vk_update_links, vk_read_only, vk_format, vk_password,
            vk_write_res_password, vk_ignore_read_only_recommend, vk_origin,
            vk_delimiter, vk_editable, vk_notify, vk_converter, vk_add_to_mru,
            vk_local, vk_corrupt_load);*/
            Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1);

            //int count = MiDataGrid.ColumnCount;
            excelWorksheet.Cells[1, 1] = "4.19 " + label3.Text;
            excelWorksheet.Range[excelWorksheet.Cells[1, 1], excelWorksheet.Cells[1, 5]].Merge();
            excelWorksheet.Cells[1, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            r = excelWorksheet.get_Range("A1", "E1");
            r.Font.Bold = true;

            excelWorksheet.Cells[2, 1] = label4.Text;
            excelWorksheet.Range[excelWorksheet.Cells[2, 1], excelWorksheet.Cells[2, 5]].Merge();
            excelWorksheet.Cells[2, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            r = excelWorksheet.get_Range("A2", "E2");
            r.Font.Bold = true;
            
            excelWorksheet.Cells[3, 1] = label2.Text + " " + label1.Text;
            excelWorksheet.Range[excelWorksheet.Cells[3, 1], excelWorksheet.Cells[3, 5]].Merge();
            excelWorksheet.Cells[3, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            r = excelWorksheet.get_Range("A4", "E3");
            r.Font.Bold = true;

            int con = 1;
            for (int i = 0; i < MiGrid.Columns.Count; i++)
            {
                excelWorksheet.Cells[4, con] = MiGrid.Columns[i].HeaderText;                
                con++;
            }

            r = excelWorksheet.Range[excelWorksheet.Cells[7, 3], excelWorksheet.Cells[MiGrid.Rows.Count + 6, 9]];            
            r.NumberFormat = "#,##0.00";
            r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;


            MiGrid.SelectAll();
            MiGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            Clipboard.SetDataObject(MiGrid.GetClipboardContent().GetText());
            excelWorksheet.Cells[7, 1].Select();
            excelWorksheet.Paste();

            progressBar1.Visible = true;
            progressBar1.Maximum = MiGrid.Rows.Count;

            ////agrega las filas a excel
            int fila = 5;
            foreach (DataGridViewRow row in MiGrid.Rows)
            {                
                int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                progressBar1.CreateGraphics().DrawString("Exportando reporte . . .  " + percent.ToString() + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 60, progressBar1.Height / 2 - 7));
                if (Convert.ToString(row.Cells[0].Value.ToString().Trim()).Contains("TOTALES"))
                {
                    r = excelWorksheet.Range[excelWorksheet.Cells[fila + 2, 1], excelWorksheet.Cells[fila + 2, 9]];
                    r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    r = excelWorksheet.Range[excelWorksheet.Cells[fila + 2, 1], excelWorksheet.Cells[fila + 2, 9]];
                    r.Font.Bold = true;   
                }
                fila++;
                progressBar1.PerformStep();
            }          

            excelWorksheet.PageSetup.PrintTitleRows = "$1:$4";
            excelApp.Columns.AutoFit();
            excelApp.Visible = true;
            progressBar1.Minimum = 0;
            progressBar1.Visible = false;            
        }

        //botón guarda
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (TipoCam == 0)
            {
                try
                {
                    Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.19", "MES", "INICIO DE ACTUALIZACION DE FLETES MENSUALES " + LblTC.Text, "SIPGAB");
                    //DateTimeFormatInfo formatoFecha = CultureInfo.CurrentCulture.DateTimeFormat;
                    //string nombreMes = formatoFecha.GetMonthName(Reporte_costos_xtarima.date.Month);
                    //string nombreMes = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Reporte_costos_xtarima.date.Month);
                    //string nombreMes = obtenerNombreMesNumero(Convert.ToInt32(Convert.ToDateTime(Reporte_costos_xtarima.FI).Month.ToString()));
                    string nombreMes = obtenerNombreMesNumero(Reporte_costos_xtarima.date.Month).ToString();
                    //string fecha = nombreMes + "_" + Reporte_costos_xtarima.date.Year.ToString();
                    string fecha = nombreMes.ToUpper();
                    decimal valor = 0;
                    string month = obtenerNombreMesNumero(Convert.ToInt32(Convert.ToDateTime(Reporte_costos_xtarima.FI).Month.ToString()));
                    string anio = Convert.ToDateTime(Reporte_costos_xtarima.FI).Year.ToString();

                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "select " + fecha + " from tb_cat_costosprod2 where prod_clave ='05001ML24E'";
                    cmnd1.CommandText = "select costo from tb_cat_costosprod2 where prod_clave ='05001ML24E' and mes = '" + fecha + "' and año = '" + anio + "'";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        valor = reader1.GetSqlDecimal(0).Value;
                    }
                    reader1.Close();

                    if (valor > 0)
                    {
                        if (MessageBox.Show("Ya fueron registrados los costos para este mes, ¿deseas reemplazarlos?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            thisConnection.Close();
                            return;
                        }
                        else
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "delete tb_cat_costosprod2 where mes = '" + fecha + "' and año = '" + anio + "'";
                            reader1 = cmnd1.ExecuteReader();
                            reader1.Dispose();
                        }
                    }
                    //string query = "select * from tb_cat_costosprod";
                    string query = "select * from tb_cat_costosprod2 where mes = '" + month + "' and año ='" + anio + "'";
                    da = new SqlDataAdapter(query, thisConnection);
                    da.Fill(ds, "temporal");
                    temporal = ds.Tables["temporal"];

                    //query = "select lin_nombre, lin_clave  from tb_cat_linea order by lin_nombre";
                    //da = new SqlDataAdapter(query, thisConnection);
                    //da.Fill(ds, "linea");
                    //linea = ds.Tables["linea"];
                    makeyacasi();
                    progressBar1.Visible = true;
                    progressBar1.Value = 0;
                    progressBar1.Maximum = tem.Rows.Count;
                    progressBar1.Update();
                    //MessageBox.Show("Cargando Costos");
                    foreach (DataRow registro in tem.Select("coscajp > 0"))
                    {
                        int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                        progressBar1.CreateGraphics().DrawString("Actualizando costos . . .  " + percent.ToString() + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 60, progressBar1.Height / 2 - 7));
                        DataRow row = yacasi.NewRow();
                        row["producto"] = registro["producto"].ToString().Trim();
                        row["nproducto"] = registro["nproducto"].ToString().Trim();
                        row["linea"] = registro["CVEL"].ToString();
                        //row["nomlinea"] = registro["NOMLINEA"].ToString();
                        row["movimiento"] = registro["tipo"].ToString();
                        decimal cost = Convert.ToDecimal(registro["costarp"].ToString()) / Convert.ToDecimal(registro["totcajp"].ToString());
                        row["costo"] = System.Math.Round(cost, 2, MidpointRounding.ToEven);
                        //row["costo"] = Convert.ToDecimal(Convert.ToDecimal(registro["costarp"].ToString()) / Convert.ToDecimal(registro["totcajp"].ToString()));
                        //foreach (DataRow dr in linea.Select("lin_clave = '" + Convert.ToString(registro["cvel"].ToString().Trim()) + "'"))
                        //    row["nomlinea"] = Convert.ToString(dr["lin_nombre"].ToString().Trim());
                        yacasi.Rows.Add(row);
                        progressBar1.PerformStep();
                    }
                    progressBar1.Value = 0;
                    progressBar1.Maximum = tem.Rows.Count;
                    progressBar1.Update();

                    //MessageBox.Show("Actuliza los Costos");
                    DataView dv = new DataView();
                    dv = yacasi.DefaultView;
                    dv.Sort = "movimiento, producto, nproducto, linea, nomlinea";
                    yacasi = dv.ToTable();
                    bool find = false;
                    string qry = "";
                    foreach (DataRow regis in yacasi.Rows)
                    {
                        find = false;
                        int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                        progressBar1.CreateGraphics().DrawString("Actualizando costos . . .  " + percent.ToString() + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 60, progressBar1.Height / 2 - 7));
                        foreach (DataRow prod in temporal.Select("prod_clave ='" + Convert.ToString(regis["producto"].ToString()) + "' and movimiento = '" + Convert.ToString(regis["movimiento"].ToString()) + "' and mes = '" + month + "' and año = '" + anio + "'"))
                        {
                            find = true;
                            cmnd1 = thisConnection.CreateCommand();
                            ///cmnd1.CommandText = "update tb_cat_costosprod set " + fecha + " = '" + Convert.ToDecimal(regis["costo"].ToString()) + "' where prod_clave = '" + regis["producto"].ToString() + "'";
                            cmnd1.CommandText = "update tb_cat_costosprod2 set costo = '" + Convert.ToDecimal(regis["costo"].ToString()) + "' where " +
                                                "prod_clave = '" + Convert.ToString(regis["producto"].ToString()).Trim() + "' " +
                                                "and movimiento = '" + Convert.ToString(regis["movimiento"].ToString()) + "' and mes = '" + month + "' and año = '" + anio + "'";
                            reader1 = cmnd1.ExecuteReader();
                            reader1.Close();

                            qry = "update tb_cat_costosprod2 set costo = " + Convert.ToDecimal(regis["costo"].ToString()) + " where " +
                                        "prod_clave = " + Convert.ToString(regis["producto"].ToString()).Trim() + " " +
                                        "and movimiento = " + Convert.ToString(regis["movimiento"].ToString()) + " and mes = " + month + " and año = " + anio + "";
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "INSERT INTO tb_fletes_query(query) VALUES ('" + qry + "')";
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();
                            

                        }
                        if (find == false)
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            //System.Math.Round (n, 2, MidpointRounding.ToEven);
                            //cmnd1.CommandText = "insert into tb_cat_costosprod (prod_clave, prod_nombre, lin_clave, lin_nombre, " + fecha + ") values " +
                            //                  "('" + Convert.ToString(regis["producto"].ToString()) + "', '" + Convert.ToString(regis["nproducto"].ToString().Replace("'", " ")) + "', " +
                            //                  "'" + Convert.ToString(regis["linea"].ToString()) + "', '" + Convert.ToString(regis["nomlinea"].ToString()) + "', " +
                            //                  "'" + Convert.ToDecimal(regis["costo"].ToString()) + "')";     
                            decimal cost = System.Math.Round(Convert.ToDecimal(regis["costo"].ToString()), 2, MidpointRounding.ToEven);
                            cmnd1.CommandText = "insert into tb_cat_costosprod2 (prod_clave, lin_clave, mes, año, costo, movimiento) values ('" + Convert.ToString(regis["producto"].ToString()).Trim() + "', " +
                                                "'" + Convert.ToString(regis["linea"].ToString()).Trim() + "', '" + month + "', '" + anio + "', " +
                                                "'" + cost + "', '" + Convert.ToString(regis["movimiento"].ToString()) + "')";
                            reader1 = cmnd1.ExecuteReader();
                            reader1.Dispose();

                            qry = "insert into tb_cat_costosprod2 (prod_clave, lin_clave, mes, año, costo, movimiento) values (" + Convert.ToString(regis["producto"].ToString()).Trim() + ", " +
                                        "" + Convert.ToString(regis["linea"].ToString()).Trim() + ", " + month + ", " + anio + ", " +
                                        "" + cost + ", " + Convert.ToString(regis["movimiento"].ToString()) + ")";
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "INSERT INTO tb_fletes_query(query) VALUES ('" + qry + "')";
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();
                        }
                        progressBar1.PerformStep();
                    }
                    thisConnection.Close();
                    progressBar1.Visible = false;
                    progressBar1.Value = 0;
                    progressBar1.Update();
                    Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "F", "4.19", "MES", "FIN DE ACTUALIZACION DE FLETES MENSUALES " + LblTC.Text, "SIPGAB");
                    MessageBox.Show("Costos Registrados", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    lblUltimoCalculo.Text = ultima_fecha_calculada();
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
            else // guardo los Costos Por Semana
            {
                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.19", "SEM", "INICIO DE ACTUALIZACION DE FLETES SEMANALES " + LblTC.Text + ". Costos Registrados de la " + LblSem.Text, "SIPGAB");

                thisConnection.Open();
                string Cadena = "Select NoSem, Año, Lin_Clave, Prod_Clave FROM Tb_Costos_DetProd Where Año = '" + Año.ToString() + "' AND NoSem = '" + NoSem.ToString() + "'";
                SqlDataAdapter da = new SqlDataAdapter(Cadena, thisConnection);
                DataSet ds = new DataSet();
                da.Fill(ds,"Costos");
                DataTable Costos = ds.Tables["Costos"];
                Cadena = "Select * from tb_costos_FleteSem Where Anio = '" + Año.ToString() + "' AND Semana = '" + NoSem.ToString() + "'";
                da = new SqlDataAdapter(Cadena, thisConnection);
                ds = new DataSet();
                da.Fill(ds, "FleSem");
                DataTable FleSem = ds.Tables["FleSem"];
                //if (Costos.Rows.Count == 0)
                //{
                //    MessageBox.Show("Error: No se ha generado la Informacion de Ventas, Generar las Ventas de Esta " + LblSem.Text , "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                //    thisConnection.Close();
                //    return;
                //}
                if (FleSem.Rows.Count > 0)
                {
                    Cadena = "Delete from tb_costos_FleteSem Where Anio = '" + Año.ToString() + "' AND Semana = '" + NoSem.ToString() + "'";
                    SqlCommand cmd1 = new SqlCommand(Cadena, thisConnection);
                    cmd1.ExecuteNonQuery();
                }
                Cadena = "UPDATE Tb_Costos_DetProd SET FletesVta = 0" +
                         " WHERE AÑO = '" + Año.ToString() + "' AND NOSEM = '" + NoSem.ToString() + "'";
                SqlCommand cmd = new SqlCommand(Cadena, thisConnection);
                cmd.ExecuteNonQuery();
                int i = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    decimal imp = Convert.ToDecimal(dataGridView1.Rows[i].Cells["ImpFlete"].Value.ToString());
                    string cvepro = dataGridView1.Rows[i].Cells["prod_clave"].Value.ToString();
                    Cadena = "UPDATE Tb_Costos_DetProd SET FletesVta = FletesVta + '" + imp.ToString() + "'" +
                             " WHERE AÑO = '" +  Año.ToString() + "' AND NOSEM = '" + NoSem.ToString() + "' AND PROD_CLAVE =  '" + cvepro + "'";
                    cmd = new SqlCommand(Cadena, thisConnection);
                    cmd.ExecuteNonQuery();
                    if (dataGridView1.Rows[i].Cells["prod_clave"].Value.ToString().Trim().Length > 0)
                    {
                        //Cadena = "Insert into tb_costos_FleteSem(Semana, Anio, prod_clave, cajas, costo, tipo) Values ('" + NoSem.ToString() + "','" + Año.ToString() + "','" + dataGridView1.Rows[i].Cells["prod_clave"].Value.ToString() + "','" +
                        //         dataGridView1.Rows[i].Cells["ucajas"].Value.ToString() + "','" + dataGridView1.Rows[i].Cells["CPUCAJAS"].Value.ToString() + "','" + dataGridView1.Rows[i].Cells["moneda"].Value.ToString().Substring(0, 1) + "')";
                        Cadena = "Insert into tb_costos_FleteSem(Semana, Anio, prod_clave, cajas, costo, tipo, TipoCambio, Importe) Values ('" + NoSem.ToString() + "','" + Año.ToString() + "','" + dataGridView1.Rows[i].Cells["prod_clave"].Value.ToString() + "','" +
                                 dataGridView1.Rows[i].Cells["ucajas"].Value.ToString() + "','" + dataGridView1.Rows[i].Cells["CPUTARI"].Value.ToString() + "','" + dataGridView1.Rows[i].Cells["moneda"].Value.ToString().Substring(0, 1) + "','" +
                                 TipoCam.ToString() + "','" + dataGridView1.Rows[i].Cells["CTARIMAS"].Value.ToString() + "')";
                        cmd = new SqlCommand(Cadena, thisConnection);
                        cmd.ExecuteNonQuery();
                    }
                    i++;
                }
                thisConnection.Close();
                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "F", "4.19", "SEM", "FIN DE ACTUALIZACION DE FLETES SEMANALES " + LblTC.Text + ". Costos Registrados de la " + LblSem.Text, "SIPGAB");
                MessageBox.Show("Costos Registrados de la " + LblSem.Text, "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                lblUltimoCalculo.Text = ultima_fecha_calculada();
            }
        }

        private string obtenerNombreMesNumero(int numeroMes)
        {
            try
            {
                DateTimeFormatInfo formatoFecha = CultureInfo.CurrentCulture.DateTimeFormat;
                string nombreMes = formatoFecha.GetMonthName(numeroMes);
                return nombreMes.ToUpper();
            }
            catch
            {
                return "Desconocido";
            }
        } 

        private void Reporte_FormClosing(object sender, FormClosingEventArgs e)
        {
            //fech_ulti = "01/06/2022";
            if (fech_ulti != "")
            {
                if (Convert.ToDateTime(DateTime.Now.ToShortDateString()) == Convert.ToDateTime(fech_ulti))
                {
                    
                }
                else
                {
                    if (MessageBox.Show("¿Desea cerrar la ventana de los fletes aun no guarda el calculo de los fletes?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        //guardar registro de que cerro la ventana sin guardar
                        Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "Y", "4.19", "YES", "EL USUARIO SALIO SIN GUARDAR LA INFORMACION", "SIPGAB");
                    }
                    else
                    {
                        //
                        e.Cancel = true;
                    }
                }
            }
            count = 0;
            totalunicajas = 0; totalunitarimas = 0; totalcostoscajas = 0; totalcostostarimas = 0;
            tipomoneda = "";
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {            
            if (dataGridView1.Columns[e.ColumnIndex].Name == "prod_clave")
            {
                Producto_Factura prodfact = new Producto_Factura();
                prodfact.tabla = productos.Clone();
                prodfact.detalle = detalle_flete.Clone();
                prod = dataGridView1.Rows[e.RowIndex].Cells["prod_clave"].Value.ToString().Trim();
                foreach (DataRow row in detalle_factura.Select("producto = '" + prod + "'"))
                    prodfact.tabla.ImportRow(row);
                foreach (DataRow row in  detalle_flete.Select("producto = '" + prod + "'"))
                    prodfact.detalle.ImportRow(row);

                if (prodfact.tabla.Rows.Count > 0)
                    prodfact.ShowDialog();                
            }
        }

        public string ultima_fecha_calculada()
        {
            string fch = "";
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            if (tipo_reporte == "M")
            {
                fch = "Fecha ultima de calculo mensual de fletes el ";
                cmnd1.CommandText = "SELECT TOP 1 fecha FROM tb_registro_movimientos WHERE tipo_mov = 'F' AND folio = 'MES' ORDER BY fecha DESC ";
            }
            else
            {
                fch = "Fecha ultima de calculo semanal de fletes el ";
                cmnd1.CommandText = "SELECT TOP 1 fecha FROM tb_registro_movimientos WHERE tipo_mov = 'F' AND folio = 'SEM' ORDER BY fecha DESC ";
            }
                
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                fch += Convert.ToDateTime(reader1["fecha"].ToString().Trim()).ToLongDateString();
                fech_ulti = Convert.ToDateTime(reader1["fecha"].ToString().Trim()).ToShortDateString();
            }
            else
            {
                fch += "No existe registro";
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();
            return fch;
        }

        public void manda_correo_empaques_fletes()
        {
            string body = "<table border='1'>";
            body += "<tr><th>Conse</th><th>Prod Clave</th><th>Cajas</th><th>Costo flete</th></tr>";
            int i = 0;
            foreach (DataGridViewRow rw in dataGridView1.Rows)
            {
                if (rw.Cells["producto"].Value.ToString() == "" || rw.Cells["producto"].Value.ToString().Contains("TOTALES"))
                    continue;
                else
                {
                    i++;
                    body += "<tr><td>" + i.ToString() + "</td><td>" + rw.Cells["prod_clave"].Value.ToString() + "</td><td>" + rw.Cells["ucajas"].Value.ToString() + "</td><td>" + rw.Cells["cputari"].Value.ToString() + "</td></tr>";
                }                    
            }
            body += "</table>";

            try
            {
                GraphMail.Enviar("jgalvan@mrlucky.com.mx", "aescamilla@mrlucky.com.mx", "Calculo de fletes", body);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Mensaje no enviado: " + ex.Message);
            }
        }

        private void BtnDet_Click(object sender, EventArgs e)
        {
            DgDet.DataSource = detalle_flete;
            ReporteExcel(DgDet);
        }
    }
}
