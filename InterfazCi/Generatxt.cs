using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using MyExcel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using Microsoft.Win32;
using System.Configuration;
//using MyExcel = Microsoft.Office.Interop.Excel;



namespace InterfazCi
{
    public partial class Generatxt : Form
    {

       // SDKCONTPAQNGLib.TSdkSesion sesion = new SDKCONTPAQNGLib.TSdkSesion();
        public string Cadenaconexion = "";
        public string Archivo = "";
        public DataTable DatosReporte = null;

        public Generatxt()
        {
            InitializeComponent();
            if (Properties.Settings.Default.server != "")
            {
                Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";
                Archivo = Properties.Settings.Default.archivo;
                if (Archivo == "C:\\Users\\TOSHIBA\\Documents\\clientes\\binstala\\InterfazCi\\Compras.xls" || Archivo =="")
                    botonExcel1.mSetNombre(System.IO.Directory.GetCurrentDirectory() + "\\ArchivoExportar.txt");
            }
        }

        private void Generatxt_Load(object sender, EventArgs e)
        {
            botonExcel1.mSetEtiqueta("Archivo Exportar");
            botonExcel1.tipo = 1;
        }

        private void Generatxt_Shown(object sender, EventArgs e)
        {
            if (Cadenaconexion != "")
            {
                ciCompanyList11.Populate(Cadenaconexion);
            }
            else
            {
                this.Visible = false;
                Form4 x = new Form4();
                x.asignaformGeneratxt(this);
                x.Show();
            }
            this.Text = " Genera txt desde Contabilidad " + this.ProductVersion;
           // if (Archivo != "")
                botonExcel1.mSetNombre(System.IO.Directory.GetCurrentDirectory() + "\\ArchivoExportar.txt");


        }
        public void mllenarcomboempresas()
        {
            ciCompanyList11.Populate(Cadenaconexion);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Archivo = botonExcel1.mRegresarNombre();

            Properties.Settings.Default.Save();

            DateTime lfecha = dateTimePicker1.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');

            sfecha1 = lfecha.Day.ToString().PadLeft(2, '0') + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Year.ToString();

            DateTime lfecha2 = dateTimePicker2.Value;
            string sfecha2 = lfecha2.Year.ToString() + lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Day.ToString().PadLeft(2, '0');


            sfecha2 = lfecha2.Day.ToString().PadLeft(2, '0') + lfecha2.Month.ToString().PadLeft(2, '0') + lfecha2.Year.ToString();


            string sql = " select '001' as empresa, REPLACE(STR(s.Codigo, 4), SPACE(1), '0')  as sucursal, 'CONTPAQ ' as usuario " +
            " ,replace(Convert(varchar(10),CONVERT(date,p.fecha,106),103),'/','') as fecha " +
            " ,substring(c.Codigo,1,4) as mayor " +
            " ,substring(c.Codigo,5,2) as nivel1 " +
            " ,substring(c.Codigo,7,2) as nivel2 " +
            " ,substring(c.Codigo,9,2) as nivel3 " +
            " ,substring(c.Codigo,11,2) as nivel4 " +
            " ,substring(c.Codigo,13,2) as nivel5 " +
            " ,'004 'as region " +
            " ,REPLACE(STR(s.Codigo, 4), SPACE(1), '0')  as sucursal2 " +
           // " ,REPLACE(STR(tp.Codigo, 3), SPACE(1), '0')  as region " +
           ",mp.Concepto as concpetomovto" +
            " ,'0001' as centrocosto " +
            " ,'' as NoAuxiliar " +
            " ,replace(Convert(varchar(10),CONVERT(date,p.fecha,106),103),'/','') as fecha2 " +
            " ,case when mp.importe != 0 then '01' else '02' end as moneda " +
            " ,case when mp.importe != 0 then mp.importe else mp.importe end as monto " +
            " ,case when mp.TipoMovto = 1 then 'D' else 'C' end as naturaleza " +
            " ,p.Concepto as conceptopoliza " +
            " ,mp.Referencia as referenciamovto " +
            " ,mp.Concepto as concpetomovto1 " +
            " from Polizas p " +
            " join MovimientosPoliza mp on p.id = mp.idpoliza " +
            " join Cuentas c on c.id = mp.IdCuenta " +
            " join TiposPolizas tp on tp.Id = p.TipoPol " +
            " join SegmentosNegocio s on s.Id = mp.IdSegNeg " +
            " where replace(Convert(varchar(10),CONVERT(date,p.fecha,106),103),'/','') >= '" + sfecha1 + "'" +
            " and replace(Convert(varchar(10),CONVERT(date,p.fecha,106),103),'/','') <= '" + sfecha2 + "'" +
            " order by p.Folio ";

            //MessageBox.Show("uno");
            mTraerInformacionContabilidad(sql, ciCompanyList11.aliasbdd);
            //MessageBox.Show("dos");
            mMostrarInfo(ciCompanyList11.aliasbdd);
            MessageBox.Show("Proceso Terminado");




        }
        public void mTraerInformacionContabilidad(string lquery, string mEmpresa)
        {
            SqlConnection _conexion1 = new SqlConnection();
            //            rutadestino = "c:\\compacw\\empresas\\adtala2";
            string rutadestino = mEmpresa;

            string sempresa = rutadestino.Substring(rutadestino.LastIndexOf("\\") + 1);

            string server = Properties.Settings.Default.server;
            string user = Properties.Settings.Default.user;
            string pwd = Properties.Settings.Default.password;
            //sempresa = GetSettingValueFromAppConfigForDLL("empresa");
            //string lruta3 = obc.ToString();
            string lruta4 = @rutadestino;
            _conexion1 = new SqlConnection();
            string Cadenaconexion1 = "data source =" + server + ";initial catalog = " + sempresa + ";user id = " + user + "; password = " + pwd + ";";
            _conexion1.ConnectionString = Cadenaconexion1;
            _conexion1.Open();




            DataSet ds = new DataSet();

            string lsql = lquery.ToString();
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter(lsql, _conexion1);



            //mySqlDataAdapter.SelectCommand.Connection = _conexion1;

            //mySqlDataAdapter.SelectCommand.Connection = _conexion1;
            //mySqlDataAdapter.SelectCommand.CommandText = lsql;

            mySqlDataAdapter.Fill(ds);

            DatosReporte = ds.Tables[0];
            //if (ds.Tables.Count > 1)
                //DatosDetalle = ds.Tables[1];
            _conexion1.Close();

        }


        public MyExcel.Workbook mIniciarExcel()
        {
            MyExcel.Application excelApp = new MyExcel.Application();
            excelApp.Visible = true;
            MyExcel.Workbook newWorkbook = excelApp.Workbooks.Add();
            newWorkbook.Worksheets.Add();
            return newWorkbook;

        }

        public void mMostrarInfo(string mEmpresa)//, string lfechai, string lfechaf)
        {
            
            int lrenglon = 1;
            int lrengloninicial = 1;

            MyExcel.Worksheet sheet = null;

            if (checkBox2.Checked == true)
            {
                MyExcel.Workbook newWorkbook = mIniciarExcel();
                 sheet = (MyExcel.Worksheet)newWorkbook.Sheets[1];
            }

            string lconcepto = "";
            string lcliente = "";
            int lmismoconcepto = 0;
            
            decimal dos, tres;
            int lcolumna;


            dataGridView1.DataSource = null;
            dataGridView1.DataSource = DatosReporte;

            //before your loop
            var csv = new StringBuilder();

            //in your loop
            

            //after your loop
            //var newLine = string.Format("{0},{1}", first, second);

            

            var newLine1 = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{12},{13},{14},{15},{16},{17},{18}",  
                "empresa", "sucursal", "usuario", "fecha", "mayor", 
                "nivel1", "nivel2", "nivel3", "nivel4", "nivel5",
                "region", "sucursal", "centrocosto", "NoAuxiliar", 
                "fecha2", "moneda", "monto", "naturaleza", "conceptopoliza");

            if (checkBox2.Checked == true)
            {
                (sheet.Cells[lrenglon, 1] as MyExcel.Range).Value = "empresa";
                (sheet.Cells[lrenglon, 2] as MyExcel.Range).Value = "sucursal";
                (sheet.Cells[lrenglon, 3] as MyExcel.Range).Value = "usuario";
                (sheet.Cells[lrenglon, 4] as MyExcel.Range).Value = "fecha";
                (sheet.Cells[lrenglon, 5] as MyExcel.Range).Value = "mayor";

                (sheet.Cells[lrenglon, 6] as MyExcel.Range).Value = "nivel1";
                (sheet.Cells[lrenglon, 7] as MyExcel.Range).Value = "nivel2";
                (sheet.Cells[lrenglon, 8] as MyExcel.Range).Value = "nivel3";
                (sheet.Cells[lrenglon, 9] as MyExcel.Range).Value = "nivel4";
                (sheet.Cells[lrenglon, 10] as MyExcel.Range).Value = "nivel5";

                //(sheet.Cells[lrenglon, 11] as MyExcel.Range).Value = "nivel6";
                (sheet.Cells[lrenglon, 11] as MyExcel.Range).Value = "region";
                (sheet.Cells[lrenglon, 12] as MyExcel.Range).Value = "sucursal";
                (sheet.Cells[lrenglon, 13] as MyExcel.Range).Value = "centrocosto";
                (sheet.Cells[lrenglon, 14] as MyExcel.Range).Value = "NoAuxiliar";

                (sheet.Cells[lrenglon, 15] as MyExcel.Range).Value = "fecha2";
                (sheet.Cells[lrenglon, 16] as MyExcel.Range).Value = "moneda";
                (sheet.Cells[lrenglon, 17] as MyExcel.Range).Value = "monto";
                (sheet.Cells[lrenglon, 18] as MyExcel.Range).Value = "naturaleza";
                (sheet.Cells[lrenglon, 19] as MyExcel.Range).Value = "conceptopoliza";



                lrenglon += 1;
            }

            //sheet.Cells[1,1].

            //sheet.Cells[1, 1].value = newLine1;

            //sheet.Cells[lrenglon, lcolumna++].value = newLine1;

            
            //sheet.Cells[lrenglon, lcolumna++].value = row[0].ToString().Trim(); //Fecha

            string nombre = botonExcel1.mRegresarNombre();
           // csv.AppendLine(newLine1);
            //File.WriteAllText(nombre, newLine1.ToString());
            
            
            //File.WriteAllText(filePath, csv.ToString());


            foreach (DataRow row in DatosReporte.Rows)
            {
                //Fecha	# pedidos	cliente	importe	pendiente de facturar	# de factura	cliente	importe	Impuesto	Retención	Total


              /*  var first = row[0].ToString().Trim();
                var second = row[0].ToString().Trim();
                //Suggestion made by KyleMit
                var newLine = string.Format("{0},{1}", first, second);
                csv.AppendLine(newLine);*/

                decimal uno = decimal.Parse(row[17].ToString().Trim());

                //uno = 345.33M;
                string x = uno.ToString("0.00").PadRight(12,' ');

                x = uno.ToString("0.00");



                string ssucursal = row[1].ToString().Trim();
                newLine1 = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18} ",
                row[0].ToString().Trim(), row[1].ToString().Trim(), row[2].ToString().Trim(), row[3].ToString().Trim(), row[4].ToString().Trim(),
                row[5].ToString().Trim(), row[6].ToString().Trim(), row[7].ToString().Trim(), row[8].ToString().Trim(), row[9].ToString().Trim(),
                //row[10].ToString().Trim(), row[12].ToString().Trim(),
                "004" , ssucursal,
                row[13].ToString().Trim(), row[14].ToString(),
                row[15].ToString().Trim(), row[16].ToString().Trim(), x, row[18].ToString().Trim(), row[19].ToString().Trim()
                );
                if (checkBox1.Checked == true)
                {
                    csv.AppendLine(newLine1);
                }


                

                //var cellValue = (string)(sheet.Cells[10, 2] as MyExcel.Range).Value;

                if (checkBox2.Checked == true)
                {

                    (sheet.Cells[lrenglon, 1] as MyExcel.Range).Value = "'" + row[0].ToString().Trim(); //Fecha

                    (sheet.Cells[lrenglon, 2] as MyExcel.Range).Value = "'" + row[1].ToString().Trim(); // "sucursal";
                    (sheet.Cells[lrenglon, 3] as MyExcel.Range).Value = "'" + row[2].ToString().Trim(); //Fecha"usuario";
                    (sheet.Cells[lrenglon, 4] as MyExcel.Range).Value = "'" + row[3].ToString().Trim(); //Fecha"fecha";
                    (sheet.Cells[lrenglon, 5] as MyExcel.Range).Value = "'" + row[4].ToString().Trim(); //Fecha"mayor";

                    (sheet.Cells[lrenglon, 6] as MyExcel.Range).Value = "'" + row[5].ToString().Trim(); //Fecha"nivel1";
                    (sheet.Cells[lrenglon, 7] as MyExcel.Range).Value = "'" + row[6].ToString().Trim(); //Fecha"nivel2";
                    (sheet.Cells[lrenglon, 8] as MyExcel.Range).Value = "'" + row[7].ToString().Trim(); //Fecha"nivel3";
                    (sheet.Cells[lrenglon, 9] as MyExcel.Range).Value = "'" + row[8].ToString().Trim(); //Fecha"nivel4";
                    (sheet.Cells[lrenglon, 10] as MyExcel.Range).Value = "'" + row[9].ToString().Trim(); //nivel5

                    //(sheet.Cells[lrenglon, 11] as MyExcel.Range).Value = "'" + row[10].ToString().Trim(); //Fecha"nivel6";

                    (sheet.Cells[lrenglon, 11] as MyExcel.Range).Value = "'004"; //Fecha"sucursal"; region 


                    (sheet.Cells[lrenglon, 12] as MyExcel.Range).Value = "'" + row[1].ToString().Trim(); // "sucursal";

                    (sheet.Cells[lrenglon, 13] as MyExcel.Range).Value = "'" + row[13].ToString().Trim(); // "centro de costos";


                    (sheet.Cells[lrenglon, 14] as MyExcel.Range).Value = "'" + row[14].ToString().Trim(); //Fecha"centrocosto";
                    (sheet.Cells[lrenglon, 15] as MyExcel.Range).Value = "'" + row[15].ToString().Trim(); //Fecha"NoAuxiliar";

                    (sheet.Cells[lrenglon, 16] as MyExcel.Range).Value = "'" + row[16].ToString().Trim(); //Fecha"fecha2";
                    (sheet.Cells[lrenglon, 17] as MyExcel.Range).Value = "'" + x; //Fecha"monto";
                    (sheet.Cells[lrenglon, 18] as MyExcel.Range).Value = "'" + row[18].ToString().Trim(); //Fecha"moneda";
                    (sheet.Cells[lrenglon, 19] as MyExcel.Range).Value = "'" + row[19].ToString().Trim(); //Fecha"naturaleza";

                    // (sheet.Cells[lrenglon, 20] as MyExcel.Range).Value = "'" + row[19].ToString().Trim(); //Fecha"conceptopoliza";


                    //sheet.Cells[lrenglon, lcolumna++].Value = row[0].ToString().Trim(); //Fecha
                    /*       sheet.Cells[lrenglon, lcolumna++].value = row[1].ToString().Trim(); //serie pedidos
                           sheet.Cells[lrenglon, lcolumna++].value = row[2].ToString().Trim(); //#pedidos
                           sheet.Cells[lrenglon, lcolumna++].value = row[3].ToString().Trim(); //cliente
                           sheet.Cells[lrenglon, lcolumna++].value = row[4].ToString().Trim(); //importe
                           sheet.Cells[lrenglon, lcolumna++].value = row[5].ToString().Trim(); //pendiente de facturar

                           sheet.Cells[lrenglon, lcolumna++].value = "'" + row[13].ToString().Trim(); //fecha de facturacion

                           sheet.Cells[lrenglon, lcolumna++].value = row[6].ToString().Trim(); //serie factura
                           sheet.Cells[lrenglon, lcolumna++].value = row[7].ToString().Trim(); //# de factura
                           if (row[7].ToString().Trim() == "")
                               sheet.Cells[lrenglon, lcolumna++].value = ""; // cliente    
                           else
                               sheet.Cells[lrenglon, lcolumna++].value = row[8].ToString().Trim(); // cliente
                           sheet.Cells[lrenglon, lcolumna++].value = row[9].ToString().Trim(); // importe
                           sheet.Cells[lrenglon, lcolumna++].value = row[10].ToString().Trim(); // impuesto
                           sheet.Cells[lrenglon, lcolumna++].value = row[11].ToString().Trim(); // retencion
                           sheet.Cells[lrenglon, lcolumna++].value = row[12].ToString().Trim(); // total



                           sheet.get_Range("E" + lrenglon.ToString(), "F" + lrenglon.ToString()).Style = "Currency";
                           sheet.get_Range("K" + lrenglon.ToString(), "N" + lrenglon.ToString()).Style = "Currency";
            * */
                    lrenglon++;
                }


            }
           // sheet.Cells.EntireColumn.AutoFit();


            //MessageBox.Show("2.5");
            if (checkBox1.Checked == true)
            {
                File.WriteAllText(nombre, csv.ToString());
            }
            //MessageBox.Show("2.8");

            return;



        }
    }
}
