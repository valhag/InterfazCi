using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;

namespace InterfazCi
{
    public partial class Romero : Base 
    {

        SDKCONTPAQNGLib.TSdkSesion sesion = new SDKCONTPAQNGLib.TSdkSesion();        

        public Romero()
        {
            InitializeComponent();
        }

        private void Romero_Load(object sender, EventArgs e)
        {
            botonExcel1.mSetTipo(0);
            botonExcel1.mSetNombre(System.IO.Directory.GetCurrentDirectory() + "\\Archivo Origen.xlsx");
            //MessageBox.Show( base.Controls["TabControl1"].Name);
            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(botonExcel1);

            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(button2);

            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(progressBar1);

            botonExcel1.Top = 70;
            botonExcel1.Left = 0;

            progressBar1.Top = 175;
            progressBar1.Left = 0;


            button2.Top = 125;
            button2.Left = 25;
            button2.Text = "Enviar Informacion";

            this.Text = " Interfaz Romero " + this.ProductVersion;
        }

        private string mLlenarPolizas()
        {
            string aNombreArchivo = botonExcel1.mRegresarNombre();
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + aNombreArchivo + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");

            conn.Open();

            DataTable tables = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);  


            System.Data.OleDb.OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
            //EMPRESA	ANIO	PERIODO	TIPO POLIZA	NUMERO POLIZA	FECHA	MODULO	ESTATUS	LINEA	CTA	SCTA	SSCTA	OBRA	CC	DESCRIPCION	REFERENCIA	NATURALEZA	IMPORTE

            DataRow table = tables.Rows[0];
            string name = table["TABLE_NAME"].ToString();

            name = name.Replace("'", "");

            //name += "D7:Z";
            cmd.CommandText = "SELECT * FROM [" + name +"]";

            

            //comm.CommandText = "Select * from [" + "Sheet1" + "$A3:B6]";

            //cmd.CommandText = "SELECT * FROM [Sheet1$]";
            /*cmd.ExecuteNonQuery();
            dr = cmd.ExecuteReader();*/

            long xxx;
            xxx = 1000000;

         


            System.Data.OleDb.OleDbDataReader dr;
            dr = cmd.ExecuteReader();

/*
            DataTable dt = new DataTable();
            dt.Load(dr1);
            IEnumerable<DataRow> newRows = dt.AsEnumerable().Skip(7);
            DataTable dt2 = newRows.CopyToDataTable();

            DataTableReader dr = dt2.CreateDataReader();
            */

            Boolean noseguir = false;
            _RegPolizas.Clear();
            long ifolio = 0;
            long lfolio = -1;
            int renglonblanco = 1;
            if (dr.HasRows)
                while (noseguir == false)
                {
                    renglonblanco = 1;
                    dr.Read();

                    try
                    {
                        ifolio = long.Parse(dr["DocumentNo"].ToString());
                        //ifolio = int.Parse(dr[5].ToString());
                    }
                    catch (Exception e)
                    {
                        string test = "";
                        try
                        {
                            test = dr["DocumentNo"].ToString();
                        }
                        catch (Exception ee)
                        {
                            _RegPolizas.Add(_Poliza);
                            noseguir = true;
                            break;
                        }
                        if (dr["DocumentNo"].ToString() == "")
                        {
                            if (_Poliza._RegMovtos.Count > 0)
                                _RegPolizas.Add(_Poliza);
                            renglonblanco = 0;
                        }
                    }
                    int zzzzzz = 0;
                    if (renglonblanco != 0)
                    {
                        if (ifolio != lfolio)
                        {
                            if (ifolio ==1916000150)
                                 zzzzzz = 0;
                            if (lfolio != -1)
                                _RegPolizas.Add(_Poliza);
                            _Poliza = null;
                            _Poliza = new Poliza();
                            _Poliza.Folio = ifolio;
                            _Poliza.Concepto = "";
                            //_Poliza.FechaAlta = DateTime.Parse(dr["FECHA POLIZA"].ToString());
                            string lfecha = dr["Doc# Date"].ToString();
                            int primerdiagonal = lfecha.IndexOf('.', 0);
                            int segundadiagonal = lfecha.IndexOf('.', primerdiagonal + 1);



                            string ldia = lfecha.Substring(0, primerdiagonal);

                            string lanio = lfecha.Substring(segundadiagonal + 1);

                          //  lanio = "2019";

                            string lmes = lfecha.Substring(primerdiagonal + 1, segundadiagonal - (primerdiagonal + 1));
                            _Poliza.FechaAlta = DateTime.Parse(ldia.ToString() + "/" + lmes.ToString() + "/" + lanio.ToString());



                            Guid g;

                            g = Guid.NewGuid();

                            //_Poliza.Guid = g.ToString();

                            //_Poliza.FechaAlta = DateTime.Parse(dr[6].ToString());
                            lfolio = _Poliza.Folio;
                            //_Poliza.TipoPol = int.Parse(dr["TIPO POLIZA"].ToString());
                            _Poliza.TipoPol = 1;
                            switch (dr["Tipo Poliza"].ToString()){
                                case "Diario": _Poliza.TipoPol = 3; break;
                                case "Egresos": _Poliza.TipoPol = 2; break;
                                case "Ingresos": _Poliza.TipoPol = 1; break;

                        }
                            _Poliza.Concepto = dr["Document Header Text"].ToString().Trim();
                        }

                        //_Poliza.sMensaje = dr["UUID"].ToString().Trim();

                        //if (_Poliza.sMensaje != "")
                          //  _Poliza.sMensaje = dr["UUID"].ToString().Trim();
                        MovPoliza lRegmovto = new MovPoliza();
                        //CTA	SCTA	SSCTA

                        //lRegmovto.cuenta = dr["CTA"].ToString().Trim().PadLeft(4, '0') + dr["SCTA"].ToString().Trim().PadLeft(4, '0') + dr["SSCTA"].ToString().Trim().PadLeft(4, '0');


                        
                            

                        //lRegmovto.concepto = dr["DESCRIPCION POLIZA"].ToString();
                        //lRegmovto.referencia = dr["REFERENCIA"].ToString();

                        //lRegmovto.sn = dr["OBRA (Segmento de Negocio)"].ToString();
                        //lRegmovto.cuenta = dr[10].ToString().Trim().PadLeft(4, '0') + dr[11].ToString().Trim().PadLeft(4, '0') + dr[12].ToString().Trim().PadLeft(4, '0');
                        //NATURALEZA
                        //string naturaleza = dr["NATURALEZA"].ToString();
                        //string naturaleza = dr[19].ToString();


                        if (decimal.Parse(dr["Amount in local cur#"].ToString()) < 0)
                        {
                            lRegmovto.credito = decimal.Parse(dr["Amount in local cur#"].ToString()) * - 1;
                        }
                        else
                            lRegmovto.debito = decimal.Parse(dr["Amount in local cur#"].ToString());

                        lRegmovto.cuenta = dr["Account"].ToString().Trim();
                        string lcuenta = mBuscaCuenta(dr["Account"].ToString().Trim(), decimal.Parse(dr["Amount in local cur#"].ToString()));
                        lRegmovto.cuenta = lcuenta;

                        lRegmovto.referencia = dr["DocumentNo"].ToString();
                        lRegmovto.uuid = dr["UUID"].ToString();
                        if (lRegmovto.uuid != "")
                            lRegmovto.uuid = dr["UUID"].ToString();
                            
                        /*
                        lRegmovto.credito = decimal.Parse(dr["Amount in local cur#"].ToString());
                        lRegmovto.debito = 0;
                        if (lRegmovto.credito < 0)
                        {
                            lRegmovto.credito = 0;
                            //lRegmovto.debito = decimal.Parse(dr["IMPORTE"].ToString());
                            lRegmovto.debito = decimal.Parse(dr["Amount in local cur#"].ToString());
                        }
                        */


                        _Poliza._RegMovtos.Add(lRegmovto);
                        _Poliza.sMensaje = "";
                    }
                }



            //IEnumerable<Poliza> query = _RegPolizas.Where(pol => pol.TipoPol == 1);


            int count = _RegPolizas.Count(pol => pol.TipoPol == 3);

            return "";


        }


        private string mBuscaCuenta(string acuenta, decimal monto)
        {
            string cuenta = "";
            SqlConnection _conexion1 = new SqlConnection();
            //            rutadestino = "c:\\compacw\\empresas\\adtala2";
            string rutadestino = ciCompanyList11.aliasbdd;

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

            SqlCommand lcom = new SqlCommand();
            lcom.CommandText = "select * from cuentas where NomIdioma = @Cuenta";
            SqlParameter lparam = new SqlParameter();

            lparam.ParameterName = "@Cuenta";
            lparam.Value = acuenta;

            lcom.Parameters.Add(lparam);
            lcom.Connection = _conexion1 ;
            
            SqlDataReader ldr;
            ldr = lcom.ExecuteReader();

            if (ldr.HasRows)
            {
                ldr.Read();
                //string cuentax = ldr["Codigo"].ToString();
                cuenta = ldr["Codigo"].ToString();
            }
            else
                if (monto < 0) // cuenta debito
                    cuenta = "20101000";
                else
                    cuenta = "60184000";


        


            _conexion1.Close();

            return cuenta;
        }

        private void mGrabarPolizas(int incluir = 0)
        {

            SDKCONTPAQNGLib.TSdkEmpresa empresa = new SDKCONTPAQNGLib.TSdkEmpresa();
            SDKCONTPAQNGLib.TSdkPoliza poliza = new SDKCONTPAQNGLib.TSdkPoliza();
            SDKCONTPAQNGLib.TSdkMovimientoPoliza movimientosPoliza = new SDKCONTPAQNGLib.TSdkMovimientoPoliza();

            SDKCONTPAQNGLib.TSdkCuenta Cuenta = new SDKCONTPAQNGLib.TSdkCuenta();
            sesion = null;
            empresa = null;
            poliza = null;
            movimientosPoliza = null;
            Cuenta = null;

            sesion = new SDKCONTPAQNGLib.TSdkSesion();
            empresa = new SDKCONTPAQNGLib.TSdkEmpresa();
            poliza = new SDKCONTPAQNGLib.TSdkPoliza();
            movimientosPoliza = new SDKCONTPAQNGLib.TSdkMovimientoPoliza();

            Cuenta = new SDKCONTPAQNGLib.TSdkCuenta();
            if (sesion.conexionActiva == 0)
                sesion.iniciaConexion();

            if (sesion.conexionActiva == 1 && sesion.ingresoUsuario == 0)
                sesion.firmaUsuario();

            int uno = 0;
            string lempresa = ciCompanyList11.aliasbdd;
            if (sesion.conexionActiva == 1 && sesion.ingresoUsuario == 1)
            {
                uno = sesion.abreEmpresa(lempresa);

            }


            //int i = sesion.abreEmpresa("ctEmpresatest");
            if (sesion.conexionActiva == 0)
            {
                sesion.cierraEmpresa();
                MessageBox.Show("No se puede iniciar Contabilidad");
                return; 
            }

            empresa.setSesion(sesion);
            int lcuantos = 0;


            int cuantos = _RegPolizas.Count;

            decimal avance  =  cuantos/100;

            progressBar1.Value = 0;

            int polizaactual = 1;

            int lgenerados = 1;

            string dir = System.IO.Directory.GetCurrentDirectory();
            string now = System.DateTime.Now.ToString().Substring(0, 19);
            string file = "ErrorLog_" + now.Substring(0, 2) + now.Substring(3, 2) + now.Substring(6, 4) + "_" + now.Substring(11, 2) + now.Substring(14, 2) + now.Substring(17, 2) + ".txt";
            file = dir + "\\" + file;

            int conerror = 0;

            foreach (Poliza x in _RegPolizas)
            {

                poliza = new SDKCONTPAQNGLib.TSdkPoliza();
                poliza.setSesion(sesion);


                poliza.iniciarInfo();

                poliza.Impresa = 0;
                poliza.Diario = 0;
                //poliza.Concepto = "1";
                poliza.CodigoDiario = "";
                poliza.Concepto = x.Concepto;
                //poliza. = x.Referencia;
                poliza.Clase = SDKCONTPAQNGLib.ECLASEPOLIZA.CLASE_AFECTAR;

                poliza.Fecha = x.FechaAlta;
                //poliza.Fecha = Convert.ToDateTime("01/11/2014");

                poliza.Concepto = x.Concepto;
               //-- poliza.Guid

               // poliza.Guid = x.Guid;


                poliza.Tipo = SDKCONTPAQNGLib.ETIPOPOLIZA.TIPO_EGRESOS;
                poliza.Tipo = (SDKCONTPAQNGLib.ETIPOPOLIZA)x.TipoPol;
                poliza.Numero = poliza.getUltimoNumero(x.FechaAlta.Year, x.FechaAlta.Month, poliza.Tipo);


                //poliza.Concepto = x.Folio.ToString();
                poliza.Concepto = x.Concepto;
                //poliza.Tipo = (SDKCONTPAQNGLib.ETIPOPOLIZA.

                poliza.SistOrigen = SDKCONTPAQNGLib.ESISTORIGEN.ORIG_CONTPAQNG;

                //int idpoliza = poliza.crea();
                int lmovto = 1;

                //test
                /*Cuenta.setSesion(sesion);
                Cuenta.iniciarInfo();
                Cuenta.Codigo = "110400010042";
                Cuenta.Nombre = "Hector";
                Cuenta.CodigoCuentaAcumula = "110400010000";
                Cuenta.SistOrigen = SDKCONTPAQNGLib.ESISTORIGEN.ORIG_CONTPAQNG;
                Cuenta.Moneda = 1;
                Cuenta.Tipo = SDKCONTPAQNGLib.ECUENTATIPO.CUENTA_ACTIVODEUDORA;
                
                    Cuenta.FechaAlta = DateTime.Today;
                int zzz=0;
                Cuenta.CtaMayor = SDKCONTPAQNGLib.ECUENTADEMAYOR.CUENTA_CUENTADEMAYORNO;
                Cuenta.crea(zzz);
                */



                if (polizaactual >= avance)
                {
                    if (progressBar1.Value < 100)
                    {
                        progressBar1.Value += 1;
                        polizaactual = 1;
                    }
                }
                else
                    polizaactual++;
                
                foreach (MovPoliza y in x._RegMovtos)
                {
                    Cuenta.setSesion(sesion);

                    //string lcuenta = mBuscaCuenta(y.cuenta, y.credito);
                    

                    //Cuenta.buscaPorCodigo(y.cuenta.Trim());

                    Cuenta.buscaPorCodigo(y.cuenta);

                    //int zz = Cuenta.buscaPorId(15);
                    movimientosPoliza.iniciarInfo();
                    movimientosPoliza.setSesion(sesion);

                    Guid g;
                    g = Guid.NewGuid();
                   // movimientosPoliza.Guid = g.ToString();




                    movimientosPoliza.setSdkCuenta(Cuenta);
                    movimientosPoliza.CodigoCuenta = Cuenta.Codigo;
                    //movimientosPoliza.CodigoCuenta = y.cuenta.Trim();

                    movimientosPoliza.NumMovto = lmovto;
                    decimal limporte;
                    limporte = y.debito;
                    movimientosPoliza.TipoMovto = SDKCONTPAQNGLib.ETIPOIMPORTEMOVPOLIZA.MOVPOLIZA_CARGO;
                    if (y.credito != 0)
                    {
                        limporte = y.credito;
                        movimientosPoliza.TipoMovto = SDKCONTPAQNGLib.ETIPOIMPORTEMOVPOLIZA.MOVPOLIZA_ABONO;
                    }
                    movimientosPoliza.Importe = limporte;
                    movimientosPoliza.ImporteME = 0;
                    
                    //movimientosPoliza.Concepto = y.concepto;

                    movimientosPoliza.Concepto = x.Concepto;
                    movimientosPoliza.Referencia = y.referencia;
                    

                    int movAgregado = poliza.agregaMovimiento(movimientosPoliza);
                    //int mov1 = poliza.creaMovimiento(movimientosPoliza);
                    lmovto++;

                }
                int idpoliza = poliza.crea();

                if (idpoliza == 0)
                {
                    
                    System.IO.File.AppendAllText(file, "La poliza con DocumentNo " + movimientosPoliza.Referencia + " tiene un error revise sus datos y vuelva a enviarlo" + Environment.NewLine);
                    //MessageBox.Show( "La poliza con DocumentNo " + movimientosPoliza.Referencia + " tiene un error revise sus datos y vuelva a enviarlo" );
                    conerror = 1;
                }
                else
                    lgenerados++;


                if (idpoliza > 0)
                {
                    if (incluir == 1)
                    {
                        string lempresa1 = ciCompanyList11.aliasbdd;
                        string lCadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + lempresa1 + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";

                        SqlConnection DbConnection = new SqlConnection(lCadenaconexion);
                        SqlDataAdapter adapter = new System.Data.SqlClient.SqlDataAdapter();
                        DataSet dataset = new System.Data.DataSet();
                        DbConnection.Open();

                        string sids = "select isnull(max(id),0) from AsocCFDIs;select isnull(max(rowversion),0) from AsocCFDIs ;" +
                            " select m.nummovto, m.guid, c.codigo, p.Guid from movimientospoliza m " +
                            " join cuentas c on m.idcuenta = c.id " +
                            " join polizas p on m.idpoliza = p.id " +
                            " where m.idpoliza = " + idpoliza;

                        sids = "select next from counters WHERE Name = 'Id_AsocCFDI';" +
                            " select m.nummovto, m.guid, c.codigo, p.Guid from movimientospoliza m " +
                            " join cuentas c on m.idcuenta = c.id " +
                            " join polizas p on m.idpoliza = p.id " +
                            " where m.idpoliza = " + idpoliza;
                        SqlCommand mySqlCommand = new SqlCommand(sids);
                        mySqlCommand.Connection = DbConnection;

                        adapter.SelectCommand = mySqlCommand;
                        adapter.Fill(dataset);

                        int i = int.Parse(dataset.Tables[0].Rows[0][0].ToString());
                        //int ii = int.Parse(dataset.Tables[1].Rows[0][0].ToString()) +1; 


                        DataSet ds = new DataSet();
                        //mySqlCommand.CommandType = CommandType.StoredProcedure;
                        SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter();
                        mySqlDataAdapter.SelectCommand = mySqlCommand;


                        string luuid = "";
                        //int iii = 0;
                        int iii = 0;
                        foreach (DataRow yy in dataset.Tables[1].Rows)
                        {
                            int z = 0;
                            foreach (MovPoliza y in x._RegMovtos)
                            {
                                if (yy[2].ToString() == y.cuenta && yy[0].ToString() == (z+1).ToString())
                                {
                                    y.guid = yy[1].ToString();
                                    x._RegMovtos[iii].guid = yy[1].ToString();
                                    poliza.Guid = yy[3].ToString();
                                    luuid = y.uuid;
                                    iii++;
                                    break;
                                }
                                z++;
                                
                            }
                        }

                        if (luuid != "")
                        {
                        string lsql = "insert into AsocCFDIs values (" + i.ToString() + ",ROUND(RAND() * 1000000000,0)" + ",'" + poliza.Guid + "','" + luuid + "',";
                        //<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Referencia><Documento tipo="Poliza" edoPago="0" 
                        //cadReferencia="Póliza de Diario, ejercicio: 2015, periodo: 6, número: 57, empresa: ctgomar, guid: 675ED7F3-26BD-4958-81BD-64D37901C31C."/></Referencia>

                        string lsql1 = "'<?xml version=" + '\u0022' + "1.0" + '\u0022' + " encoding=" + '\u0022' + "UTF-8" + '\u0022' + " standalone=" + '\u0022' + "yes" + '\u0022' + "?><Referencia><Documento tipo=" + '\u0022' + "Poliza" + '\u0022' + " edoPago=" + '\u0022' + "0" + '\u0022' +
                            " cadReferencia=" + '\u0022' + "Póliza de Diario, ejercicio: " + poliza.Fecha.Year + ", periodo: " + poliza.Fecha.Month + ", número: " + poliza.Numero + ", empresa: " + lempresa1 + ", guid: " + poliza.Guid + "." + '\u0022' + "/></Referencia>'" + ",";
                        lsql += lsql1;
                        lsql += "'Contabilidad',1)";

                        mySqlCommand.CommandText = lsql;
                        mySqlCommand.Connection = DbConnection;
                        int iiii = mySqlCommand.ExecuteNonQuery();

                        i++;
                        //ii++;


                        foreach (MovPoliza y in x._RegMovtos)
                        {
                            //<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Referencia><Documento tipo="MovimientoPoliza" edoPago="0" 
                            //cadReferencia="Movimiento de la cuenta: 1010001, empresa: ctgomar, guid: 335F5025-2CD7-4D92-ABD7-6C284B4E633B."/></Referencia>
                            lsql = "insert into AsocCFDIs values (" + i.ToString() + ",ROUND(RAND() * 1000000000,0)" + ",'" + y.guid + "','" + y.uuid + "',";
                            lsql1 = "'<?xml version=" + '\u0022' + "1.0" + '\u0022' + " encoding=" + '\u0022' + "UTF-8" + '\u0022' + " standalone=" + '\u0022' + "yes" + '\u0022' + "?><Referencia><Documento tipo=" + '\u0022' + "MovimientoPoliza" + '\u0022' + " edoPago=" +
                                '\u0022' + "0" + '\u0022' + " cadReferencia=" + '\u0022' + "Movimiento de la cuenta: " + y.cuenta + ", empresa: " + lempresa1 + ", guid: " + y.guid + "." + '\u0022' + "/></Referencia>'" + ",";
                            lsql += lsql1;
                            lsql += "'Contabilidad',1)";
                            mySqlCommand.CommandText = lsql;
                            mySqlCommand.Connection = DbConnection;
                            int iiiii = mySqlCommand.ExecuteNonQuery();
                            // SqlCommand mySqlCommand1 = new SqlCommand("insert AsocCFDIs values (" + i.ToString() + "," + ii.ToString() + ",'" + y.guid + "','" + y.uuid + ""'", ")" );
                            i++;
                            //ii++;

                        }
                        lsql = "UPDATE Counters Set Next = " + i + " WHERE Name = 'Id_AsocCFDI'";
                        mySqlCommand.CommandText = lsql;
                        mySqlCommand.Connection = DbConnection;
                        int iiiiiz = mySqlCommand.ExecuteNonQuery();
                     

                        DbConnection.Close();
                    }
                    }
                    lcuantos++;
                    // MessageBox.Show("Poliza " + _Poliza.Folio.ToString().Trim() + " ya existe");
                }
                else
                {
                    //    MessageBox.Show(poliza.UltimoMsjError);
                }




            }
            MessageBox.Show(lgenerados.ToString() + " Polizas fueron creadas de " + _RegPolizas.Count.ToString());

            if (conerror==1)
                MessageBox.Show("Algunas polizas no fueron cargardas Revise el archivo " + file);

            //sesion.cierraEmpresa();
            //sesion.finalizaConexion();
            //  MessageBox.Show("Proceso Terminado");
            /*
            SDKCONTPAQNGLib.TSdkPoliza poliza = new SDKCONTPAQNGLib.TSdkPoliza();
            
            SDKCONTPAQNGLib.TSdkTipoPoliza tpoliza = new SDKCONTPAQNGLib.TSdkTipoPoliza();
            SDKCONTPAQNGLib.TSdkSesion sesion = new SDKCONTPAQNGLib.TSdkSesion();
            SDKCONTPAQNGLib.TSdkMovimientoPoliza movimientosPoliza = new SDKCONTPAQNGLib.TSdkMovimientoPoliza();
            SDKCONTPAQNGLib.TSdkCuenta cuenta = new SDKCONTPAQNGLib.TSdkCuenta();*/

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string error = mLlenarPolizas();
            if (error != "")
            {
                MessageBox.Show(error);
                return;
            }
            mGrabarPolizas(1);
        }

        private void Romero_FormClosed(object sender, FormClosedEventArgs e)
        {
            Properties.Settings.Default.archivo = botonExcel1.mRegresarNombre();
            Properties.Settings.Default.Save();
            sesion.cierraEmpresa();
            sesion.finalizaConexion();
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
    }
}
