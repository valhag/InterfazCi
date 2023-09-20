using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;

namespace InterfazCi
{
    public partial class SapToCont : Base
    {
        SDKCONTPAQNGLib.TSdkSesion sesion = new SDKCONTPAQNGLib.TSdkSesion();

        public SapToCont()
        {
            InitializeComponent();
        }

        private void SapToCont_Load(object sender, EventArgs e)
        {
            base.tabControl1.TabPages[2].Text = "Conexion SAP";

            
            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(button3);

            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(progressBar1);

            progressBar1.Top = 205;
            progressBar1.Left = 0;


            

            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(label9);

            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(dateTimePicker1);

            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(checkBox1);

            label9.Top = 75;
            label9.Left = 10;
            label9.Text = "Fecha Documentos SAP";


            dateTimePicker1.Top = 70;
            dateTimePicker1.Left = 150;


            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(label11);
            label11.Top = 105;
            label11.Left =10;
            label11.Text = "Carpeta Bitacora";

            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(textBoxFolder2);
            textBoxFolder2.Top = 100;
            textBoxFolder2.Left = 150;
            textBoxFolder2.Width = 350;
            textBoxFolder2.Text = "";



            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(label12);
            label12.Top = 135;
            label12.Left = 10;
            label12.Text = "Hora envio automatico";

            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(textBox3);
            textBox3.Top = 130;
            textBox3.Left = 150;
            textBox3.Width = 50;
            textBox3.Text = "";


            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(label10);
            label10.Top = 165;
            label10.Left = 10;
            label10.Text = "Ultima Bitacora generada";

            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(textBox1);
            textBox1.Top = 160;
            textBox1.Left = 150;
            textBox1.Width = 250;
            textBox1.Text = "";



            //dateTimePicker1.Text = "Fecha Documentos SAP";

            checkBox1.Top = 190;
            checkBox1.Left = 15;
            checkBox1.Checked = true;

            
            button3.Top = 225;
            button3.Left = 25;
            button3.Text = "Enviar Informacion";


            this.Text = " Interfaz SAP " + this.ProductVersion;

            
        }
        private string mLlenarPolizas(string sFecha)
        {
            //string aNombreArchivo = botonExcel1.mRegresarNombre();
            //SqlConnection conn = new SqlConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + aNombreArchivo + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");



            string Cadenaconexion = "data source =" + Properties.Settings.Default.serverSAP.ToString()
                + ";initial catalog =" + Properties.Settings.Default.dbSAP.ToString() + ";user id = " + Properties.Settings.Default.userSAP.ToString() +
                "; password = " + Properties.Settings.Default.pwdSAP.ToString() + ";";
            SqlConnection _conn = new SqlConnection();

            _conn.ConnectionString = Cadenaconexion;

            // si se conecto grabar los datos en el cnf

            //_conn.Close();


            _conn.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = _conn;
            //EMPRESA	ANIO	PERIODO	TIPO POLIZA	NUMERO POLIZA	FECHA	MODULO	ESTATUS	LINEA	CTA	SCTA	SSCTA	OBRA	CC	DESCRIPCION	REFERENCIA	NATURALEZA	IMPORTE

            cmd.CommandText = "SELECT " +
    "doc.TransId[ID Poliza] " +
    ",doc.RefDate Fecha " +
    ", Memo[Concepto de Poliza] " +
    ", Line_ID[ID Movimiento] " +
    ", Account[Cuenta Contable] " +
    ", Debit[Cargo] " +
    ", Credit Abono " +
    ",ExpUUID UUID " +
    " from JDT1 mov join OJDT doc " +
    " on mov.transid = doc.TransId" +
    " where doc.refdate = '" + sFecha + "'";
            //cmd.CommandText = "SELECT * FROM [Sheet1$]";



            cmd.CommandText = "SELECT doc.U_poliza[ID Poliza] ,doc.RefDate Fecha, doc.U_ConceptoP[Concepto de Poliza] , " +
"mov.ProfitCode Segmento," +
"doc.U_TipoPoliza Tipo, " + 
"Line_ID[ID Movimiento], Account[Cuenta Contable], Debit[Cargo], Credit Abono ,ExpUUID UUID, doc.ref2" +
" from JDT1 mov join OJDT doc  on mov.transid = doc.TransId where doc.refdate = '" + sFecha + "'";
            


            cmd.ExecuteNonQuery();

            long xxx;
            xxx = 1000000;


            SqlDataReader dr;
            dr = cmd.ExecuteReader();
            Boolean noseguir = false;
            _RegPolizas.Clear();
            long ifolio = 0;
            long lfolio = -1;
            int renglonblanco = 1;
            string luuidactual = "";
            if (dr.HasRows)
                while (noseguir == false)
                {
                    renglonblanco = 1;
                    dr.Read();

                    try
                    {
                        ifolio = int.Parse(dr["ID Poliza"].ToString());
                        //ifolio = int.Parse(dr[5].ToString());
                    }
                    catch (Exception e)
                    {
                        string test = "";
                        try
                        {
                            test = dr["ID Poliza"].ToString();
                        }
                        catch (Exception ee)
                        {
                            _RegPolizas.Add(_Poliza);
                            noseguir = true;
                            break;
                        }
                        if (dr["ID Poliza"].ToString() == "")
                        //{
                        //    _RegPolizas.Add(_Poliza);
                        //    noseguir = true;
                        //    break;
                        //}
                        //else
                        {
                            if (_Poliza._RegMovtos.Count > 0)
                                _RegPolizas.Add(_Poliza);
                            renglonblanco = 0;
                        }
                    }
                    if (renglonblanco != 0)
                    {
                        if (ifolio != lfolio)
                        {
                            if (lfolio != -1)
                                _RegPolizas.Add(_Poliza);
                            _Poliza = null;
                            _Poliza = new Poliza();
                            _Poliza.Folio = ifolio;
                            _Poliza.FechaAlta = DateTime.Parse(dr["Fecha"].ToString());
                            //_Poliza.FechaAlta = DateTime.Parse(ldia.ToString() + "/" + lmes.ToString() + "/" + lanio.ToString());

                            //_Poliza.FechaAlta = DateTime.Parse(dr[6].ToString());
                            //lfolio = _Poliza.Folio;
                            //_Poliza.TipoPol = int.Parse(dr["TIPO POLIZA"].ToString());
                            _Poliza.TipoPol = int.Parse(dr["Tipo"].ToString().Trim());
                            _Poliza.Concepto = dr["Concepto de Poliza"].ToString().Trim();
                            _Poliza.Referencia = dr["Id Poliza"].ToString().Trim();

                            lfolio = _Poliza.Folio;

                        }

                        MovPoliza lRegmovto = new MovPoliza();
                        //CTA	SCTA	SSCTA


                        lRegmovto.cuenta = dr["Cuenta Contable"].ToString().Trim();
                        
                        lRegmovto.debito = 0;
                        lRegmovto.credito = 0;
                        if (decimal.Parse(dr["Cargo"].ToString()) != 0)
                            lRegmovto.credito = decimal.Parse(dr["cargo"].ToString());

                        else
                            lRegmovto.debito = decimal.Parse(dr["abono"].ToString());

                        if (luuidactual == "")
                            luuidactual = dr["UUID"].ToString();
                        if (luuidactual != dr["UUID"].ToString() && dr["UUID"].ToString() != "")
                            luuidactual = dr["UUID"].ToString();
                        lRegmovto.uuid = luuidactual;
                        lRegmovto.sn = dr["Segmento"].ToString().Trim();
                        lRegmovto.referencia = dr["ref2"].ToString().Trim();
                        _Poliza._RegMovtos.Add(lRegmovto);
                        _Poliza.sMensaje = "";
                    }
                }
            return "";
        }

        private void button3_Click(object sender, EventArgs e)
        {

            DateTime lfecha = dateTimePicker1.Value;
            string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');

            string dir = System.IO.Directory.GetCurrentDirectory();
            dir = textBoxFolder2.Text;
            string now = System.DateTime.Now.ToString().Substring(0, 19);
            string file = "ErrorLog_" + now.Substring(0, 2) + now.Substring(3, 2) + now.Substring(6, 4) + "_" + now.Substring(11, 2) + now.Substring(14, 2) + now.Substring(17, 2) + ".txt";
            file = dir + "\\" + file;

            string error = mLlenarPolizas(sfecha1);
            if (error != "")
            {
                MessageBox.Show(error);
                return;
            }
            mGrabarPolizas(file,1);
        }
        private void mGrabarPolizas(string file, int incluir = 0)
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

            decimal avance = cuantos / 100;

            progressBar1.Value = 0;

            int polizaactual = 1;

            int lgenerados = 0;

            

            int conerror = 0;

            int idpoliza = 0;
            foreach (Poliza x in _RegPolizas)
            {

                poliza = new SDKCONTPAQNGLib.TSdkPoliza();
                poliza.setSesion(sesion);


                poliza.iniciarInfo();

                poliza.Impresa = 0;
                poliza.Diario = 0;
                //poliza.Concepto = "1";
                poliza.CodigoDiario = "";
                poliza.Concepto = x.Concepto == null ? "": x.Concepto;
                //poliza. = x.Referencia;
                poliza.Clase = SDKCONTPAQNGLib.ECLASEPOLIZA.CLASE_AFECTAR;

                poliza.Fecha = x.FechaAlta;
                //poliza.Fecha = Convert.ToDateTime("01/11/2014");

                poliza.Concepto = x.Concepto == null ? "" : x.Concepto;
                //-- poliza.Guid


                // poliza.Guid = x.Guid;


                poliza.Tipo = SDKCONTPAQNGLib.ETIPOPOLIZA.TIPO_EGRESOS;
                poliza.Tipo = (SDKCONTPAQNGLib.ETIPOPOLIZA)x.TipoPol;
                poliza.Numero = poliza.getUltimoNumero(x.FechaAlta.Year, x.FechaAlta.Month, poliza.Tipo);

                poliza.Numero = int.Parse(x.Folio.ToString());

                //poliza.Concepto = x.Folio.ToString();
                poliza.Concepto = x.Concepto == null ? "" : x.Concepto;                 //poliza.Tipo = (SDKCONTPAQNGLib.ETIPOPOLIZA.

                poliza.SistOrigen = SDKCONTPAQNGLib.ESISTORIGEN.ORIG_CONTPAQNG;

                
                //int idpoliza1 = poliza.crea();
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

                    int zz = Cuenta.buscaPorCodigo(y.cuenta);

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
                    limporte = y.credito;
                    movimientosPoliza.TipoMovto = SDKCONTPAQNGLib.ETIPOIMPORTEMOVPOLIZA.MOVPOLIZA_CARGO;
                    if (y.debito != 0)
                    {
                        limporte = y.debito;
                        movimientosPoliza.TipoMovto = SDKCONTPAQNGLib.ETIPOIMPORTEMOVPOLIZA.MOVPOLIZA_ABONO;
                    }
                    movimientosPoliza.Importe = limporte;
                    movimientosPoliza.ImporteME = 0;

                    //movimientosPoliza.Concepto = y.concepto;

                    movimientosPoliza.Concepto = x.Concepto == null ? "" : x.Concepto;
                    movimientosPoliza.Referencia = y.referencia == null ? "" : x.Referencia ;
                    if (Cuenta.AplicaSegNeg == 1)
                        movimientosPoliza.SegmentoNegocio = y.sn;

                    int movAgregado = 0;
                    movAgregado = poliza.agregaMovimiento(movimientosPoliza);
                    //movAgregado = poliza.agregaMovimiento(movimientosPoliza);

                    //int mov1 = poliza.creaMovimiento(movimientosPoliza);
                    string error1 = poliza.getMensajeError();
                    lmovto++;

                
                }
                
                idpoliza = 0;
                idpoliza = poliza.crea();

                string error = "";
                    error = poliza.getMensajeError();

                if (error != "")
                {

                    System.IO.File.AppendAllText(file, "La poliza con DocumentNo " + x.Referencia + " tiene un error" + error + " revise sus datos y vuelva a enviarlo" + Environment.NewLine);
                    //MessageBox.Show( "La poliza con DocumentNo " + movimientosPoliza.Referencia + " tiene un error revise sus datos y vuelva a enviarlo" );
                    conerror = 1;
                }
                else
                {
                    System.IO.File.AppendAllText(file, "La poliza con DocumentNo " + x.Referencia + " se genero correctamente" + Environment.NewLine);
                    lgenerados++;
                }


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
                            " select m.nummovto, m.guid, c.codigo, p.Guid from movimie ntospoliza m " +
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
                                if (yy[2].ToString() == y.cuenta && yy[0].ToString() == (z + 1).ToString())
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
            
            if (lgenerados> 0)
                System.IO.File.AppendAllText(file, lgenerados.ToString() + " Polizas fueron creadas de " + _RegPolizas.Count.ToString() + Environment.NewLine);
            //MessageBox.Show(lgenerados.ToString() + " Polizas fueron creadas de " + _RegPolizas.Count.ToString());

            //if (conerror == 1)
            //  System.IO.File.AppendAllText(lgenerados.ToString() + " Polizas fueron creadas de " + _RegPolizas.Count.ToString() + Environment.NewLine);
            //MessageBox.Show("Algunas polizas no fueron cargardas Revise el archivo " + file);

            //sesion.cierraEmpresa();
            //sesion.finalizaConexion();
            //  MessageBox.Show("Proceso Terminado");

            System.IO.File.AppendAllText(file,"Proceso Terminado");
            textBox1.Text =  Path.GetFileName(file) ;
            Clipboard.SetText(file);
            /*
            SDKCONTPAQNGLib.TSdkPoliza poliza = new SDKCONTPAQNGLib.TSdkPoliza();
            
            SDKCONTPAQNGLib.TSdkTipoPoliza tpoliza = new SDKCONTPAQNGLib.TSdkTipoPoliza();
            SDKCONTPAQNGLib.TSdkSesion sesion = new SDKCONTPAQNGLib.TSdkSesion();
            SDKCONTPAQNGLib.TSdkMovimientoPoliza movimientosPoliza = new SDKCONTPAQNGLib.TSdkMovimientoPoliza();
            SDKCONTPAQNGLib.TSdkCuenta cuenta = new SDKCONTPAQNGLib.TSdkCuenta();*/

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true && textBox3.Text  != "" && textBox3.Text.Length ==5)
            {
                string lhora = textBox3.Text.Substring(0,2);
                string lminutos = textBox3.Text.Substring(3, 2);


                if (checkBox1.Checked == true && DateTime.Now.Hour.ToString() == lhora && DateTime.Now.Minute.ToString() == lminutos  && (DateTime.Now.Second > 1 && DateTime.Now.Second < 5))
                {

                    string dir = System.IO.Directory.GetCurrentDirectory();
                    dir = textBoxFolder2.Text;
                    string now = System.DateTime.Now.ToString().Substring(0, 19);
                    string file = "ErrorLog_" + now.Substring(0, 2) + now.Substring(3, 2) + now.Substring(6, 4) + "_" + now.Substring(11, 2) + now.Substring(14, 2) + now.Substring(17, 2) + ".txt";
                    file = dir + "\\" + file;

                    DateTime lfecha = DateTime.Now.AddDays(-1);
                    string sfecha1 = lfecha.Year.ToString() + lfecha.Month.ToString().PadLeft(2, '0') + lfecha.Day.ToString().PadLeft(2, '0');

                    string error = mLlenarPolizas(sfecha1);
                    if (error != "")
                    {
                        MessageBox.Show(error);
                        return;
                    }
                    mGrabarPolizas(file);
                    //MessageBox.Show("Proceso Terminado");
                }
            }

            
        }
    }



}
