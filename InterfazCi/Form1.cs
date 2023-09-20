using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data.Odbc;

namespace InterfazCi
{
    public partial class Form1 : Form
    {
        SDKCONTPAQNGLib.TSdkSesion sesion = new SDKCONTPAQNGLib.TSdkSesion();        
        public string Cadenaconexion="";
        public string Archivo = "";
        public List<Poliza> _RegPolizas = new List<Poliza>();
        public Poliza _Poliza = new Poliza();
        public Form1()
        {
            InitializeComponent();

            if (Properties.Settings.Default.server != "")
            {
//                Properties.Settings.Default.server = "Toshiba-pc";
  //              Properties.Settings.Default.database = "GeneralesSQL";
    //            Properties.Settings.Default.user = "sa";
      //          Properties.Settings.Default.password = "ady123";

                Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";
              //  MessageBox.Show(Cadenaconexion);
                Archivo = Properties.Settings.Default.archivo;
            }
        }

        private string mLlenarPolizas()
        {
            string aNombreArchivo = botonExcel1.mRegresarNombre();
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + aNombreArchivo + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");

            conn.Open();
            System.Data.OleDb.OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
            cmd.CommandText = "SELECT * FROM [Sheet1$]";
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception ee)
            {
                return ee.Message;
            }

            long xxx;
            xxx = 1000000;


            System.Data.OleDb.OleDbDataReader dr;
            dr = cmd.ExecuteReader();
            Boolean noseguir = false;
            _RegPolizas.Clear();
            long ifolio=0;
            long lfolio = -1;
            if (dr.HasRows)
                while (noseguir == false)
                {
                    dr.Read();

                    try
                    {
                        ifolio = int.Parse(dr[1].ToString());
                    }
                    catch (Exception e)
                    {
                        _RegPolizas.Add(_Poliza);
                        noseguir = true;
                        break;
                    }
                    if (ifolio != lfolio)
                    {
                        if (lfolio !=-1)
                            _RegPolizas.Add(_Poliza);
                        _Poliza = null;
                        _Poliza = new Poliza();
                        _Poliza.Folio = ifolio;
                        switch (dr["Jrnl"].ToString().Trim())
                        {
                            case "160":
                                _Poliza.TipoPol = 2;
                                //_Poliza.TipoPol = int(SDKCONTPAQNGLib.ETIPOPOLIZA.TIPO_INGRESOS);
                                break;
                            case "165":
                                _Poliza.TipoPol = 1;
                                break;
                            case "200":
                                _Poliza.TipoPol = 3;
                                break;
                            case "310":
                                _Poliza.TipoPol = 3;
                                break;

                        }
                        

                        string lfecha = dr["Date"].ToString().Trim();
                        int primerdiagonal = lfecha.IndexOf('/', 0);
                        int segundadiagonal = lfecha.IndexOf('/', primerdiagonal+1);


                        _Poliza.Concepto = dr["Vendor"].ToString().Trim();

                        string ldia = lfecha.Substring(0,primerdiagonal);

                        string lanio = lfecha.Substring(segundadiagonal + 1);
                        string lmes = lfecha.Substring(primerdiagonal + 1, segundadiagonal - (primerdiagonal + 1));
                        _Poliza.FechaAlta = DateTime.Parse(ldia.ToString() + "/" + lmes.ToString() + "/" + lanio.ToString());
                        lfolio = _Poliza.Folio;
                        //_Poliza.TipoPol = 1; 
                    }

                    MovPoliza lRegmovto = new MovPoliza();
                    lRegmovto.cuenta = dr["Account"].ToString();
                    string credito = dr["Credit"].ToString();
                    if (credito =="")
                        lRegmovto.credito = 0;
                    else
                        lRegmovto.credito = decimal.Parse(credito);

                    string debito = dr["Debit"].ToString();
                    if (debito == "")
                        lRegmovto.debito = 0;
                    else
                        lRegmovto.debito = decimal.Parse(debito);

                    lRegmovto.concepto = dr["Your reference"].ToString();
                    _Poliza._RegMovtos.Add(lRegmovto);
                    _Poliza.sMensaje = "";
                    
                }
            return "";
        }


        private string mLlenarPolizasMicroplaneSQL()
        {
            string dsn = textBox1.Text;

            dsn = "DSN=" + textBox1.Text;
            dsn = "DSN=" + textBox1.Text + ";UID=Reports;Pwd=;";
            OdbcConnection DbConnection = new OdbcConnection(dsn);
            try
            {
                DbConnection.Open();
            }
            catch (Exception eeeee)
            {
                return "";
            }


            int mes = 0;

            switch (comboBox1.Text)
            {
                case "Enero": mes =1; break;
                case "Febrero": mes =2; break;
                case "Marzo": mes =3; break;
                case "Abril": mes =4; break;
                case "Mayo": mes =5; break;
                case "Junio": mes =6; break;
                case "Julio": mes =7; break;
                case "Agosto": mes =8; break;
                case "Septiembre": mes =9; break;
                case "Octubre": mes =10; break;
                case "Noviembre": mes =11; break;
                case "Diciembre": mes =12; break;
            }



            string ssql = " SELECT GBKMUT.dagbknr as Jrnl, GBKMUT.bkstnr as [Entry no.], ltrim(str(year(GBKMUT.datum)))+replace(str(month(GBKMUT.datum),2),' ','0')+replace(str(day(GBKMUT.datum),2),' ','0') as Date, " +
                "GBKMUT.reknr as Account, " +
            " GRTBK.oms25_0 as [Account description], " +
            " GBKMUT.faktuurnr as [Our Ref], GBKMUT.docnumber as [Your Reference], GBKMUT.bkstnr_sub as [SO no.], GBKMUT.bdr_val as cantidad, GBKMUT.valcode as [Cur.], GBKMUT.oms25 as Description, GBKMUT.crdnr as [Vendor Number],  " +
            " VENDOR.cmp_code , " +
            " VENDOR.cmp_name as Vendor, " +
            " GBKMUT.res_id, GBKMUT.DocAttachmentID " +
            " FROM GBKMUT  " +
            " join GRTBK on GBKMUT.reknr = GRTBK.reknr  " +
            " join CICMPY as VENDOR on ltrim(GBKMUT.crdnr) = ltrim(VENDOR.cmp_code)  " +
            "  WHERE (month(datum)=" + mes + ") AND (year(datum)=" + textBox2.Text + ") AND (GBKMUT.dagbknr In (160,165,200,310))  " +
            " ORDER BY GBKMUT.bkstnr ";

            
            OdbcCommand DbCommand = DbConnection.CreateCommand();
            DbCommand.CommandText = ssql;
            OdbcDataReader dr = DbCommand.ExecuteReader();



            

            //string aNombreArchivo = botonExcel1.mRegresarNombre();

            /*
            string aNombreArchivo = textBox1.Text;
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + aNombreArchivo + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");

            conn.Open();
            System.Data.OleDb.OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
            cmd.CommandText = "SELECT * FROM [Sheet1$]";
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception ee)
            {
                return ee.Message;
            }

            long xxx;
            xxx = 1000000;
             * */
          

          //  System.Data.OleDb.OleDbDataReader dr;
            //dr = DbCommand.ExecuteReader();
            Boolean noseguir = false;
            _RegPolizas.Clear();
            long ifolio = 0;
            long lfolio = -1;
            if (dr.HasRows)
                while (noseguir == false)
                {
                    dr.Read();

                    try
                    {
                        ifolio = int.Parse(dr[1].ToString());
                    }
                    catch (Exception e)
                    {
                        _RegPolizas.Add(_Poliza);
                        noseguir = true;
                        break;
                    }
                    if (ifolio != lfolio)
                    {
                        if (lfolio != -1)
                            _RegPolizas.Add(_Poliza);
                        _Poliza = null;
                        _Poliza = new Poliza();
                        _Poliza.Folio = ifolio;
                        switch (dr["Jrnl"].ToString().Trim())
                        {
                            case "160":
                                _Poliza.TipoPol = 2;
                                //_Poliza.TipoPol = int(SDKCONTPAQNGLib.ETIPOPOLIZA.TIPO_INGRESOS);
                                break;
                            case "165":
                                _Poliza.TipoPol = 1;
                                break;
                            case "200":
                                _Poliza.TipoPol = 3;
                                break;
                            case "310":
                                _Poliza.TipoPol = 3;
                                break;

                        }


                        string lfecha = dr["Date"].ToString().Trim();
                        int primerdiagonal = lfecha.IndexOf('/', 0);
                        int segundadiagonal = lfecha.IndexOf('/', primerdiagonal + 1);


                        _Poliza.Concepto = dr["Vendor"].ToString().Trim();


                        string ldia = lfecha.Substring(6, 2);
                        string lanio = lfecha.Substring(0,4);
                        string lmes = lfecha.Substring(4,2);
                        /*
                        string ldia = lfecha.Substring(0, primerdiagonal);
                        string lanio = lfecha.Substring(segundadiagonal + 1);
                        string lmes = lfecha.Substring(primerdiagonal + 1, segundadiagonal - (primerdiagonal + 1));
                        */


                        _Poliza.FechaAlta = DateTime.Parse(ldia.ToString() + "/" + lmes.ToString() + "/" + lanio.ToString());
                        lfolio = _Poliza.Folio;
                        //_Poliza.TipoPol = 1; 
                    }

                    MovPoliza lRegmovto = new MovPoliza();
                    lRegmovto.cuenta = dr["Account"].ToString();


                    decimal cantidad = decimal.Parse(dr["Cantidad"].ToString());
                    if (cantidad < 0)
                    {
                        lRegmovto.credito = cantidad * -1;
                        lRegmovto.debito = 0;
                    }
                    else
                    {
                        lRegmovto.credito = 0;
                        lRegmovto.debito = cantidad;
                    }
                    /*

                    string credito = dr["Credit"].ToString();
                    if (credito == "")
                        lRegmovto.credito = 0;
                    else
                        lRegmovto.credito = decimal.Parse(credito);

                    string debito = dr["Debit"].ToString();
                    if (debito == "")
                        lRegmovto.debito = 0;
                    else
                        lRegmovto.debito = decimal.Parse(debito);
                     */

                    lRegmovto.concepto = dr["Your reference"].ToString();
                    lRegmovto.sn = "";
                    _Poliza._RegMovtos.Add(lRegmovto);
                    _Poliza.sMensaje = "";

                }
            return "";
        }
        private string  mLlenarPolizasGranVision()
        {
            string aNombreArchivo = botonExcel1.mRegresarNombre();
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + aNombreArchivo + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");

            conn.Open();
            System.Data.OleDb.OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
            //EMPRESA	ANIO	PERIODO	TIPO POLIZA	NUMERO POLIZA	FECHA	MODULO	ESTATUS	LINEA	CTA	SCTA	SSCTA	OBRA	CC	DESCRIPCION	REFERENCIA	NATURALEZA	IMPORTE

            cmd.CommandText = "SELECT * FROM [Hoja1$]";
            //cmd.CommandText = "SELECT * FROM [Sheet1$]";
            cmd.ExecuteNonQuery();

            long xxx;
            xxx = 1000000;


            System.Data.OleDb.OleDbDataReader dr;
            dr = cmd.ExecuteReader();
            Boolean noseguir = false;
            _RegPolizas.Clear();
            int ifolio = 0;
            int lfolio = -1;
            int renglonblanco = 1;
            if (dr.HasRows)
                while (noseguir == false)
                {
                    renglonblanco = 1;
                    dr.Read();

                    try
                    {
                        ifolio = int.Parse(dr["NUMERO POLIZA"].ToString());
                        //ifolio = int.Parse(dr[5].ToString());
                    }
                    catch (Exception e)
                    {
                        string test = "";
                        try
                        {
                            test = dr["NUMERO POLIZA"].ToString();
                        }
                        catch (Exception ee)
                        {
                            _RegPolizas.Add(_Poliza);
                            noseguir = true;
                            break;
                        }
                        if (dr["NUMERO POLIZA"].ToString() == "")
                        //{
                        //    _RegPolizas.Add(_Poliza);
                        //    noseguir = true;
                        //    break;
                        //}
                        //else
                        {
                            if (_Poliza._RegMovtos.Count>0)
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
                            _Poliza.Concepto = "";
                            //_Poliza.FechaAlta = DateTime.Parse(dr["FECHA POLIZA"].ToString());
                            string lfecha = dr["FECHA POLIZA"].ToString();
                            int primerdiagonal = lfecha.IndexOf('/', 0);
                        int segundadiagonal = lfecha.IndexOf('/', primerdiagonal+1);


                        
                        string ldia = lfecha.Substring(0,primerdiagonal);

                        string lanio = lfecha.Substring(segundadiagonal + 1);
                        string lmes = lfecha.Substring(primerdiagonal + 1, segundadiagonal - (primerdiagonal + 1));
                        _Poliza.FechaAlta = DateTime.Parse(ldia.ToString() + "/" + lmes.ToString() + "/" + lanio.ToString());

                            //_Poliza.FechaAlta = DateTime.Parse(dr[6].ToString());
                            //lfolio = _Poliza.Folio;
                            //_Poliza.TipoPol = int.Parse(dr["TIPO POLIZA"].ToString());
                            _Poliza.TipoPol = int.Parse(dr[4].ToString());
                        }

                        MovPoliza lRegmovto = new MovPoliza();
                        //CTA	SCTA	SSCTA

                        lRegmovto.cuenta = dr["CTA"].ToString().Trim().PadLeft(4, '0') + dr["SCTA"].ToString().Trim().PadLeft(4, '0') + dr["SSCTA"].ToString().Trim().PadLeft(4, '0');

                        lRegmovto.concepto = dr["DESCRIPCION POLIZA"].ToString();
                        lRegmovto.referencia = dr["REFERENCIA"].ToString();

                        lRegmovto.sn = dr["OBRA (Segmento de Negocio)"].ToString();
                        //lRegmovto.cuenta = dr[10].ToString().Trim().PadLeft(4, '0') + dr[11].ToString().Trim().PadLeft(4, '0') + dr[12].ToString().Trim().PadLeft(4, '0');
                        //NATURALEZA
                        string naturaleza = dr["NATURALEZA"].ToString();
                        //string naturaleza = dr[19].ToString();
                        lRegmovto.credito = decimal.Parse(dr["IMPORTE"].ToString());
                        //lRegmovto.credito = decimal.Parse(dr[20].ToString());
                        lRegmovto.debito = 0;
                        if (naturaleza == "C")
                        {
                            lRegmovto.credito = 0;
                            //lRegmovto.debito = decimal.Parse(dr["IMPORTE"].ToString());
                            lRegmovto.debito = decimal.Parse(dr[20].ToString());
                        }
                        _Poliza._RegMovtos.Add(lRegmovto);
                        _Poliza.sMensaje = "";
                    }
                }
            return "";
        }
            





        
    

        private void mGrabarPolizas()
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
                sesion.cierraEmpresa();

            empresa.setSesion(sesion);
            int lcuantos = 0;
            foreach (Poliza x in _RegPolizas)
            {

                poliza = new SDKCONTPAQNGLib.TSdkPoliza();
                poliza.setSesion(sesion);
                

                poliza.iniciarInfo();
                
                poliza.Impresa = 0;
                poliza.Diario = 0;
                poliza.Concepto = "1";
                poliza.CodigoDiario = "";
                poliza.Clase = SDKCONTPAQNGLib.ECLASEPOLIZA.CLASE_AFECTAR;
                
                poliza.Fecha = x.FechaAlta;

                string lfecha = x.FechaAlta.ToShortDateString();
                //poliza.Fecha = lfecha;
                //poliza.Fecha = Convert.ToDateTime("01/11/2014");
                //poliza.Numero = x.Folio;
                poliza.Concepto = x.Concepto;


                //poliza.Tipo = SDKCONTPAQNGLib.ETIPOPOLIZA.TIPO_INGRESOS;
                poliza.Tipo = (SDKCONTPAQNGLib.ETIPOPOLIZA)x.TipoPol;

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
                
                foreach (MovPoliza y in x._RegMovtos)
                {
                    Cuenta.setSesion(sesion);
                    

            
                    Cuenta.buscaPorCodigo(y.cuenta.Trim());

                    //int zz = Cuenta.buscaPorId(15);
                    movimientosPoliza.iniciarInfo();
                    movimientosPoliza.setSesion(sesion);
                    



                    movimientosPoliza.setSdkCuenta(Cuenta);
                    movimientosPoliza.CodigoCuenta = Cuenta.Codigo;
                    movimientosPoliza.CodigoCuenta = y.cuenta.Trim();
                    movimientosPoliza.SegmentoNegocio = y.sn;

                    movimientosPoliza.Referencia = y.referencia;
                    movimientosPoliza.Concepto = y.concepto;
                    movimientosPoliza.NumMovto = lmovto;
                    decimal limporte;
                    limporte = y.debito;
                    movimientosPoliza.TipoMovto = SDKCONTPAQNGLib.ETIPOIMPORTEMOVPOLIZA.MOVPOLIZA_CARGO;
                    if (y.credito!=0)
                    {
                        limporte = y.credito;
                        movimientosPoliza.TipoMovto = SDKCONTPAQNGLib.ETIPOIMPORTEMOVPOLIZA.MOVPOLIZA_ABONO;
                    }
                    movimientosPoliza.Importe = limporte;
                    movimientosPoliza.ImporteME = 0;
                    movimientosPoliza.Concepto = y.concepto;
                    
                    
                    int movAgregado = poliza.agregaMovimiento(movimientosPoliza);
                    //int mov1 = poliza.creaMovimiento(movimientosPoliza);
                    lmovto++;

                }
                int idpoliza = poliza.crea();
                
                if (idpoliza > 0)
                {
                    lcuantos++;
                   // MessageBox.Show("Poliza " + _Poliza.Folio.ToString().Trim() + " ya existe");
                }
                else
                {
                    MessageBox.Show(poliza.UltimoMsjError);
                }
                
                


            }
            MessageBox.Show(lcuantos.ToString() + " Polizas fueron creadas");
 

            
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

        private void button1_Click(object sender, EventArgs e)
        {

            
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        /*    if (Cadenaconexion == "")
                ciCompanyList11.Populate(Cadenaconexion);
            else
            {
                Form4 x = new Form4();
                x.Show();
            }*/
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            
            /*if (Cadenaconexion != "")
                ciCompanyList11.Populate(Cadenaconexion);
            else
            {
                Form4 x = new Form4();
                x.Show();
            }*/
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            DateTime zz = DateTime.Parse("01/02/2015");
            //DateTime zz = DateTime.Parse("01/02/2014");

            //if (DateTime.Today > zz)
             //   return;
            //string error = mLlenarPolizas();
            string error = mLlenarPolizasMicroplaneSQL();
            //string error = mLlenarPolizasGranVision();
            if (error != "")
            {
                MessageBox.Show(error);
                return;
            }
            else
            {
                MessageBox.Show( _RegPolizas.Count() + " Polizas detectadas ");
            }
            //mLlenarPolizasGranVision(); 
            mGrabarPolizas();
        }

        public void mllenarcomboempresas ()
        {
            ciCompanyList11.Populate(Cadenaconexion);
        }
        private void Form1_Shown(object sender, EventArgs e)
        {
            
            if (Cadenaconexion != "")
            {
                ciCompanyList11.Populate(Cadenaconexion);
            }
            else
            {
                this.Visible = false;
                Form4 x = new Form4();
                x.asignaform1(this);
                x.Show();
            }
            textBox1.Text = Properties.Settings.Default.DNS;
            textBox2.Text = Properties.Settings.Default.ejercicio;


            if (Archivo != "")
                botonExcel1.mSetNombre(Archivo);
            this.Text = " Interfaz Microplane Contabilidad " + this.ProductVersion;
            //this.Text = " Interfaz Gran Vision Contabilidad " + this.ProductVersion;

            if (this.Text.Substring(0, 34) == " Interfaz Microplane Contabilidad ")
            {
                comboBox1.Items.Add("Enero");
                comboBox1.Items.Add("Febrero");
                comboBox1.Items.Add("Marzo");
                comboBox1.Items.Add("Abril");
                comboBox1.Items.Add("Mayo");
                comboBox1.Items.Add("Junio");
                comboBox1.Items.Add("Julio");
                comboBox1.Items.Add("Agosto");
                comboBox1.Items.Add("Septiembre");
                comboBox1.Items.Add("Octubre");
                comboBox1.Items.Add("Noviembre");
                comboBox1.Items.Add("Diciembre");
                comboBox1.SelectedIndex = 0;
            }

             
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Properties.Settings.Default.archivo = botonExcel1.mRegresarNombre();
            Properties.Settings.Default.DNS=textBox1.Text;
            Properties.Settings.Default.ejercicio = textBox2.Text;

            Properties.Settings.Default.Save();
            sesion.cierraEmpresa();
            sesion.finalizaConexion();
        }

        private void botonExcel1_Load(object sender, EventArgs e)
        {

        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar))
            {
                e.Handled = false;
            }
            else if (Char.IsControl(e.KeyChar))
            {
                e.Handled = false;
            }
            else if (Char.IsSeparator(e.KeyChar))
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }
           
            
                
        
    }
}

        

        
    

