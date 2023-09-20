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
using System.Globalization;

namespace InterfazCi
{
    public partial class TabPolizas : Base
    {

        protected List<Poliza> _RegPolizas = new List<Poliza>();
        protected Poliza _Poliza = new Poliza();
         SDKCONTPAQNGLib.TSdkSesion sesion = new SDKCONTPAQNGLib.TSdkSesion();

        public TabPolizas()
        {
            InitializeComponent();
        }

        private void TabPolizas_Load(object sender, EventArgs e)
        {
            base.Controls["TabControl1"].Controls["TabPage1"].Controls.Add(botonExcel1);
            botonExcel1.Top = 60;
            botonExcel1.Left = 0;
            botonExcel1.mSetTipo(1);

            base.mOcultaTab3();

            //base.Controls["TabControl1"].Controls["TabPage3"].Visible = false;
            base.ciCompanyList11.mrecorrertxt(92);

            tabControl1.TabPages.Add("TabPage4","Configuracion Cuentas");

            this.ciCompanyList11.SelectedItem += new EventHandler(OnComboChange);


        }

        private void OnComboChange(object sender, EventArgs e)
        {

            //MessageBox.Show("uno");
            
                GetData("select * from sincronizacioncuentas");

        }

        private void button2_Click(object sender, EventArgs e)
        {
            CultureInfo ci = new CultureInfo("es-MX");
            ci = new CultureInfo("es-MX");
            //DateTime fecha = DateTime.Parse( DateTime.Now.ToString("dd/MM/yyyy", ci));

            DateTime xfecha2 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

            DateTime xfechalimite = DateTime.ParseExact("01/12/2023", "dd/MM/yyyy", ci, DateTimeStyles.None);
            DateTime xfecha = DateTime.ParseExact(xfecha2.ToString("dd/MM/yyyy", ci), "dd/MM/yyyy", ci, DateTimeStyles.None);

            if (xfecha >= xfechalimite)

            {
                MessageBox.Show("Error en configuracion");
                return;
            }


            string error = mLlenarPolizas();
            if (error != "")
            {
                MessageBox.Show(error);
                return;
            }
            mGrabarPolizas(0);
        }

        private string mLlenarPolizas()
        {

            CultureInfo ci = new CultureInfo("es-MX");
            ci = new CultureInfo("es-MX");
            var mes = DateTime.Now.ToString("MMMM", ci);

            string aNombreArchivo = botonExcel1.mRegresarNombre();

            string line;
            //List<Cliente> lista = new List<Cliente>();   //Creamos una lista para ir metiendo los clientes
            string path = @"C:\compacw\DIARIO DE POLIZAS 2022.txt";

            path = aNombreArchivo;
            /* try
             {*/
            StreamReader sr = new StreamReader(path);
            line = sr.ReadLine();
            line = sr.ReadLine();
            line = sr.ReadLine();
            line = sr.ReadLine();
            line = sr.ReadLine();
            line = sr.ReadLine();

            int lentrar = 0;
            int lencabezado = 0;
            string lfecha;
            string lfolio;
            string ltipo = "";
            string lreferencia;
            string lusuario;
            string llineatipo;
            string lcuenta;
            //string lcuenta2;
            string limporte1;
            string limporte2;
            string limporteiva;
            string criterio = "";
            _RegPolizas.Clear();
            MovPoliza lRegmovto = new MovPoliza();



            while (line != null)
            {
                line = sr.ReadLine();
                if (lentrar == 0)
                {
                    lentrar = 1;
                    lencabezado = 1;
                }

                if (line == null)
                    break;

                if (line == "\t")
                    continue;

                string[] laux = line.Split('\t'); //Separamos line por los tabuladores, este método devuelve un array


                if (laux.Count() == 1)
                {
                    // nueva poliza
                    if (_Poliza._RegMovtos.Count() == 0 || _Poliza._RegMovtos.Count() == 1)
                        _Poliza.sMensaje = "La poliza numero " + _Poliza.Folio + " tiene " + _Poliza._RegMovtos.Count() + " movimientos";
                    _RegPolizas.Add(_Poliza);
                    lencabezado = 1;
                    _Poliza = new Poliza();
                }
                else
                {

                    if (lencabezado == 0)
                    {
                        lRegmovto = new MovPoliza();
                        if (laux[1] == "D")
                        {
                            lcuenta = laux[3];
                            limporte1 = laux[10];
                            lcuenta = "111001000";
                            lRegmovto.cuenta = lcuenta;
                            if (decimal.Parse(limporte1) < 0)
                                lRegmovto.debito = decimal.Parse(limporte1) * -1;
                            else
                                lRegmovto.credito = decimal.Parse(limporte1);

                            _Poliza._RegMovtos.Add(lRegmovto);

                        }
                        else 
                        if (laux[1] == "K")
                        {
                            lcuenta = laux[3];
                            limporte1 = laux[10];
                            lRegmovto.cuenta = lcuenta;
                            lRegmovto.cuenta = "203001000";
                            /*
                            criterio = " cuentasap = '" + lcuenta + "'";
                            var foundRows = table.Select(criterio);



                            if (foundRows.Length > 0)
                                lRegmovto.cuenta = foundRows[0][2].ToString();
                            else
                            {
                                // no esta sincronizado
                                lRegmovto.cuenta = lRegmovto.cuenta;
                                lRegmovto.error = "La cuenta " + lRegmovto.cuenta + " no esta sincronizada";
                            }*/
                            if (decimal.Parse(limporte1) < 0)
                                lRegmovto.debito = decimal.Parse(limporte1) * -1;
                            else
                                lRegmovto.credito = decimal.Parse(limporte1);

                            _Poliza._RegMovtos.Add(lRegmovto);

                            if (laux.Count() > 19)
                            {
                                lcuenta = laux[20];
                                limporte1 = laux[22];
                                lRegmovto = new MovPoliza();
                                lRegmovto.cuenta = lcuenta;

                                criterio = " cuentasap = '" + lcuenta + "'";
                                var foundRows = table.Select(criterio);

                                if (foundRows.Length > 0)
                                    lRegmovto.cuenta = foundRows[0][2].ToString();
                                else
                                {
                                    // no esta sincronizado
                                    lRegmovto.cuenta = lRegmovto.cuenta;
                                    lRegmovto.error = "La cuenta " + lRegmovto.cuenta + " no esta sincronizada";
                                }
                                if (decimal.Parse(limporte1) < 0)
                                    lRegmovto.debito = decimal.Parse(limporte1) * -1;
                                else
                                    lRegmovto.credito = decimal.Parse(limporte1);
                                _Poliza._RegMovtos.Add(lRegmovto);
                            }



                        }
                        else
                                {
                                    if (laux.Count() > 15)
                                    {
                                        llineatipo = laux[19];
                                        lcuenta = laux[20];
                                        limporte1 = laux[22];
                                        //limporte2 = laux[21];

                                        lRegmovto.cuenta = lcuenta;

                                        criterio = " cuentasap = '" + lcuenta + "'";
                                        var foundRows = table.Select(criterio);

                                        if (foundRows.Length > 0)
                                            lRegmovto.cuenta = foundRows[0][2].ToString();
                                        else
                                        {
                                            // no esta sincronizado
                                            lRegmovto.cuenta = lRegmovto.cuenta;
                                            lRegmovto.error = "La cuenta " + lRegmovto.cuenta + " no esta sincronizada";
                                        }
                                        lRegmovto.concepto = "";
                                        lRegmovto.referencia = "";
                                        lRegmovto.credito = 0;
                                        lRegmovto.debito = 0;
                                        if (llineatipo == "40")
                                            lRegmovto.credito = decimal.Parse(limporte1);
                                        if (llineatipo == "50")
                                            lRegmovto.debito = decimal.Parse(limporte1) * -1;

                                        _Poliza._RegMovtos.Add(lRegmovto);
                                    }
                                }



                    }
                    if (lencabezado == -2) // son los que llevan 2 importes que vienen de iva
                    {
                        llineatipo = laux[19];
                        lcuenta = laux[20];
                        limporte1 = laux[22];
                        limporte2 = laux[14];

                        lRegmovto = new MovPoliza();
                        lRegmovto.cuenta = lcuenta;

                        criterio = " cuentasap = '" + lcuenta + "'";
                        var foundRows = table.Select(criterio);

                        if (foundRows.Length > 0)
                            lRegmovto.cuenta = foundRows[0][1].ToString();
                        else
                        {
                            // no esta sincronizado
                            lRegmovto.cuenta = lRegmovto.cuenta;
                            lRegmovto.error = "La cuenta " + lRegmovto.cuenta + " no esta sincronizada";

                        }

                        lRegmovto.debito = decimal.Parse(limporte1);

                        _Poliza._RegMovtos.Add(lRegmovto);
                        lRegmovto = new MovPoliza();
                        lRegmovto.cuenta = lcuenta;

                        criterio = " cuentasap = '" + lcuenta + "'";
                        foundRows = table.Select(criterio);

                        if (foundRows.Length > 0)
                            lRegmovto.cuenta = foundRows[0][1].ToString();
                        else
                        {
                            // no esta sincronizado
                            lRegmovto.cuenta = lRegmovto.cuenta;
                            lRegmovto.error = "La cuenta " + lRegmovto.cuenta + " no esta sincronizada";
                        }

                        lRegmovto.debito = decimal.Parse(limporte2);

                        _Poliza._RegMovtos.Add(lRegmovto);


                    }


                    if (lencabezado == -1) //LostFocus que tienen iva
                    {
                        if (laux.Count() > 15)
                        {
                            if (ltipo == "SU" && laux[22] == "")
                            {
                                llineatipo = laux[19];
                                lcuenta = laux[20];
                                limporte2 = laux[22];
                                limporte1 = laux[14];
                                limporteiva = laux[10];
                                lencabezado = 2;

                                lRegmovto = new MovPoliza();
                                lRegmovto.cuenta = "0000001"; //cuenta de iva

                                criterio = " cuentasap = '" + lcuenta + "'";
                                var foundRows = table.Select(criterio);

                                if (foundRows.Length > 0)
                                    lRegmovto.cuenta = foundRows[0][1].ToString();
                                else
                                {
                                    // no esta sincronizado
                                    lRegmovto.cuenta = lRegmovto.cuenta;
                                }


                                lRegmovto.credito = decimal.Parse(limporteiva);

                                _Poliza._RegMovtos.Add(lRegmovto);


                                lRegmovto = new MovPoliza();
                                lRegmovto.cuenta = lcuenta;

                                criterio = " cuentasap = '" + lcuenta + "'";
                                foundRows = table.Select(criterio);

                                if (foundRows.Length > 0)
                                    lRegmovto.cuenta = foundRows[0][1].ToString();
                                else
                                {
                                    // no esta sincronizado
                                    lRegmovto.cuenta = lRegmovto.cuenta;
                                }
                                lRegmovto.debito = decimal.Parse(limporte1);

                                _Poliza._RegMovtos.Add(lRegmovto);
                                lRegmovto = new MovPoliza();
                                lRegmovto.cuenta = lcuenta;
                                lRegmovto.debito = decimal.Parse(limporte2);

                                _Poliza._RegMovtos.Add(lRegmovto);
                                lencabezado = -2;


                            }
                        }


                    }



                    if (lencabezado == -20) //KZ
                    {
                        if (laux.Count() > 15)
                        {
                            if (ltipo == "SU" && laux[22] == "")
                            {
                                llineatipo = laux[19];
                                lcuenta = laux[20];
                                limporte2 = laux[22];
                                limporte1 = laux[14];
                                limporteiva = laux[10];
                                lencabezado = 2;

                                lRegmovto = new MovPoliza();
                                lRegmovto.cuenta = "0000001"; //cuenta de iva

                                criterio = " cuentasap = '" + lcuenta + "'";
                                var foundRows = table.Select(criterio);

                                if (foundRows.Length > 0)
                                    lRegmovto.cuenta = foundRows[0][1].ToString();
                                else
                                {
                                    // no esta sincronizado
                                    lRegmovto.cuenta = lRegmovto.cuenta;
                                }


                                lRegmovto.credito = decimal.Parse(limporteiva);

                                _Poliza._RegMovtos.Add(lRegmovto);


                                lRegmovto = new MovPoliza();
                                lRegmovto.cuenta = lcuenta;

                                criterio = " cuentasap = '" + lcuenta + "'";
                                foundRows = table.Select(criterio);

                                if (foundRows.Length > 0)
                                    lRegmovto.cuenta = foundRows[0][1].ToString();
                                else
                                {
                                    // no esta sincronizado
                                    lRegmovto.cuenta = lRegmovto.cuenta;
                                }
                                lRegmovto.debito = decimal.Parse(limporte1);

                                _Poliza._RegMovtos.Add(lRegmovto);
                                lRegmovto = new MovPoliza();
                                lRegmovto.cuenta = lcuenta;
                                lRegmovto.debito = decimal.Parse(limporte2);

                                _Poliza._RegMovtos.Add(lRegmovto);
                                lencabezado = -2;


                            }
                        }


                    }

                    if (lencabezado == 1) // leer los encabezados
                    {
                        lfecha = laux[9];
                        lfolio = laux[2];
                        
                        if (lfolio == "100000014")
                            lfolio = laux[2];
                        if (lfolio == "Mon." || lfolio == "MXN" || lfolio == "")
                            ltipo = laux[8];
                        else
                        {
                            ltipo = laux[8];
                            lreferencia = laux[17];
                            lusuario = laux[29];
                            lencabezado = -1;
                            if (ltipo == "SA")
                                lencabezado = 0; //normales
                            else
                                if (ltipo == "AB")
                                lencabezado = 0;
                            else
                                if (ltipo == "KZ")
                                    lencabezado = -2;
                                lencabezado = 0;
                            _Poliza.FechaAlta = DateTime.Parse(lfecha, ci);
                            _Poliza.Folio = long.Parse(lfolio);

                            switch (ltipo)
                            {
                                case "DZ": _Poliza.TipoPol= 1; break;
                                case "KZ": _Poliza.TipoPol = 2; break;
                                default:
                                        _Poliza.TipoPol = 3; break;
                            }

                           

                        }
                        














                    }

                    

                    
                
                }
                

                //     return "";

            }
            sr.Close();

            if (_Poliza._RegMovtos.Count > 0)
            {
                _RegPolizas.Add(_Poliza);
                lencabezado = 1;
                _Poliza = new Poliza();
            }

            return "";
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

            decimal avance = cuantos / 100;

            progressBar1.Value = 0;

            int polizaactual = 1;

            int lgenerados = 0;

            string dir = System.IO.Directory.GetCurrentDirectory();

            CultureInfo ci = new CultureInfo("es-MX");
            ci = new CultureInfo("es-MX");
            string now = DateTime.Now.ToString("dd/mmm/yyyy hh:mm:ss", ci);
            //string now = System.DateTime.Now.ToString().Substring(0, 19);
            string file = "ErrorLog_" + now.Substring(0, 2) + now.Substring(3, 2) + now.Substring(6, 4) + "_" + now.Substring(11, 2) + now.Substring(14, 2) + now.Substring(17, 2) + ".txt";
            file = dir + "\\" + file;

            int conerrormovtos = 0;
            int conerror = 0;

            foreach (Poliza p in _RegPolizas)
            {
                if (p.sMensaje != "")
                //if (p._RegMovtos.Count() == 0 || p._RegMovtos.Count() == 1)
                {
                    //System.IO.File.AppendAllText(file, "La poliza con DocumentNo " + p.Folio + " tiene "+ p._RegMovtos.Count()  + " movimientos"+   Environment.NewLine);
                    System.IO.File.AppendAllText(file, p.sMensaje + Environment.NewLine);
                    conerrormovtos = 1;
                }
                foreach (MovPoliza m in p._RegMovtos)
                {
                    if (m.error != "")
                    {
                        System.IO.File.AppendAllText(file, "La poliza con DocumentNo " + p.Folio + " tiene el sig. error "+  m.error + Environment.NewLine);
                        conerrormovtos = 1;
                        p.sMensaje = "La poliza con DocumentNo " + p.Folio + " tiene el sig. error " + m.error;
                    }
                }

            }

            if (conerrormovtos == 1)
            {
                MessageBox.Show("Algunas polizas tienen 1 o 0 movimientos, revise bitacora ubicada en " + file);
            }
            foreach (Poliza x in _RegPolizas)
            {
                if (x.sMensaje == "")
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


                    switch (x.TipoPol)
                    {
                        case 1: poliza.Tipo= SDKCONTPAQNGLib.ETIPOPOLIZA.TIPO_INGRESOS; break;
                        case 2: poliza.Tipo = SDKCONTPAQNGLib.ETIPOPOLIZA.TIPO_EGRESOS; break;
                        default:
                            poliza.Tipo = SDKCONTPAQNGLib.ETIPOPOLIZA.TIPO_DIARIO; break;
                    }

                    poliza.Numero = poliza.getUltimoNumero(x.FechaAlta.Year, x.FechaAlta.Month, poliza.Tipo);


                    poliza.Concepto = x.Concepto;

                    poliza.SistOrigen = SDKCONTPAQNGLib.ESISTORIGEN.ORIG_CONTPAQNG;

                    int lmovto = 1;
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

                    string problemacuenta = "";
                    foreach (MovPoliza y in x._RegMovtos)
                    {
                        Cuenta.setSesion(sesion);
                        
                        //string lcuenta = mBuscaCuenta(y.cuenta, y.credito);


                        //Cuenta.buscaPorCodigo(y.cuenta.Trim());



                        int encontrado = Cuenta.buscaPorCodigo(y.cuenta);

                        if (encontrado == 0)
                        {
                            if (problemacuenta == "")
                                problemacuenta = "Revisar cuenta(s) ";
                            problemacuenta += y.cuenta + ", ";
                        }
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
                        if (x.Concepto == null)
                            x.Concepto = "";
                        movimientosPoliza.Concepto = x.Concepto;
                        movimientosPoliza.Referencia = y.referencia;


                        int movAgregado = poliza.agregaMovimiento(movimientosPoliza);
                        //int mov1 = poliza.creaMovimiento(movimientosPoliza);
                        lmovto++;

                    }
                    int idpoliza = poliza.crea();

                    if (idpoliza == 0)
                    {

                        if (problemacuenta == "")
                            System.IO.File.AppendAllText(file, "La poliza con DocumentNo " + x.Folio + " tiene un error revise sus datos y vuelva a enviarlo" + Environment.NewLine);
                        else
                            System.IO.File.AppendAllText(file, "La poliza con DocumentNo " + x.Folio + " " + problemacuenta + Environment.NewLine);

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
                            int i = 1;
                            if (dataset.Tables[0].Rows.Count > 0)
                                i = int.Parse(dataset.Tables[0].Rows[0][0].ToString());
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
            }
            MessageBox.Show(lgenerados.ToString() + " Polizas fueron creadas de " + _RegPolizas.Count.ToString());

            if (conerror == 1)
                MessageBox.Show("Algunas polizas no fueron cargardas Revise el archivo " + file);

            //sesion.cierraEmpresa();
            //sesion.finalizaConexion();
            //  MessageBox.Show("Proceso Terminado");

                }

        private void TabPolizas_Shown(object sender, EventArgs e)
        {
            tabControl1.Controls["TabPage4"].Controls.Add(dataGridView1);
            base.Controls["TabControl1"].Controls["TabPage4"].Controls.Add(dataGridView1);
            base.Controls["TabControl1"].Controls["TabPage4"].Controls.Add(button3);

            button3.Top = 0;
            button3.Left = 440;


            base.Controls["TabControl1"].Controls["TabPage4"].Controls.Add(textBox1);
            base.Controls["TabControl1"].Controls["TabPage4"].Controls.Add(textBox2);


            textBox1.Top = 0;
            textBox2.Top = 0;
            textBox1.Left = 20;
            textBox2.Left = 220;

            

            /*
            foreach (DataGridViewRow row in dataGridViewUsuarios.Rows)
            {
                if (row.Cells[2].Value.ToString().Contains(ctxtbusquedaUsuarios.txt.Text))
                {
                    //rowIndex = row.Index;
                    dataGridViewUsuarios.Rows[row.Index].Selected = true;
                    dataGridViewUsuarios.FirstDisplayedScrollingRowIndex = row.Index;
                    dataGridViewUsuarios.Refresh();
                    break;
                }
            }*/
            dataGridView1.Top = 20;
            dataGridView1.Left = 5;

            dataGridView1.DataSource = bindingSource1;

            


            GetData("select * from sincronizacioncuentas");

        }

        

        private BindingSource bindingSource1 = new BindingSource();
        private SqlDataAdapter dataAdapter = new SqlDataAdapter();

        DataTable table;
        private void GetData(string selectCommand)
        {
            
            try
            {
                // Specify a connection string.
                // Replace <SQL Server> with the SQL Server for your Northwind sample database.
                // Replace "Integrated Security=True" with user login information if necessary.
                string lempresa1 = ciCompanyList11.aliasbdd;

                string connectionString = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + lempresa1 + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";


                
                // Create a new data adapter based on the specified query.
                dataAdapter = new SqlDataAdapter(selectCommand, connectionString);

                // Create a command builder to generate SQL update, insert, and
                // delete commands based on selectCommand.
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);

                bindingSource1.DataSource = null;

                // Populate a new data table and bind it to the BindingSource.
                table = new DataTable
                {
                    Locale = CultureInfo.InvariantCulture
                };
                dataAdapter.Fill(table);
                bindingSource1.DataSource = table;

                // Resize the DataGridView columns to fit the newly loaded content.

                if (dataGridView1.Columns.Count > 0 )
                {
                    dataGridView1.RowHeadersVisible = false;
                    dataGridView1.Columns[0].Visible = false;
                    dataGridView1.Columns[1].Width = dataGridView1.Width / 2;
                    dataGridView1.Columns[2].Width = dataGridView1.Width / 2;
                }
                
                //dataGridView1.AutoResizeColumns(
                //  DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader);
            }
            catch (SqlException)
            {
                MessageBox.Show("La empresa actualmente seleccionada no es compatible con la interfaz sap");
            }
        }

        private void dataGridView1_RowLeave(object sender, DataGridViewCellEventArgs e)
        {
        //    bindingSource1.EndEdit();
            //this.myTableAdapter.Update(this.myDataSet.Customers);
          //  dataAdapter.Update(table);
        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            //DataRow x = table.NewRow();
            //x[0] = "uno";
            //x[1] = "dos";
            //table.Rows.Add(x);


            //dataAdapter.Update(table);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                dataAdapter.Update(table);
            }
            catch (Exception exceptionObj)
            {
                MessageBox.Show(exceptionObj.Message.ToString());
            }
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Delete)
                {
                MessageBox.Show("delete pressed");
                dataGridView1.Rows.Remove(dataGridView1.CurrentRow);
                //DataGridView1.Rows.Remove(DataGridView1.Rows[DataGridView1.Rows.Count - 1]);
                e.Handled = true;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[1].Value != null)
                if (row.Cells[1].Value.ToString().Contains(textBox1.Text))
                {
                    //rowIndex = row.Index;
                    dataGridView1.Rows[row.Index].Selected = true;
                    dataGridView1.FirstDisplayedScrollingRowIndex = row.Index;
                    dataGridView1.Refresh();
                    break;
                }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[2].Value != null)
                    if (row.Cells[2].Value.ToString().Contains(textBox2.Text))
                    {
                        //rowIndex = row.Index;
                        dataGridView1.Rows[row.Index].Selected = true;
                        dataGridView1.FirstDisplayedScrollingRowIndex = row.Index;
                        dataGridView1.Refresh();
                        break;
                    }
            }
        }
    }
}
