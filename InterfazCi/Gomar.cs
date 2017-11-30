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

namespace InterfazCi
{
    public partial class Gomar : Form
    {
        SDKCONTPAQNGLib.TSdkSesion sesion = new SDKCONTPAQNGLib.TSdkSesion();        
        public string Cadenaconexion="";
        public string Archivo = "";
        public List<Poliza> _RegPolizas = new List<Poliza>();
        public Poliza _Poliza = new Poliza();
        public Gomar()
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
            int ifolio=0;
            int lfolio = -1;
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

        private string mLlenarPolizasGomar()
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
            int ifolio = 0;
            int lfolio = 0;
            string ltipo = dr.GetName(2).ToString();
            string laplicacion = dr.GetName(5).ToString();
            int atipo = 0;
            string documento = "";
            switch (ltipo)
            {
                case "Recepción o poliza contable" :
                    atipo = 3; // compras
                    documento = "Compras";
                    break;
                case "Suc.": case "Suc#":
                    atipo = 3; 
                    documento = "Facturas";
                    break;
                case "Nombre Proveedor":
                    atipo = 2; 
                    documento = "PagoProveedor";
                    break;
                case "Documento Relacionado":
                    atipo = 3;
                    documento = "NotaCreditoCliente";
                    break;
                case "Docto Relacionado":
                    atipo = 3;
                    documento = "NotaCreditoProveedor";
                    break;
                case "Nombre Cliente":
                    atipo = 1;
                    documento = "PagoCliente";
                    break;
                case "Anticipo Cliente":
                    atipo = 1;
                    documento = "AnticipoCliente";
                    break;
                case "Anticipo Proveedor":
                    atipo = 2;
                    documento = "AnticipoProveedor";
                    break;
                case "Anticipo Aplicado":
                    atipo = 3;
                    if (laplicacion == "Nombre de cliente")                         
                        documento = "AplicacionCliente";
                    else
                    
                        documento = "AplicacionProveedor";
                    break;
            }
            

            if (dr.HasRows)
                while (noseguir == false)
                {
                    dr.Read();

                    //try
                    //{
                    //    ifolio = 0;
                    //}
                    //catch (Exception e)
                    //{
                    //    _RegPolizas.Add(_Poliza);
                    //    noseguir = true;
                    //    break;
                    //}
                    //if (ifolio != lfolio)
                    //{
                        //if (lfolio != -1)
                          //  _RegPolizas.Add(_Poliza);
                        _Poliza = null;
                        _Poliza = new Poliza();
                        _Poliza.Folio = ifolio;
                        _Poliza.TipoPol = atipo;



                        string lfecha = "";
                        try
                        {
                            lfecha = dr["Fecha"].ToString().Trim();
                        }
                        catch (Exception eee)
                        { noseguir = true; break; }
                        int primerdiagonal = lfecha.IndexOf('/', 0);
                        int segundadiagonal = lfecha.IndexOf('/', primerdiagonal + 1);


                        

                        string ldia = lfecha.Substring(0, primerdiagonal);

                        string lanio = lfecha.Substring(segundadiagonal + 1);
                        string lmes = lfecha.Substring(primerdiagonal + 1, segundadiagonal - (primerdiagonal + 1));


                        _Poliza.FechaAlta = DateTime.Parse(ldia.ToString() + "/" + lmes.ToString() + "/" + lanio.ToString());

                        switch (documento){
                            case "Compras":
                                _Poliza.Concepto = "Compra "+ dr[0].ToString().Trim();
                                _Poliza.Referencia = dr[0].ToString().Trim();
                                break;
                            case "Facturas":
                                _Poliza.Concepto = "Facturacion";
                                _Poliza.Referencia = dr[0].ToString().Trim();
                                break;
                            case "PagoProveedor":
                                _Poliza.Concepto = "Pago Proveedor";
                                _Poliza.Referencia = dr[0].ToString().Trim();
                                break;
                            case "Documento Relacionado":
                                _Poliza.Concepto = "Nota Credito Cliente";
                                _Poliza.Referencia = dr[0].ToString().Trim();
                                break;
                            case "Docto Relacionado":
                                _Poliza.Concepto = "Nota Credito Proveedor";
                                _Poliza.Referencia = dr[0].ToString().Trim();
                                break;
                            case "PagoCliente":
                                _Poliza.Concepto = "Pago Cliente";
                                _Poliza.Referencia = dr[0].ToString().Trim();
                                break;
                            case "AnticipoCliente":
                                _Poliza.Concepto = "Anticipo Cliente";
                                _Poliza.Referencia = dr[0].ToString().Trim();
                                break;
                            case "AnticipoProveedor":
                                _Poliza.Concepto = "Anticipo Proveedor";
                                _Poliza.Referencia = dr[0].ToString().Trim();
                                break;
                            case "AplicacionCliente":
                                _Poliza.Concepto = "Aplicacion Anticipo Cliente";
                                _Poliza.Referencia = dr[0].ToString().Trim();
                                break;
                            case "AplicacionProveedor":
                                _Poliza.Concepto = "Aplicacion Anticipo Proveedor";
                                _Poliza.Referencia = dr[0].ToString().Trim();
                                break;
                        }
                        
                        
                        
                        //lfolio = _Poliza.Folio;
                        //_Poliza.TipoPol = 1; 
                    //}

                    MovPoliza lRegmovto1 = new MovPoliza();
                    MovPoliza lRegmovto2 = new MovPoliza();
                    MovPoliza lRegmovto3 = new MovPoliza();
                    MovPoliza lRegmovto4 = new MovPoliza();

                    switch (documento)
                    {
                        case "Compras":
                            lRegmovto1.cuenta = textBoxCompraCuenta1.Text;
                            lRegmovto1.credito = 0;
                            lRegmovto1.debito = decimal.Parse(dr["Subtotal Factura"].ToString());
                            lRegmovto1.concepto = "Compra " + dr[0].ToString().Trim();
                            lRegmovto1.referencia = dr[0].ToString().Trim();
                            lRegmovto1.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto1);

                            lRegmovto2.cuenta = textBoxCompraCuenta2.Text;
                            lRegmovto2.credito = 0;
                            lRegmovto2.debito = decimal.Parse(dr["Total Cr Tributario"].ToString()); 
                            lRegmovto2.concepto = "Compra " + dr[0].ToString().Trim();
                            lRegmovto2.referencia = dr[0].ToString().Trim();
                            lRegmovto2.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto2);

                            lRegmovto3.cuenta = dr["contable"].ToString();
                            lRegmovto3.credito = decimal.Parse(dr["Total Factura"].ToString()); 
                            lRegmovto3.debito = 0;
                            lRegmovto3.concepto = "Compra " + dr[0].ToString().Trim();
                            lRegmovto3.referencia = dr[0].ToString().Trim();
                            lRegmovto3.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto3);
                            break;
                        case "Facturas":
                            lRegmovto1.cuenta = dr["contable"].ToString();
                            lRegmovto1.credito = 0; 
                            lRegmovto1.debito = decimal.Parse(dr["Total Factura"].ToString()); ;
                            _Poliza.Concepto = "Facturacion";
                            _Poliza.Referencia = dr[0].ToString().Trim();
                            lRegmovto1.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto1);

                            lRegmovto2.cuenta = textBoxFacturaCuenta1.Text;
                            lRegmovto2.credito = decimal.Parse(dr["Subtotal Factura"].ToString()); ;
                            lRegmovto2.debito = 0;
                            lRegmovto2.concepto = "Facturacion";
                            lRegmovto2.referencia = dr[0].ToString().Trim();
                            lRegmovto2.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto2);

                            lRegmovto3.cuenta = textBoxFacturaCuenta2.Text;
                            lRegmovto3.credito = decimal.Parse(dr["Total Impuestos"].ToString()); ;
                            lRegmovto3.debito = 0;
                            lRegmovto3.concepto = "Facturacion";
                            lRegmovto3.referencia = dr[0].ToString().Trim();
                            lRegmovto3.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto3);

                            break;   
                    case "PagoProveedor":
                            lRegmovto1.cuenta = dr["Cta Contable"].ToString();
                            lRegmovto1.credito = 0; 
                            lRegmovto1.debito = decimal.Parse(dr["Cobrado"].ToString());
                            lRegmovto1.concepto = "Pago Proveedor";
                            lRegmovto1.uuid = dr["UUID"].ToString();
                            _Poliza.Concepto = "Pago Proveedor";
                            _Poliza.Referencia = dr[0].ToString().Trim();
                            _Poliza._RegMovtos.Add(lRegmovto1);

                            lRegmovto2.cuenta = textBoxPagoProveedor1.Text;
                            lRegmovto2.credito = decimal.Parse(dr["Cobrado"].ToString()); ;
                            lRegmovto2.debito = 0;
                            lRegmovto2.concepto = "Pago Proveedor";
                            lRegmovto2.referencia = dr[0].ToString().Trim();

                            lRegmovto2.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto2);

                            lRegmovto3.cuenta = textBoxPagoProveedor2.Text;
                            lRegmovto3.debito = decimal.Parse(dr["IVA"].ToString()); ;
                            lRegmovto3.credito = 0;
                            lRegmovto3.concepto = "Pago Proveedor";
                            lRegmovto3.referencia = dr[0].ToString().Trim();
                            lRegmovto3.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto3);

                            lRegmovto4.cuenta = textBoxPagoProveedor3.Text;
                            lRegmovto4.credito = decimal.Parse(dr["IVA"].ToString()); ;
                            lRegmovto4.debito = 0;
                            lRegmovto4.concepto = "Pago Proveedor";
                            lRegmovto4.referencia = dr[0].ToString().Trim();
                            lRegmovto4.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto4);

                            break;

                            //"Nota Credito Cliente";
                    case "NotaCreditoCliente":

                            string lcuenta = dr["CUENTA CTE"].ToString();
                            lcuenta = "429" + lcuenta.Substring(3);
                            lRegmovto1.cuenta = lcuenta;
                            lRegmovto1.credito = 0;
                            lRegmovto1.debito = decimal.Parse(dr[5].ToString());
                            lRegmovto1.concepto = "Nota Credito Cliente";
                            _Poliza.Concepto = "Nota Credito Cliente";
                            _Poliza.Referencia = dr[0].ToString().Trim();
                            lRegmovto1.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto1);

                            lRegmovto2.cuenta = textBoxNotaCreditoCteCuenta1.Text;
                            lRegmovto2.debito = decimal.Parse(dr[6].ToString()); ;
                            lRegmovto2.credito = 0;
                            lRegmovto2.concepto = "Nota Credito Cliente";
                            lRegmovto2.referencia = dr[0].ToString().Trim();
                            lRegmovto2.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto2);

                            lRegmovto3.cuenta = dr["CUENTA CTE"].ToString();
                            lRegmovto3.credito = decimal.Parse(dr[7].ToString()); ;
                            lRegmovto3.debito = 0;
                            lRegmovto3.concepto = "Nota Credito Cliente";
                            lRegmovto3.referencia = dr[0].ToString().Trim();
                            lRegmovto3.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto3);
                            break;
                    case "NotaCreditoProveedor":

                            string lcuenta1 = dr["Cuenta proveedor"].ToString();
                            //lcuenta1 = "429" + lcuenta1.Substring(3);
                            lRegmovto1.cuenta = lcuenta1;
                            lRegmovto1.credito = 0;
                            lRegmovto1.debito = decimal.Parse(dr["Total"].ToString());
                            lRegmovto1.concepto = "Nota Credito Proveedor";
                            _Poliza.Concepto = "Nota Credito Proveedor";
                            _Poliza.Referencia = dr[0].ToString().Trim();
                            lRegmovto1.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto1);

                            lRegmovto2.cuenta = textBoxNotaCreditoProvCuenta1.Text;
                            lRegmovto2.debito = 0;
                            lRegmovto2.credito = decimal.Parse(dr["Subtotal 1"].ToString()); ;;
                            lRegmovto2.concepto = "Nota Credito Proveedor";
                            lRegmovto2.referencia = dr[0].ToString().Trim();
                            lRegmovto2.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto2);

                            lRegmovto3.cuenta = textBoxNotaCreditoProvCuenta2.Text;
                            lRegmovto3.credito = decimal.Parse(dr["Impuesto"].ToString()); ;
                            lRegmovto3.debito = 0;
                            lRegmovto3.concepto = "Nota Credito Proveedor";
                            lRegmovto3.referencia = dr[0].ToString().Trim();
                            lRegmovto3.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto3);
                            break;
                    case "PagoCliente":
                            lRegmovto1.cuenta = dr["Banco"].ToString();
                            lRegmovto1.credito = 0;
                            lRegmovto1.debito = decimal.Parse(dr["Cobrado"].ToString());
                            lRegmovto1.concepto = "Pago Cliente";
                            _Poliza.Concepto = "Pago Cliente";
                            _Poliza.Referencia = dr[0].ToString().Trim();
                            lRegmovto1.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto1);

                            lRegmovto2.cuenta = dr["Cta Contable"].ToString();
                            lRegmovto2.credito = decimal.Parse(dr["Cobrado"].ToString()); ;
                            lRegmovto2.debito = 0;
                            lRegmovto2.concepto = "Pago Proveedor";
                            lRegmovto2.referencia = dr[0].ToString().Trim();
                            lRegmovto2.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto2);

                            lRegmovto3.cuenta = textBoxpagoCuenta1.Text;
                            lRegmovto3.debito = decimal.Parse(dr["IVA"].ToString()); ;
                            lRegmovto3.credito = 0;
                            lRegmovto3.concepto = "Pago Cliente";
                            lRegmovto3.referencia = dr[0].ToString().Trim();
                            lRegmovto3.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto3);

                            lRegmovto4.cuenta = textBoxpagoCuenta2.Text;
                            lRegmovto4.credito = decimal.Parse(dr["IVA"].ToString()); ;
                            lRegmovto4.debito = 0;
                            lRegmovto4.concepto = "Pago Cliente";
                            lRegmovto4.referencia = dr[0].ToString().Trim();
                            lRegmovto4.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto4);

                            break;
                    case "AnticipoCliente":

                            lcuenta = dr["Banco"].ToString();
                            
                            lRegmovto1.cuenta = lcuenta;
                            lRegmovto1.credito = 0;
                            lRegmovto1.debito = decimal.Parse(dr["Cobrado"].ToString());
                            lRegmovto1.concepto = "Anticipo Cliente";
                            _Poliza.Concepto = "Anticipo Cliente";
                            _Poliza.Referencia = dr[0].ToString().Trim();
                            lRegmovto1.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto1);

                            lcuenta = dr["Cta Contable"].ToString();
                            lcuenta = "225" + lcuenta.Substring(3);

                            lRegmovto2.cuenta = lcuenta;
                            lRegmovto2.debito = 0 ;
                            lRegmovto2.credito = decimal.Parse(dr["Importe"].ToString());
                            lRegmovto2.concepto = "Anticipo Cliente";
                            lRegmovto2.referencia = dr[0].ToString().Trim();
                            lRegmovto2.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto2);

                            lRegmovto3.cuenta = textBoxAnticipoCte1.Text;
                            lRegmovto3.credito = decimal.Parse(dr["IVA"].ToString()); ;
                            lRegmovto3.debito = 0;
                            lRegmovto3.concepto = "Anticipo Cliente";
                            lRegmovto3.referencia = dr[0].ToString().Trim();
                            lRegmovto3.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto3);
                            break;
                    case "AnticipoProveedor":

                            lcuenta = dr["Cta Contable"].ToString();
                            lcuenta = "117" + lcuenta.Substring(3);
                            lRegmovto1.cuenta = lcuenta;
                            lRegmovto1.debito = decimal.Parse(dr["Importe"].ToString());
                            lRegmovto1.credito = 0;
                            lRegmovto1.concepto = "Anticipo Proveedor";
                            _Poliza.Concepto = "Anticipo Proveedor";
                            _Poliza.Referencia = dr[0].ToString().Trim();
                            lRegmovto1.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto1);

                            

                            lRegmovto2.cuenta = textBoxAnticipoProv1.Text;
                            lRegmovto2.credito = 0;
                            lRegmovto2.debito = decimal.Parse(dr["IVA"].ToString());
                            lRegmovto2.concepto = "Anticipo Proveedor";
                            lRegmovto2.referencia = dr[0].ToString().Trim();
                            lRegmovto2.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto2);

                            lRegmovto3.cuenta = dr["Banco"].ToString();
                            lRegmovto3.debito = 0;
                            lRegmovto3.credito = decimal.Parse(dr["Cobrado"].ToString());
                            lRegmovto3.concepto = "Anticipo Proveedor";
                            lRegmovto3.referencia = dr[0].ToString().Trim();
                            lRegmovto3.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto3);
                            break;
                    case "AplicacionCliente":
                            lcuenta = dr["Contable"].ToString();
                            lcuenta = "225" + lcuenta.Substring(3);
                            lRegmovto1.cuenta = lcuenta;
                            lRegmovto1.credito = 0;
                            lRegmovto1.debito = decimal.Parse(dr["Anticipo Aplicado"].ToString());
                            lRegmovto1.concepto = "Aplicacion Anticipo Cliente";
                            _Poliza.Concepto = "Aplicacion Anticipo Cliente";
                            _Poliza.Referencia = dr[0].ToString().Trim();
                            lRegmovto1.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto1);

                            lRegmovto2.cuenta = textBoxAplCte1.Text;
                            lRegmovto2.debito = decimal.Parse(dr["IVA"].ToString()); ;
                            lRegmovto2.credito = 0;
                            lRegmovto2.concepto = "Aplicacion Anticipo Cliente";
                            lRegmovto2.referencia = dr[0].ToString().Trim();
                            lRegmovto2.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto2);

                            lRegmovto3.cuenta = dr["Contable"].ToString();
                            lRegmovto3.credito = decimal.Parse(dr["Anticipo Aplicado"].ToString()); ;
                            lRegmovto3.debito = 0;
                            lRegmovto3.concepto = "Aplicacion Anticipo Cliente";
                            lRegmovto3.referencia = dr[0].ToString().Trim();
                            lRegmovto3.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto3);

                            lRegmovto4.cuenta = textBoxAplCte1.Text;
                            lRegmovto4.credito = decimal.Parse(dr["IVA"].ToString()); ;
                            lRegmovto4.debito = 0;
                            lRegmovto4.concepto = "Aplicacion Anticipo Cliente";
                            lRegmovto4.referencia = dr[0].ToString().Trim();
                            lRegmovto4.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto4);
                            break;
                    case "AplicacionProveedor":
                            lcuenta = dr[4].ToString();
                            
                            lRegmovto1.cuenta = lcuenta;
                            lRegmovto1.credito = 0;
                            lRegmovto1.debito = decimal.Parse(dr["Anticipo Aplicado"].ToString());
                            lRegmovto1.concepto = "Aplicacion Anticipo Proveedor";
                            _Poliza.Concepto = "Aplicacion Anticipo Proveedor";
                            _Poliza.Referencia = dr[0].ToString().Trim();
                            lRegmovto1.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto1);

                            lRegmovto2.cuenta = textBoxAplProv1.Text;
                            lRegmovto2.debito = decimal.Parse(dr["IVA"].ToString()); ;
                            lRegmovto2.credito = 0;
                            lRegmovto2.concepto = "Aplicacion Anticipo Proveedor";
                            lRegmovto2.referencia = dr[0].ToString().Trim();
                            lRegmovto2.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto2);

                            
                            lcuenta = "117" + lcuenta.Substring(3);
                            lRegmovto3.cuenta = lcuenta;
                            lRegmovto3.credito = decimal.Parse(dr["Anticipo Aplicado"].ToString()); ;
                            lRegmovto3.debito = 0;
                            lRegmovto3.concepto = "Aplicacion Anticipo Proveedor";
                            lRegmovto3.referencia = dr[0].ToString().Trim();
                            lRegmovto3.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto3);

                            lRegmovto4.cuenta = textBoxAplProv1.Text;
                            lRegmovto4.credito = decimal.Parse(dr["IVA"].ToString()); ;
                            lRegmovto4.debito = 0;
                            lRegmovto4.concepto = "Aplicacion Anticipo Proveedor";
                            lRegmovto4.referencia = dr[0].ToString().Trim();
                            lRegmovto4.uuid = dr["UUID"].ToString();
                            _Poliza._RegMovtos.Add(lRegmovto4);
                            break;
                    
                }

                    if (_Poliza.TipoPol == 3)
                    {
                        


                    }
                    _RegPolizas.Add(_Poliza);



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
                            _Poliza.FechaAlta = DateTime.Parse(dr["FECHA POLIZA"].ToString());
                            //_Poliza.FechaAlta = DateTime.Parse(dr[6].ToString());
                            lfolio = _Poliza.Folio;
                            //_Poliza.TipoPol = int.Parse(dr["TIPO POLIZA"].ToString());
                            _Poliza.TipoPol = int.Parse(dr[4].ToString());
                        }

                        MovPoliza lRegmovto = new MovPoliza();
                        //CTA	SCTA	SSCTA

                        lRegmovto.cuenta = dr["CTA"].ToString().Trim().PadLeft(4, '0') + dr["SCTA"].ToString().Trim().PadLeft(4, '0') + dr["SSCTA"].ToString().Trim().PadLeft(4, '0');
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
            





        
    

        private void mGrabarPolizas(int incluir=0)
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
                //poliza.Concepto = "1";
                poliza.CodigoDiario = "";
                poliza.Concepto = x.Concepto;
                //poliza. = x.Referencia;
                poliza.Clase = SDKCONTPAQNGLib.ECLASEPOLIZA.CLASE_AFECTAR;
                
                poliza.Fecha = x.FechaAlta;
                //poliza.Fecha = Convert.ToDateTime("01/11/2014");
                
                poliza.Concepto = x.Concepto;


                poliza.Tipo = SDKCONTPAQNGLib.ETIPOPOLIZA.TIPO_EGRESOS;
                poliza.Tipo = (SDKCONTPAQNGLib.ETIPOPOLIZA)x.TipoPol;
                poliza.Numero = poliza.getUltimoNumero(x.FechaAlta.Year, x.FechaAlta.Month, poliza.Tipo);
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
                    movimientosPoliza.Referencia = x.Referencia;
                    
                    
                    int movAgregado = poliza.agregaMovimiento(movimientosPoliza);
                    //int mov1 = poliza.creaMovimiento(movimientosPoliza);
                    lmovto++;

                }
                int idpoliza = poliza.crea();
                
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

                        int i = int.Parse(dataset.Tables[0].Rows[0][0].ToString()) ; 
                        //int ii = int.Parse(dataset.Tables[1].Rows[0][0].ToString()) +1; 
                        

                        DataSet ds = new DataSet();
                        //mySqlCommand.CommandType = CommandType.StoredProcedure;
                        SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter();
                        mySqlDataAdapter.SelectCommand = mySqlCommand;


                        string luuid = "";
                        foreach (DataRow yy in dataset.Tables[1].Rows)
                        {
                            foreach (MovPoliza y in x._RegMovtos)
                            {
                                if (yy[2].ToString() == y.cuenta)
                                {
                                    y.guid = yy[1].ToString();
                                    poliza.Guid = yy[3].ToString();
                                    luuid = y.uuid;
                                }
                            }
                        }


                        string lsql = "insert into AsocCFDIs values (" + i.ToString() + ",ROUND(RAND() * 1000000000,0)" + ",'" + poliza.Guid + "','" + luuid + "',";
                        //<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Referencia><Documento tipo="Poliza" edoPago="0" 
                        //cadReferencia="Póliza de Diario, ejercicio: 2015, periodo: 6, número: 57, empresa: ctgomar, guid: 675ED7F3-26BD-4958-81BD-64D37901C31C."/></Referencia>

                        string lsql1 = "'<?xml version=" + '\u0022' + "1.0" + '\u0022' + " encoding=" + '\u0022' + "UTF-8" + '\u0022' + " standalone=" + '\u0022' + "yes" + '\u0022' + "?><Referencia><Documento tipo=" + '\u0022' + "Poliza" + '\u0022' + " edoPago=" + '\u0022' + "0" + '\u0022' +
                            " cadReferencia=" + '\u0022' + "Póliza de Diario, ejercicio: " + poliza.Fecha.Year + ", periodo: " + poliza.Fecha.Month +", número: " + poliza.Numero +", empresa: " + lempresa1 + ", guid: " + poliza.Guid + "." + '\u0022' + "/></Referencia>'" + ",";
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
                              lsql = "insert into AsocCFDIs values (" + i.ToString() + ",ROUND(RAND() * 1000000000,0)"+ ",'" + y.guid + "','" + y.uuid + "',";
                              lsql1 = "'<?xml version=" + '\u0022' + "1.0" + '\u0022' + " encoding=" + '\u0022' + "UTF-8" + '\u0022' + " standalone=" + '\u0022' + "yes" + '\u0022' + "?><Referencia><Documento tipo=" + '\u0022' + "MovimientoPoliza" + '\u0022' + " edoPago=" +
                                  '\u0022' + "0" + '\u0022' + " cadReferencia=" + '\u0022' + "Movimiento de la cuenta: " + y.cuenta + ", empresa: " + lempresa1 + ", guid: " + y.guid + "." + '\u0022' + "/></Referencia>'" + ",";
                            lsql+=lsql1;
                            lsql += "'Contabilidad',1)";
                            mySqlCommand.CommandText = lsql;
                            mySqlCommand.Connection = DbConnection;
                            int iii = mySqlCommand.ExecuteNonQuery();
                               // SqlCommand mySqlCommand1 = new SqlCommand("insert AsocCFDIs values (" + i.ToString() + "," + ii.ToString() + ",'" + y.guid + "','" + y.uuid + ""'", ")" );
                            i++;
                            //ii++;
                        
                        }
                        lsql = "UPDATE Counters Set Next = " + i + " WHERE Name = 'Id_AsocCFDI'";
                        mySqlCommand.CommandText = lsql;
                        mySqlCommand.Connection = DbConnection;
                        int iiiii = mySqlCommand.ExecuteNonQuery();
                        
                        DbConnection.Close() ;
                    }
                    lcuantos++;
                   // MessageBox.Show("Poliza " + _Poliza.Folio.ToString().Trim() + " ya existe");
                }
                else
                {
                //    MessageBox.Show(poliza.UltimoMsjError);
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
            if (Cadenaconexion == "")
                ciCompanyList11.Populate(Cadenaconexion);
            else
            {
                Form4 x = new Form4();
                x.Show();
            }
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
            string error = mLlenarPolizasGomar();
            //string error = mLlenarPolizasGranVision();
            if (error != "")
            {
                MessageBox.Show(error);
                return;
            }
            //mLlenarPolizasGranVision(); 
            mGrabarPolizas(1);
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
                x.asignaformGomar(this);
                x.Show();
            }
            if (Archivo != "")
                botonExcel1.mSetNombre(Archivo);
            this.Text = " Interfaz Gomar Contabilidad " + this.ProductVersion;
            //this.Text = " Interfaz Gran Vision Contabilidad " + this.ProductVersion;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Properties.Settings.Default.archivo = botonExcel1.mRegresarNombre();
            Properties.Settings.Default.Save();
            sesion.cierraEmpresa();
            sesion.finalizaConexion();
        }

        private void botonExcel1_Load(object sender, EventArgs e)
        {

        }
           
            
                
        
    }
}

        

        
    

