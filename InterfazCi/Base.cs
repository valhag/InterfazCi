using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using InterfazCi;

namespace InterfazCi
{
    public partial class Base : Form
    {

        public string Cadenaconexion = "";
        public string Cadenaconexionsap = "";

        public string Archivo = "";

        protected List<Poliza> _RegPolizas = new List<Poliza>();
        protected Poliza _Poliza = new Poliza();

        

        public Base()
        {
            InitializeComponent();
            if (Properties.Settings.Default.server != "")
            {
                Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";



                Archivo = Properties.Settings.Default.archivo;

                //                    botonExcel1.mSetNombre(System.IO.Directory.GetCurrentDirectory() + "\\ArchivoExportar.txt");
            }
            if (Properties.Settings.Default.serverSAP != "")
            {
                Cadenaconexionsap = "data source =" + Properties.Settings.Default.serverSAP +
               ";initial catalog =" + Properties.Settings.Default.dbSAP + " ;user id = " + Properties.Settings.Default.userSAP +
               "; password = " + Properties.Settings.Default.pwdSAP + ";";
            }
        }

        public void mOcultaTab3()
        {
            tabControl1.TabPages.Remove(tabPage3);
        }

        private void Base_Load(object sender, EventArgs e)
        {
            txtServer.Text = Properties.Settings.Default.server;
            //txtBD.Text = Properties.Settings.Default.database;
            txtUser.Text = Properties.Settings.Default.user;
            txtPass.Text = Properties.Settings.Default.password;

            TxtServerSAP.Text = Properties.Settings.Default.serverSAP;
            txtBDSAP.Text = Properties.Settings.Default.dbSAP;
            txtUserSAP.Text = Properties.Settings.Default.userSAP;
            txtpwdSAP.Text = Properties.Settings.Default.pwdSAP;


            //this.ciCompanyList11.SelectedItem += new EventHandler(OnComboChange);

           // this.ciCompanyList12.SelectedItem += new EventHandler(OnComboChange);
        }

        private void OnComboChange(object sender, EventArgs e)
        {

            MessageBox.Show("uno");

        }

        private void Base_Shown(object sender, EventArgs e)
        {
            if (Cadenaconexion != "")
            {
                ciCompanyList11.Populate(Cadenaconexion);
                ciCompanyList12.Populate(Cadenaconexion);
            }
            else
            {
                /*this.Visible = false;
                Form4 x = new Form4();
                //x.asignaformGeneratxt(this);
                x.Show();*/
                
                    MessageBox.Show("Configure Conexion");

            }

            if (Cadenaconexionsap == "")
            {
                if (TxtServerSAP.Visible == true)
                    MessageBox.Show("Configure Conexion sap");

            }
            // if (Archivo != "")


        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (mValida())
            {
                Properties.Settings.Default.server = txtServer.Text;
                Properties.Settings.Default.database = "GeneralesSQL";
                Properties.Settings.Default.user = txtUser.Text;
                Properties.Settings.Default.password = txtPass.Text;

                Properties.Settings.Default.Save();

                MessageBox.Show("Conexion Correcta");
                Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";
                mllenarcomboempresas();
                
            }
            else
                MessageBox.Show("Valores de conexion incorrectos");
        }

        public void mllenarcomboempresas()
        {
            ciCompanyList11.Populate(Cadenaconexion);
        }

        private bool mValida()
        {
            string Cadenaconexion = "data source =" + txtServer.Text + ";initial catalog =" + txtBD.Text + ";user id = " + txtUser.Text + "; password = " + txtPass.Text + ";";
            SqlConnection _con = new SqlConnection();

            _con.ConnectionString = Cadenaconexion;
            try
            {
                _con.Open();
                // si se conecto grabar los datos en el cnf
                _con.Close();
                return true;
            }
            catch (Exception ee)
            {
                return false;
            }
        }

        private bool mValidasap()
        {
            string Cadenaconexion = "data source =" + TxtServerSAP.Text + ";initial catalog =" + txtBDSAP.Text + ";user id = " + txtUserSAP.Text + "; password = " + txtpwdSAP.Text + ";";
            SqlConnection _con = new SqlConnection();

            _con.ConnectionString = Cadenaconexion;
            try
            {
                _con.Open();
                // si se conecto grabar los datos en el cnf
                _con.Close();
                return true;
            }
            catch (Exception ee)
            {
                return false;
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (mValidasap())
            {
                Properties.Settings.Default.serverSAP = TxtServerSAP.Text;
                Properties.Settings.Default.dbSAP = txtBDSAP.Text;
                Properties.Settings.Default.userSAP = txtUserSAP.Text;
                Properties.Settings.Default.pwdSAP= txtpwdSAP.Text;

                Properties.Settings.Default.Save();

                MessageBox.Show("Conexion Correcta");
                Cadenaconexionsap = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";
                mllenarcomboempresas();

            }
            else
                MessageBox.Show("Valores de conexion incorrectos");
        }

        private void ciCompanyList11_Load(object sender, EventArgs e)
        {

        }
    }
}
