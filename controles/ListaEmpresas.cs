using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace InterfazCi
{
    public partial class ListaEmpresas : UserControl
    {
        string Cadenaconexion;
        public ListaEmpresas(string aCadenaConexion)
        {
            Cadenaconexion = aCadenaConexion;
            InitializeComponent();
        }
        public ListaEmpresas()
        {
            InitializeComponent();
        }

        private void ListaEmpresas_Load(object sender, EventArgs e)
        {
            DataTable Empresas = null;
            mlistarEmpresas(ref Empresas);
            if (Empresas != null)
            {
                mllenaList(Empresas);
            }
            else
            {
                MessageBox.Show("Es necesario que configure correctamente los datos de la configuracion de la conexion a sqlserver");
            }
        }

        private void mllenaList(DataTable Empresas)
        {

            comboBox1.Items.Clear();
            comboBox1.DataSource = Empresas;
            comboBox1.DisplayMember = "NombreEmpresa";
            comboBox1.ValueMember = "RutaEmpresa";

        }

        private void mlistarEmpresas(ref DataTable Empresas)
        {
            SqlConnection DbConnection = new SqlConnection(Cadenaconexion);


            SqlCommand mySqlCommand = new SqlCommand("select nombreempresa,rutaempresa from NOM10000", DbConnection);
            DataSet ds = new DataSet();
            //mySqlCommand.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter();
            mySqlDataAdapter.SelectCommand = mySqlCommand;

            try
            {
                mySqlDataAdapter.Fill(ds);
                Empresas = ds.Tables[0];

            }
            catch (Exception ee)
            {

            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
