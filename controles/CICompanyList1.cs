using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace controles
{
    public partial class CICompanyList1 : UserControl
    {
        string Cadenaconexion="";
        public string aliasbdd = "";
        public event EventHandler SelectedItem;
        public ComboBox micombo = new ComboBox();


        public CICompanyList1(string aCadena)
        {
            InitializeComponent();
            
            Cadenaconexion = aCadena;

        }

      


        public CICompanyList1()
        {
            InitializeComponent();
        }

        private void CICompanyList1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndexChanged += new EventHandler(OnSelectedItem);
        }

        public void OnSelectedItem(object sender, EventArgs e)
        {
            // Delegate the event to the caller
            if (SelectedItem != null)
                SelectedItem(this.comboBox1, e);
        }

        public void mrecorrertxt(int left=0)
        {
            comboBox1.Left = left;
        }
        public void Populate(string aCadena)
        {
            Cadenaconexion = aCadena;
            DataTable Empresas = null;
            mTraerEmpresas(ref Empresas);
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
            if (comboBox1.Items.Count == 0)
            {
                comboBox1.Items.Clear();
                comboBox1.DataSource = Empresas;
                comboBox1.DisplayMember = "nombre";
                comboBox1.ValueMember = "aliasbdd";
            }

        }

        private void mTraerEmpresas(ref DataTable Empresas)
        {
            SqlConnection DbConnection = new SqlConnection(Cadenaconexion);


            SqlCommand mySqlCommand = new SqlCommand("select nombre,aliasbdd from ListaEmpresas", DbConnection);
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

        

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            ComboBox cmb = (ComboBox)sender;
            if (cmb.SelectedIndex != -1)
            {
                int selectedIndex = cmb.SelectedIndex;


                DataRowView selectedCar = (DataRowView)cmb.SelectedItem;
                aliasbdd = selectedCar.Row[1].ToString();
            }

            

        }

        /*public class ComboboxItem
        {
            public string Text { get; set; }
            public string Value { get; set; }
            public override string ToString() { return Text; }
        }*/
    }
}
