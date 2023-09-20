using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace controles
{
    public partial class BotonExcel : UserControl
    {
        public int tipo = 0;
        public BotonExcel()
        {
            InitializeComponent();
        }
        public string mRegresarNombre()
        {
            return textBox1.Text;
        }

        public void mSetNombre(string aNombre)
        {
            textBox1.Text = aNombre;
        }

        public void mSetEtiqueta(string aNombre)
        {
            label1.Text = aNombre;
        }

        public void mSetTipo(int aTipo)
        {
            tipo = aTipo;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog fbDialog;
            fbDialog = new System.Windows.Forms.OpenFileDialog();


            switch (tipo)
            {
                case 0:
                fbDialog.DefaultExt = "xls,xlsx";
                fbDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                break;
                case 1:
                fbDialog.DefaultExt = "txt";
                fbDialog.Filter = "Txt Files|*.txt";
                break;

            }

            if (fbDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                   textBox1.Text = fbDialog.FileName; 
            }
        }

        private void BotonExcel_Load(object sender, EventArgs e)
        {

        }
    }
}
