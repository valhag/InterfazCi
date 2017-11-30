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
        private void button1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog fbDialog;
            fbDialog = new System.Windows.Forms.OpenFileDialog();
            fbDialog.DefaultExt = "xls,xlsx";

            
            fbDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
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
