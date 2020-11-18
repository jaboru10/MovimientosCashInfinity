using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelResidencia
{
    public partial class Configuracion : MetroFramework.Forms.MetroForm
    {
        public Configuracion()
        {
            InitializeComponent();
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            try
            {
                Properties.Settings.Default.host = metroTextBox1.Text;
                Properties.Settings.Default.Save();
                this.DialogResult = DialogResult.OK;
            }
            catch (Exception)
            {
                MessageBox.Show("Compruebe que el nº introducido en operacion es correcto y que existe algún valor en la casilla de Host.");
                metroTextBox1.Focus();
            }
        }

        private void Configuracion_Load(object sender, EventArgs e)
        {
            metroTextBox1.Text = Properties.Settings.Default.host;
           
        }

        private void metroTextBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                HayConexion();
            }
        }

        private void HayConexion()
        {
            try
            {
                Farmanager.Farmanager far = new Farmanager.Farmanager(metroTextBox1.Text, 3306);
                if (far.IsOpenConnection())
                {
                    MessageBox.Show(String.Format("¡Conexión con {0} correcta!", metroTextBox1.Text));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "¡No hay conexión!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }

        }
    }
}
