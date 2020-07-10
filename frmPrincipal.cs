using System;
using System.Windows.Forms;
namespace Parte_Diario
{
    public partial class frmPrincipal : Form
    {

        public frmPrincipal()
        {
            InitializeComponent();


        }

        private void nuevoParteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ParteDiario parte2 = new ParteDiario();
            // Set the Parent Form of the Child window.
            parte2.MdiParent = this;
            // Display the new form.

            parte2.Show();

        }

        private void configuraciónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmConfiguracion config = new frmConfiguracion();

            config.MdiParent = this;
            config.Show();
        }
    }
}
