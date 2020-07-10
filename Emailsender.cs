using System;
using System.Windows.Forms;

namespace Parte_Diario
{
    public partial class Emailsender : Form
    {
        private Email correo;
        public Emailsender()
        {
            InitializeComponent();
            correo = new Email();
        }

        private void Emailsender_Load(object sender, EventArgs e)
        {
            btnEnviar.Text = "Enviar";
        }

        private void btnEnviar_Click(object sender, EventArgs e)
        {
            correo.Mensaje();
            correo.enviar();
        }
    }
}
