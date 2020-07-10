using System;
using System.Windows.Forms;

namespace Parte_Diario
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmPrincipal()); //frmPrincipal());//Emailsender() );//  ParteDiario());
        }
    }
}
