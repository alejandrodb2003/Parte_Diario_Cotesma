using System;
using System.Net.Mail;
using System.Windows.Forms;

namespace Parte_Diario
{
    public class Email
    {
        private readonly MailMessage email;
        private readonly SmtpClient smtp;
        public Email()
        {
            email = new MailMessage();
            smtp = new SmtpClient();
            smtp.Host = "correo.cotesma.com.ar";
            smtp.Port = 993;
            smtp.EnableSsl = true;
            
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new System.Net.NetworkCredential("alejandro.barboza@cotesma.com.ar", "camila5@emanuel");
        }

        public void Mensaje()
        {

            email.To.Add(new MailAddress("alejandro.barboza@cotesma.com.ar"));
            email.From = new MailAddress("alejandrobarboza@gmail.com");
            email.Subject = "Asunto ( " + DateTime.Now.ToString("dd / MMM / yyy hh:mm:ss") + " ) ";
            email.Body = "<p style='font-weight: 400;'><strong>Alejandro Daniel Barboza</strong><br />" +
                         "<strong><em>Soporte t&eacute;cnico domiciliario<span>&nbsp;</span></em></strong><strong>&nbsp;</strong></p>" +
                         "<p style='font-weight: 400;'>Tel. +54 2972 427000 | Cel. +549 11 6509 8678</p>" +
                         "<p style='font-weight: 400;'><span><a href='mailto:alejandro.barboza@cotesma.com.ar'>" +
                         "alejandro.barboza@cotesma.com.ar</a></span><span>&nbsp;</span>|<span>&nbsp;</span><span>" +
                         "<a href='https://www.cotesma.coop/'>cotesma.coop</a></span></p>" +
                         "<p style='font-weight: 400;'><span><img src='blob:https://www.cotesma.com.ar/2a781d78-1414-453b-9bfd-dd3d5526d1c3'" +
                         " width='257' height='39' /></span></p>";
            //email.Attachments.Add();
            email.IsBodyHtml = true;

            email.Priority = MailPriority.Normal;

        }

        public int enviar()
        {
            int rta = 0;


            //string output = null;
            try
            {
                smtp.Send(email);
                email.Dispose();
                MessageBox.Show("Corre electrónico fue enviado satisfactoriamente.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error enviando correo electrónico: " + ex.Message);
            }

            return rta;
        }
    }
}
