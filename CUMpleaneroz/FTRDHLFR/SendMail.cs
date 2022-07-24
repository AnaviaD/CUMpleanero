using System;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.IO;

namespace CUMpleanero
{
    class SendMail
    {
        bool notificar = true;
        string htmlmsg = "<!DOCTYPE html><html lang=\"en\" xmlns=\"http://www.w3.org/1999/xhtml\">"+
"<head>"+
    "<meta charset=\"utf-8\" />"+
    "<title></title>"+
    "<style> td.data, th.data {border: 1px solid black;}"+
    "</style>"+
"</head>"+
"<body>"+
    "<table style=\"width:100%;border-color:black; margin-bottom: 0px;border-collapse:collapse;border-width:1px;\">"+
        "<tr>"+
            "<td style=\"background-color:black;height:60px;\"><img src=\"http://www.ftr.com.mx/images/logo_alt.png\" style=\"padding-left:20px;\" /></td>"+
        "</tr>"+
        "<tr>"+
            "<td style=\"font-family:Verdana; font-size:13px;padding:10px;\">Mensaje enviado desde FTR.</td>"+
        "</tr>"+
        "<tr><td style=\"font-family:Verdana; font-size:13px;padding:4px;\">{-Parte1-}</td></tr>"+
        "<tr><td style=\"font-family:Verdana; font-size:13px;padding:4px;\">{-Parte2-}</td></tr>"+
        "<tr><td style=\"font-family:Verdana; font-size:13px;padding:4px;\">{-Parte3-}</td></tr>"+
        "<tr>"+
            
        "</tr>"+
        "<tr>"+
            "<td style=\"background-color:black;height:30px;font-family:Verdana;font-size:14px;text-align:center;font-weight:bold;color:white\">Verificacion Carta Porte</td>"+
        "</tr>"+
    "</table>"+

"</body>"+
"</html>";

        public SendMail(bool avisar)
        {
            notificar = avisar;
        }

        public int Notificar(string destinatario, string mensaje, string mensaje2, string mensaje3, string topico, string attachedDoc)
        {
            IDictionary<string, string> integrantes = new Dictionary<string, string>();

            if (!notificar)
                return 0;

            // el contenido de la plantilla
            string plantilla = htmlmsg;
            plantilla = plantilla.Replace("{-Parte1-}", mensaje);
            plantilla = plantilla.Replace("{-Parte2-}", mensaje2);
            plantilla = plantilla.Replace("{-Parte3-}", mensaje3);
            // Proceso normal de correo.
            MailMessage Sendmail = new MailMessage();
            SmtpClient emailClient = new SmtpClient("correo.ftr.com.mx", 25);    //  10.1.1.224    SRVFTREX1.ftr.local
            MailAddress direccionMailFrom = new MailAddress("verificacionFormato@ftr.com.mx");
            Sendmail.From = direccionMailFrom;
            Sendmail.IsBodyHtml = true;
            //Sendmail.IsBodyHtml = false;
            Sendmail.Subject = topico;
            Sendmail.Body = plantilla;
            Sendmail.Priority = MailPriority.High;

            emailClient.EnableSsl = false;
            emailClient.Port = 25;

            Sendmail.To.Add(destinatario);

            Sendmail.To.Add("desaftr02 @ftr.com.mx");
            Sendmail.To.Add("acastillo@ftr.com.mx");
            Sendmail.To.Add("ecarrillo@ftr.com.mx");
            Sendmail.To.Add("desaftr04@ftr.com.mx");
            Sendmail.To.Add("jsalinasd@ftr.com.mx");
            Sendmail.To.Add("jdmeza@ftr.com.mx");
            Sendmail.To.Add("malvarez@ftr.com.mx");
            Sendmail.To.Add("mmorales@ftr.com.mx");
            Sendmail.To.Add("idiaz@ftr.com.mx");
            Sendmail.To.Add("groblero@ftr.com.mx");
            Sendmail.To.Add("vcalderon@ftr.com.mx");
            Sendmail.To.Add("marmas@ftr.com.mx");
            Sendmail.To.Add("rbaltazar@ftr.com.mx");
            
            Sendmail.To.Add("ftrccp@ftr.com.mx");

            try
            {
                System.Net.Mail.Attachment attachmentOri;
                attachmentOri = new System.Net.Mail.Attachment(attachedDoc);
                Sendmail.Attachments.Add(attachmentOri);

                System.Net.Mail.Attachment attachment;


                if (topico.Contains("amazon"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\Amazon-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("apl"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\APL-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("generico"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\Generico-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("adient"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\Generico-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("android"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\Android-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                
                if (topico.Contains("celtic"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\Celtic-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("dhlsupply"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\DHL-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("jbhunt"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\JBH-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("leartoluca"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\Generico-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("matson"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\Generico-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("schneider"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\Schneider-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("stellantis"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\Stellantis-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("stellantis-mts"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\Stellantis-mts-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("transplace"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\830059015-Transplace-EjemploCorrecto.xml");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("tremec"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\Tremec-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("learmexico"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\lear-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("truper"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\Trupper-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("upds"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\DHL-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }
                if (topico.Contains("test"))
                {
                    attachment = new System.Net.Mail.Attachment(@"\\10.1.1.30\FTProot\Formatos\Generico-EjemploCorrecto.xlsx");
                    Sendmail.Attachments.Add(attachment);
                }

            }
            catch (Exception)
            {

                throw;
            }


            

            emailClient.Send(Sendmail);
            Sendmail.Attachments.Dispose();
            Sendmail.Dispose();

            return 0;
        }
    }
}
