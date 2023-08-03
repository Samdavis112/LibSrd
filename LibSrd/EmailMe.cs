using System.Net;
using System.Net.Mail;

namespace LibSrd
{
    public class EmailMe
    {
        private static string smtp_server = "smtp.gmail.com";
        private static string Email = SecretVariables.senderEmail;
        private static string Password = SecretVariables.senderPassword;

        public static void SendEmail(string subject, string recipient, string body)
        {
            //Initialise message instance
            var smtpClient = new SmtpClient(smtp_server)
            {
                Credentials = new NetworkCredential(Email, Password),
                EnableSsl = true,
            };

            smtpClient.Send(Email, recipient, subject, body);
        }
    }
}
