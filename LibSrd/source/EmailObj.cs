using System.Net;
using System.Net.Mail;

namespace LibSrd
{
    public class EmailObj
    {
        private const string smtp_server = "smtp.gmail.com";
        private string Email;
        private string Password;

        public EmailObj(string email=null, string password=null)
        {
            Email = email;
            Password = password;

            if(Email == null && Password == null)
            {
                Password = SecretVariables.senderPassword;
                Email = SecretVariables.senderEmail;
            }
        }
        public void SendEmail(string subject, string recipient, string body)
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
