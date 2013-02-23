using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

using Microsoft.Exchange.WebServices.Data;
using System.Diagnostics;

namespace OutlookTaskExtract
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Impersonation
            // http://msdn.microsoft.com/en-us/library/dd633680%28v=exchg.80%29.aspx

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
        }

        private void extract_Click(object sender, EventArgs e)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            service.Credentials = new WebCredentials("rassad@gbhem.org", "Question96");
            service.AutodiscoverUrl("rassad@gbhem.org", RedirectionUrlValidationCallback);

            var items = service.FindItems(WellKnownFolderName.Tasks, 
                                            new ItemView(100) { PropertySet = new PropertySet(BasePropertySet.FirstClassProperties) });

            PropertySet _customPropertySet = new PropertySet(BasePropertySet.FirstClassProperties, AppointmentSchema.MyResponseType, AppointmentSchema.IsMeeting, AppointmentSchema.ICalUid);
            _customPropertySet.RequestedBodyType = BodyType.Text;
            
            foreach (Task task in items.Items)
            {
                task.Load(_customPropertySet);
                String subject = task.Subject;
                DateTime? startDate = task.StartDate;
                DateTime? dueDate = task.DueDate;
                Boolean reminderOn = task.IsReminderSet;
                DateTime reminderDate;
                try
                {
                    reminderDate = task.ReminderDueBy;
                }
                catch { }
                Int32 reminderMinutes = task.ReminderMinutesBeforeStart;
                DateTime? dateCompleted = task.CompleteDate;
                Double percentComplete = task.PercentComplete;
                Int32? totalWork = task.TotalWork;
                Int32? acutalWork = task.ActualWork;
                String billingInfo = task.BillingInformation;
                StringList categories = task.Categories;
                StringList contacts = task.Contacts;
                String mileage = task.Mileage;
                String notes = task.Body;
                String sensitivity = task.Sensitivity.ToString();
                String status = task.Status.ToString();

                String priority = task.Importance.ToString();
                Boolean isPrivate = (sensitivity.ToLower() == "private");
            }
        }

        private static bool CertificateValidationCallBack(object sender,
                                                            System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                                                            System.Security.Cryptography.X509Certificates.X509Chain chain,
                                                            System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}
