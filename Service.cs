using Microsoft.Exchange.WebServices.Data;
using System;

namespace WindowsFormsApp1
{
    internal class Service
    {
        private ExchangeService mService;
        public Service()
        {
            this.mService = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            this.mService.Credentials = new WebCredentials("devin@denaliai.com", "Welcome2018");
            this.mService.UseDefaultCredentials = false;
            this.mService.TraceEnabled = true;
            this.mService.TraceFlags = TraceFlags.All;
            this.mService.AutodiscoverUrl("devin@denaliai.com", RedirectionUrlValidationCallback);
        }
        public ExchangeService getMService() { return mService; }

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
