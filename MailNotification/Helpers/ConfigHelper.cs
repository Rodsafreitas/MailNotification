using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailNotification.Helpers
{
    public class ConfigHelper
    {
       public static string EmailAddress { get { return ConfigurationManager.AppSettings["EmailAddress"].ToString(); } }
        public static string EmailPassword { get { return ConfigurationManager.AppSettings["EmailPassword"].ToString(); } }
        public static string ExchangeUrl { get { return ConfigurationManager.AppSettings["ExchangeUrl"].ToString(); } }
    }
}
