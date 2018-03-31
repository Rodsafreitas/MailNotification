using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace MailNotification.Services
{
    public class ExchangeMail
    {
        public String Address { get; private set; }
        public String Password { get; private set; }
        public String ExchangeUrl { get; private set; }
        public DateTime Time { get; set; }

        private ExchangeService exchangeService;
        
        public ExchangeMail(String Address, String Password, String ExchangeUrl)
        {
            this.Address = Address;
            this.Password = Password;
            this.ExchangeUrl = ExchangeUrl;
        }
        
        public void Authentication()
        {
            exchangeService = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            exchangeService.Credentials = new NetworkCredential(this.Address, this.Password);
            exchangeService.Url = new Uri(this.ExchangeUrl);
            exchangeService.KeepAlive = true;
        }
        public FindItemsResults<Item> FetchMail()
        {
            Mailbox mailbox = new Mailbox(this.Address, "");
            FolderId folder = new FolderId(WellKnownFolderName.Inbox, mailbox);
            ItemView view = new ItemView(10);

            return exchangeService.FindItems(folder, view);
        }

    }
}
