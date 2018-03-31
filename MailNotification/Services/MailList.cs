using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MailNotification.Services
{
    public static class MailList
    {
        public static List<Item> lista = new List<Item>();
        public static List<Item> ListMail(FindItemsResults<Item> items, DateTime limitTime)
        {
            lista.Clear();

            foreach (var i in items)
            {
                i.Load();

                if (DateTime.Compare(limitTime, i.DateTimeCreated) <= 0)
                   lista.Add(i);
                
            }
            return lista;
        }

    }
}
