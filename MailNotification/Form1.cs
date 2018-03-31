using MailNotification.Helpers;
using MailNotification.Services;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Net;
using System.Threading;
using System.Windows.Forms;

namespace MailNotification
{
    public partial class Form1 : Form
    {
        public ExchangeMail mail;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            this.WindowState = FormWindowState.Minimized;

            notifyIcon1.Icon = SystemIcons.Application;

            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (mail == null)
            {
                mail = new ExchangeMail(ConfigHelper.EmailAddress,
                                        ConfigHelper.EmailPassword,
                                        ConfigHelper.ExchangeUrl);

                //Authentication mail
                mail.Authentication();

                //Emails from the last three minutes
                mail.Time = DateTime.Now.AddMinutes(-3);

            }

            FindItemsResults<Item> item = mail.FetchMail();

            //Call method show to pass to notify 
            Show(MailList.ListMail(item, mail.Time));

        }
        private void Show(List<Item> itemFilter)
        {
            //Show on notify
            foreach (Item item in itemFilter)
            {
                notifyIcon1.BalloonTipTitle = $"New mail from : {item.LastModifiedName}";
                notifyIcon1.BalloonTipText = item.Subject;
                notifyIcon1.ShowBalloonTip(0);

                Thread.Sleep(1000);
            }
        }
    }
}
