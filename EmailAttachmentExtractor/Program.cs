using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MailKit;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.IO;
using System.Configuration;
using Org.BouncyCastle.Asn1.X509;

namespace EmailAttachmentExtractor
{
	internal class Program
	{
		static void Main(string[] args)
		{

			// Get constant from app.config file
			var appSettings = ConfigurationManager.AppSettings;
			string IMAP_USERNAME = appSettings["ImapUsername"];
			string IMAP_PASSWORD = appSettings["ImapPassword"];
			string IMAP_SERVER = appSettings["ImapServer"];
			int IMAP_PORT = Convert.ToInt32(appSettings["ImapPort"]);
			string FOLDER = appSettings["Folder"];
			string SUBJECT = appSettings["Subject"];

            using (var client = new ImapClient())
			{
				client.Connect(IMAP_SERVER, IMAP_PORT, SecureSocketOptions.SslOnConnect);
				client.Authenticate(IMAP_USERNAME, IMAP_PASSWORD);
                IMailFolder mailFolder = client.GetFolder(FOLDER);
				mailFolder.Open(FolderAccess.ReadOnly); //Reads from specific folder
				//client.Inbox.Open(FolderAccess.ReadOnly);

                var uids = mailFolder.Search(SearchQuery.SubjectContains(SUBJECT));

				foreach (var uid in uids)
				{
					var message = mailFolder.GetMessage(uid);

					foreach (var attachment in message.Attachments)
					{
						using (var stream = File.Create("test\\" + attachment.ContentDisposition.FileName))
						{
							if (attachment is MessagePart)
							{
								//var part = (MessagePart)attachment;
								//part.Message.WriteTo(stream);
							}
							else
							{
								var part = (MimePart)attachment;
								part.Content.DecodeTo(stream);
							}
						}
					}
				}
				client.Disconnect(true);
			}
		}
	}
}
