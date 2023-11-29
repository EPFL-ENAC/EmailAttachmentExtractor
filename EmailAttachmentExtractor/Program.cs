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
			string SOURCEFOLDER = appSettings["SourceFolder"];
			string DESTINATIONFOLDER = appSettings["DestinationFolder"];
			string SUBJECT = appSettings["Subject"];

            using (var client = new ImapClient())
			{
				client.Connect(IMAP_SERVER, IMAP_PORT, SecureSocketOptions.SslOnConnect);
				client.Authenticate(IMAP_USERNAME, IMAP_PASSWORD);
                IMailFolder sourceFolder = client.GetFolder(SOURCEFOLDER);
				IMailFolder destinationFolder = client.GetFolder(DESTINATIONFOLDER);
				sourceFolder.Open(FolderAccess.ReadWrite);//Reads from specific folder

                var uids = sourceFolder.Search(SearchQuery.SubjectContains(SUBJECT)); //Only searches email with specific subject

				foreach (var uid in uids)
				{
					var message = sourceFolder.GetMessage(uid);

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
                var uidMap = sourceFolder.MoveTo(uids, destinationFolder);
                client.Disconnect(true);
			}
		}
	}
}
