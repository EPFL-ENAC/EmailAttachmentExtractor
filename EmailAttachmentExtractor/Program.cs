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

namespace EmailAttachmentExtractor
{
	internal class Program
	{
		static void Main(string[] args)
		{

			// Get constant from app.config file
			var appSettings = ConfigurationManager.AppSettings;
			string IMAP_USERNAME = appSettings["ImapUsername"];//.ToString();
			string IMAP_PASSWORD = appSettings["ImapPassword"].ToString();
			string IMAP_SERVER = appSettings["ImapServer"].ToString();
			int IMAP_PORT = Convert.ToInt32(appSettings["ImapPort"].ToString());

			using (var client = new ImapClient())// new ProtocolLogger("imap.log")))
			{
				client.Connect(IMAP_SERVER, IMAP_PORT, SecureSocketOptions.SslOnConnect);
				client.Authenticate(IMAP_USERNAME, IMAP_PASSWORD);
				client.Inbox.Open(FolderAccess.ReadOnly);

				var uids = client.Inbox.Search(SearchQuery.All);

				foreach (var uid in uids)
				{
					var message = client.Inbox.GetMessage(uid);

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
