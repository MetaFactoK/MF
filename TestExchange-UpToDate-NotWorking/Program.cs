using Azure.Identity;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestExchange_OldVersion_Working
{
	class Program
	{
		static string tenantId = "[TENANT_ID]";
		static string clientId = "[CLIENT_ID]";
		static string clientSecret = "[CLIENT_SECRET]";
		static string sMailbox = "[MAIL]";

		static async Task Main(string[] args)
		{
			await testInbox();
		}

		static async Task test()
		{
			Console.WriteLine("Begin test");
			try
			{
				var options = new TokenCredentialOptions
				{
					AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
				};

				string[] scopes = { "https://graph.microsoft.com/.default" };

				ClientSecretCredential csc = new ClientSecretCredential(tenantId, clientId, clientSecret, options);

				GraphServiceClient m_graphClient = new GraphServiceClient(csc, scopes);

				var user = await m_graphClient.Users[sMailbox].Request().GetAsync();

				Console.WriteLine(user.DisplayName);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			Console.WriteLine("End test");
		}

		static async Task testInbox()
		{
			Console.WriteLine("Begin test");
			try
			{
				var options = new TokenCredentialOptions
				{
					AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
				};

				string[] scopes = { "https://graph.microsoft.com/.default" };

				ClientSecretCredential csc = new ClientSecretCredential(tenantId, clientId, clientSecret, options);

				GraphServiceClient m_graphClient = new GraphServiceClient(csc, scopes);

				var inbox = await m_graphClient.Users[sMailbox].MailFolders.Inbox.Request().GetAsync();

				Console.WriteLine(inbox.DisplayName);

				var inboxMsg = await m_graphClient.Users[sMailbox].MailFolders.Inbox.Messages.Request().GetAsync();

				Console.WriteLine(inboxMsg.Count + "msg");
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			Console.WriteLine("End test");
		}
	}
}
