async void testReadGraphAPI()
    {
      try
      {
        var options = new TokenCredentialOptions
        {
          AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
        };

        string[] scopes = { "https://graph.microsoft.com/.default" };

        ClientSecretCredential csc = new ClientSecretCredential([MY_TENANT_ID], [MY_CLIENT_ID], [MY_CLIENT_SECRET], options);

        GraphServiceClient graphClient = new GraphServiceClient(csc, scopes);

        var user = await graphClient.Users[sMailbox].Request().GetAsync();

        return;
	  }
      catch (Exception ex)
      {
        Console.WriteLine(ex.Message);
      }
    }