using System;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;

namespace App
{
  class Program
  {
    static async Task Main(string[] args)
    {
      LogInfo("Starting App");

      try
      {
        var config = LoadAppSettings();

        var authProvider = CreateAuthProvider(config);

        var graphClient = new GraphServiceClient(authProvider);

        var siteUri = new Uri(config.SiteUrl);
        var site = await graphClient.Sites.GetByPath(siteUri.AbsolutePath, siteUri.Host)
                                        .Request()
                                        .Select(s => new { s.Id })
                                        .GetAsync();
        LogInfo($"Site ID: {site.Id}");

        var items = await graphClient.Sites[site.Id]
                                        .Lists[config.ListTitle]
                                        .Items
                                        .Request()
                                        .Expand(i => i.Fields)
                                        .Top(10)
                                        .GetAsync();

        foreach (var item in items) 
        {
            LogInfo(string.Format("[{0}] {1}", config.ListTitle, item.Fields.AdditionalData["Title"]));
        }

      }
      catch (Exception exception)
      {
        LogError("Uncaught Application Exception", exception);
      }

    }

    #region Authentication
    private static ClientCredentialProvider CreateAuthProvider(AppConfig config)
    {
      var authority = $"https://login.microsoftonline.com/{config.TenantId}/v2.0";
      var cert = GetCertificateFromStore(config.CertificateThumbprint);

      var msalClient = ConfidentialClientApplicationBuilder.Create(config.ClientId)
                                              .WithCertificate(cert)
                                              .WithAuthority(authority)
                                              .Build();

      return new ClientCredentialProvider(msalClient);
    }
    #endregion

    #region App Config
    private static AppConfig LoadAppSettings()
    {
      AppConfig appConfig = new AppConfig();
      try
      {
        var userSecrets = new ConfigurationBuilder()
            .AddUserSecrets<Program>()
            .Build();

        // Core Authentication and Site Configuration
        appConfig.ClientId = userSecrets["ClientId"];
        appConfig.TenantId = userSecrets["TenantId"];
        appConfig.SiteUrl = userSecrets["SiteUrl"];
        appConfig.CertificateThumbprint = userSecrets["CertificateThumbprint"];

        // Custom Configuration
        appConfig.ListTitle = userSecrets["ListTitle"];
      }
      catch (Exception ex)
      {
        LogError("Unable to load app configuration", ex);
        appConfig = null;
      }

      return appConfig;
    }

    private static X509Certificate2 GetCertificateFromStore(string thumbprint, StoreName storeName = StoreName.My, StoreLocation storeLocation = StoreLocation.CurrentUser)
    {

      X509Store store = new X509Store(storeName, storeLocation);
      try
      {
        store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
        X509Certificate2Collection certificates = store.Certificates.Find(
            X509FindType.FindByThumbprint, thumbprint, false);
        if (certificates.Count == 1)
        {
          return certificates[0];
        }
        else
        {
          return null;
        }
      }
      finally
      {
        store.Close();
      }
    }
    #endregion

    #region Logging Helpers
    private static void LogError(string message, Exception exception = null)
    {
      Console.ForegroundColor = ConsoleColor.Red;
      Console.WriteLine($"[{String.Format("{0:u}", DateTime.Now)}] {message}");
      if (exception != null)
      {
        Console.WriteLine(exception);
      }
      Console.ResetColor();
    }

    private static void LogWarning(string message)
    {
      Console.ForegroundColor = ConsoleColor.Yellow;
      Console.WriteLine($"[{String.Format("{0:u}", DateTime.Now)}] {message}");
      Console.ResetColor();
    }

    private static void LogInfo(string message)
    {
      Console.ForegroundColor = ConsoleColor.White;
      Console.WriteLine($"[{String.Format("{0:u}", DateTime.Now)}] {message}");
      Console.ResetColor();
    }
    #endregion
  }
}
