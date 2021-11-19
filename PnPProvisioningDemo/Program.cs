using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using PnP.Core.Services;
using PnP.Framework;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using System;
using System.Threading.Tasks;

namespace PnPProvisioningDemo
{
  class Program
  {
    private static ProvisioningTemplate srcTemplate;
    static async Task Main(string[] args)
    {
      var host = Host.CreateDefaultBuilder()
        .ConfigureLogging((hostingContext, logging) =>
        {
          logging.AddEventSourceLogger();
          logging.AddConsole();
        })
            .ConfigureServices((hostingContext, services) =>
            {
              // Read the custom configuration from the appsettings.<environment>.json file
              var customSettings = new CustomSettings();
              hostingContext.Configuration.Bind("CustomSettings", customSettings);

              // Add the PnP Core SDK services
              services.AddPnPCore(options =>
              {
                options.Sites.Add("DemoSite",
                    new PnP.Core.Services.Builder.Configuration.PnPCoreSiteOptions
                    {
                      SiteUrl = customSettings.DemoSiteUrl
                    });
                options.Sites.Add("DemoTargetSite",
                    new PnP.Core.Services.Builder.Configuration.PnPCoreSiteOptions
                    {
                      SiteUrl = customSettings.DemoTargetSiteUrl
                    });
              });

              services.AddPnPCoreAuthentication(
                      options =>
                      {
                        options.Credentials.Configurations.Add("interactive",
                            new PnP.Core.Auth.Services.Builder.Configuration.PnPCoreAuthenticationCredentialConfigurationOptions
                            {
                              ClientId = customSettings.ClientId,
                              TenantId = customSettings.TenantId,
                              Interactive = new PnP.Core.Auth.Services.Builder.Configuration.PnPCoreAuthenticationInteractiveOptions
                              {
                                RedirectUri = customSettings.RedirectUri
                              }
                            });

                        // Configure the default authentication provider
                        options.Credentials.DefaultConfiguration = "interactive";

                        // Map the site defined in AddPnPCore with the 
                        // Authentication Provider configured in this action
                        options.Sites.Add("DemoSite",
                              new PnP.Core.Auth.Services.Builder.Configuration.PnPCoreAuthenticationSiteOptions
                              {
                                AuthenticationProviderName = "interactive"
                              });
                        options.Sites.Add("DemoTargetSite",
                            new PnP.Core.Auth.Services.Builder.Configuration.PnPCoreAuthenticationSiteOptions
                            {
                              AuthenticationProviderName = "interactive"
                            });
                      });
            })
            // Let the builder know we're running in a console
            .UseConsoleLifetime()
            // Add services to the container
            .Build();

      await host.StartAsync();

      using (var scope = host.Services.CreateScope())
      {
        var pnpContextFactory = scope.ServiceProvider.GetRequiredService<IPnPContextFactory>();

        using (var context = await pnpContextFactory.CreateAsync("DemoSite"))
        {          
          using (ClientContext csomContext = PnPCoreSdk.Instance.GetClientContext(context))
          {
            // Use CSOM to load the web
            csomContext.Load(csomContext.Web, p => p.Title);
            csomContext.ExecuteQueryRetry();
            srcTemplate = getProvisioningTemplate(csomContext.Web);
          }
          XMLTemplateProvider provider =
            new XMLFileSystemTemplateProvider(@"C:\temp\PnPProvisioningDemo", "");
          FileSystemConnector fsConnector = new FileSystemConnector(@"C:\temp\PnPProvisioningDemo", "");
          string templateName = "prov_template_lda_demo.xml";
          provider.SaveAs(srcTemplate, templateName);
        }

        using (var context = await pnpContextFactory.CreateAsync("DemoTargetSite"))
        {
          using (ClientContext csomContext = PnPCoreSdk.Instance.GetClientContext(context))
          {
            csomContext.Load(csomContext.Web, p => p.Title);
            csomContext.ExecuteQueryRetry();
            applyProvisioningTemplate(csomContext.Web, srcTemplate);
          }
        }
        Console.ReadLine();
      }
    }

    private static ProvisioningTemplate getProvisioningTemplate(Web web)
    {
      var ptci = new ProvisioningTemplateCreationInformation(web);
      ptci.PersistBrandingFiles = false;
      ptci.IncludeSearchConfiguration = true;
      ptci.PersistBrandingFiles = false;

      ptci.HandlersToProcess = Handlers.Fields | Handlers.ContentTypes | Handlers.SiteSecurity;

      ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
      {
        Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
      };
      ProvisioningTemplate template = web.GetProvisioningTemplate(ptci);

      foreach (PnP.Framework.Provisioning.Model.Field fld in template.SiteFields)
      {
        if (fld.SchemaXml.Contains("Group=\"0 Customer Group\""))
        {
          Console.WriteLine(fld.SchemaXml);
        }
      }
      template.SiteFields.RemoveAll(f => !f.SchemaXml.Contains("Group=\"0 Customer Group\""));

      foreach (PnP.Framework.Provisioning.Model.ContentType ct in template.ContentTypes)
      {
        if (ct.Group.Contains("0 Customer"))
        {
          Console.WriteLine(ct.Name);
        }
      }
      template.ContentTypes.RemoveAll(c => !c.Group.Contains("0 Customer"));

      return template;
    }

    private static void applyProvisioningTemplate(Web web, ProvisioningTemplate template)
    {
      ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation();

      ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
      {
        Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
      };

      web.ApplyProvisioningTemplate(template, ptai);
    }
  }
}
