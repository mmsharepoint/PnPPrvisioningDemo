using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using PnP.Framework;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;


namespace PnPProvisioningLookupDemo
{
  class Program
  {
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

        ProvisioningTemplate srcTemplate = new ProvisioningTemplate();
        using (var context = await pnpContextFactory.CreateAsync("DemoSite"))
        {      
          var locationlist = await context.Web.Lists.GetByTitleAsync("Locations");
          List<Model.Location> locations = new List<Model.Location>();
          await foreach (var listItem in locationlist.Items)
          {
            Model.Location loc = new Model.Location();
            loc.Title = listItem.Title;
            loc.IDold = listItem.Id;
            locations.Add(loc);
          }
          ListInstance list = new ListInstance();
          list.Title = "Locations";
          list.Url = "Lists/Locations";
          list.TemplateType = 100;
          list.TemplateFeatureID = new Guid("00bfea71-de22-43b2-a848-c05709900100");
          list.DataRows.KeyColumn = "Title";
          list.DataRows.UpdateBehavior = UpdateBehavior.Overwrite;
          for(int i=1; i <= locations.Count; i++)
          {
            DataRow row = new DataRow();            
            row.Values.Add("Title", locations[i-1].Title);
            locations[i - 1].IDnew = i;
            list.DataRows.Add(row);
          }
          srcTemplate.Lists.Add(list);
          //DataRow row = new DataRow();
          //row.Values.Add("Title", "Claus Clausen");
          //row.Values.Add("Firstname", "Claus");
          //row.Values.Add("Lastname", "Clausen");
          //row.Values.Add("Street", "Stresemannstrasse");
          //row.Values.Add("StreetNo", "15");
          //row.Values.Add("EmployeeLocation", "1");
          //row.Values.Add("Salary", "2250");
          //list.DataRows.Add(row);
          var employeelist = await context.Web.Lists.GetByTitleAsync("Employees");
          List<Model.Employee> employees = new List<Model.Employee>();
          await foreach (var listItem in employeelist.Items)
          {
            Model.Employee emp = new Model.Employee();
            emp.Title = listItem.Title;
            emp.Firstname = listItem["Firstname"].ToString();
            emp.Lastname = listItem["Lastname"].ToString();
            emp.Street = listItem["Street"].ToString();
            emp.StreetNo = listItem["StreetNo"].ToString();
            emp.Salary = listItem["Salary"].ToString();
            var x = listItem["EmployeeLocation"] as IFieldLookupValue;
            emp.OldEmployeeLocation = x.LookupId;
            employees.Add(emp);
          }          
          ListInstance list2 = new ListInstance();
          list2.Title = "Employees";
          list2.Url = "Lists/Employees";
          list2.TemplateType = 100;
          list2.TemplateFeatureID = new Guid("00bfea71-de22-43b2-a848-c05709900100");
          for (int i = 1; i <= employees.Count; i++)
          {
            DataRow row = new DataRow();
            row.Values.Add("Title", employees[i - 1].Title);
            row.Values.Add("Firstname", employees[i - 1].Firstname);
            row.Values.Add("Lastname", employees[i - 1].Lastname);
            row.Values.Add("Street", employees[i - 1].Street);
            row.Values.Add("StreetNo", employees[i - 1].StreetNo);
            row.Values.Add("Salary", employees[i - 1].Salary);
            Model.Location refloc = locations.FindLast(l => l.IDold == employees[i - 1].OldEmployeeLocation);
            row.Values.Add("EmployeeLocation", refloc.IDnew.ToString());
            list2.DataRows.Add(row);
          }
          srcTemplate.Lists.Add(list2);

          
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
      }
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
