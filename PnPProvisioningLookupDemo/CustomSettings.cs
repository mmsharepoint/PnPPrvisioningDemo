using System;

namespace PnPProvisioningLookupDemo
{
  class CustomSettings
  {
    public string ClientId { get; set; }
    public string TenantId { get; set; }
    public string DemoSiteUrl { get; set; }
    public string DemoTargetSiteUrl { get; set; }
    public Uri RedirectUri { get; set; }
  }
}
