using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telerik.Sitefinity.Localization;

namespace Progress.Sitefinity.Office365.Mail.Sender.Configuration
{
    [ObjectInfo("Office365ConfigDescription", ResourceClassId = "Office365ConfigDescription")]
    public class Office365ConfigDescription : Resource
    {
        [ResourceEntry("DefaultTenantID", Value = "Default Azure tenant ID", Description = "Property Caption.", LastModified = "2023/05/09")]
        public string DefaultTenantId => base["DefaultTenantId"];

        [ResourceEntry("DefaultClientId", Value = "Default application (client) ID", Description = "Property Caption.", LastModified = "2023/05/09")]
        public string DefaultClientId => base["DefaultClientId"];

        [ResourceEntry("DefaultClientSecret", Value = "Default client secret value", Description = "Property Caption.", LastModified = "2023/05/09")]
        public string DefaultClientSecret => base["DefaultClientSecret"];

        [ResourceEntry("DefaultScopes", Value = "Default Microsoft Graph scopes (multiple comma-separated ones can be specified)", Description = "Property Caption.", LastModified = "2023/05/09")]
        public string DefaultScopes => base["DefaultScopes"];
    }
}
