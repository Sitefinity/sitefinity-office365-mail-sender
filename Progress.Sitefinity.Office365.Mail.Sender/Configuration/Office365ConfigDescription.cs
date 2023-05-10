using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telerik.Sitefinity.Localization;

namespace Progress.Sitefinity.Office365.Mail.Sender.Configuration
{
    /// <summary>
    /// Represents the string resources for the backend configuration of the Office 365 mail sender.
    /// </summary>
    [ObjectInfo("Office365ConfigDescription", ResourceClassId = "Office365ConfigDescription")]
    public class Office365ConfigDescription : Resource
    {
        /// <summary>
        /// Message: Azure tenant ID
        /// </summary>
        [ResourceEntry("TenantID",
            Value = "Azure Active Directory tenant ID",
            Description = "Describes configuration element.",
            LastModified = "2023/05/09")]
        public string TenantId => base["TenantId"];

        /// <summary>
        /// Message: Application (client) ID
        /// </summary>
        [ResourceEntry("ClientId",
            Value = "Application (client) ID",
            Description = "Describes configuration element.",
            LastModified = "2023/05/09")]
        public string ClientId => base["ClientId"];

        /// <summary>
        /// Message: Client secret value
        /// </summary>
        [ResourceEntry("ClientSecret",
            Value = "Client secret value",
            Description = "Describes configuration element.",
            LastModified = "2023/05/09")]
        public string ClientSecret => base["ClientSecret"];

        /// <summary>
        /// Message: Microsoft Graph scopes (multiple comma-separated ones can be specified)
        /// </summary>
        [ResourceEntry("Scopes",
            Value = "Microsoft Graph scopes (multiple comma-separated ones can be specified)",
            Description = "Describes configuration element.",
            LastModified = "2023/05/09")]
        public string Scopes => base["Scopes"];
    }
}
