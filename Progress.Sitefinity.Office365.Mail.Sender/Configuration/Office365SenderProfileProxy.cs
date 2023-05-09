using Telerik.Sitefinity.Services.Notifications.Configuration;

namespace Progress.Sitefinity.Office365.Mail.Sender.Configuration
{
    /// <summary>
    /// A proxy class which declares properties required to configure a Office 365 SMTP client
    /// </summary>
    public class Office365SenderProfileProxy : SenderProfileProxy
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Office365SenderProfileProxy" /> class.
        /// </summary>
        public Office365SenderProfileProxy()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Office365SenderProfileProxy" /> class.
        /// </summary>
        /// <param name="senderProfile">The sender profile.</param>
        public Office365SenderProfileProxy(ISenderProfile senderProfile) : base(senderProfile)
        {
        }

        /// <summary>
        /// Gets or sets the default sender email address.
        /// </summary>
        /// <value>The default sender email address.</value>
        public string DefaultSenderEmailAddress
        {
            get
            {
                return this.CustomProperties[Office365Keys.DefaultSenderEmailAddress];
            }

            set
            {
                this.CustomProperties[Office365Keys.DefaultSenderEmailAddress] = value;
            }
        }

        /// <summary>
        /// Gets or sets the default sender name.
        /// </summary>
        /// <value>The the default sender name.</value>
        public string DefaultSenderName
        {
            get
            {
                return this.CustomProperties[Office365Keys.DefaultSenderName];
            }

            set
            {
                this.CustomProperties[Office365Keys.DefaultSenderName] = value;
            }
        }

        /// <summary>
        /// Gets or sets the Azure tenant ID.
        /// </summary>
        /// <value>The Azure tenant ID.</value>
        public string TenantId
        {
            get
            {
                return this.CustomProperties[Office365Keys.TenantId];
            }

            set
            {
                this.CustomProperties[Office365Keys.TenantId] = value;
            }
        }

        /// <summary>
        /// Gets or sets the application (client) ID.
        /// </summary>
        /// <value>The application (client) ID.</value>
        public string ClientId
        {
            get
            {
                return this.CustomProperties[Office365Keys.ClientId];
            }

            set
            {
                this.CustomProperties[Office365Keys.ClientId] = value;
            }
        }

        /// <summary>
        /// Gets or sets the client secret value.
        /// </summary>
        /// <value>The client secret value.</value>
        public string ClientSecret
        {
            get
            {
                return this.CustomProperties[Office365Keys.ClientSecret];
            }

            set
            {
                this.CustomProperties[Office365Keys.ClientSecret] = value;
            }
        }

        /// <summary>
        /// Gets or sets the Microsoft Graph scopes.
        /// </summary>
        /// <value>The Microsoft Graph scopes.</value>
        public string Scopes
        {
            get
            {
                return this.CustomProperties[Office365Keys.Scopes];
            }

            set
            {
                this.CustomProperties[Office365Keys.Scopes] = value;
            }
        }

        /// <summary>
        /// Gets or sets the sender type.
        /// </summary>
        /// <value>The sender type.</value>
        public string SenderType
        {
            get
            {
                return this.CustomProperties[Office365Keys.SenderType];
            }

            set
            {
                this.CustomProperties[Office365Keys.SenderType] = value;
            }
        }

        /// <summary>
        /// Gets or sets the size of the batch.
        /// </summary>
        /// <value>The size of the batch.</value>
        public int BatchSize
        {
            get
            {
                int batchSize;
                int.TryParse(this.CustomProperties[Office365Keys.BatchSize], out batchSize);
                return batchSize;
            }

            set
            {
                this.CustomProperties[Office365Keys.BatchSize] = value.ToString();
            }
        }

        /// <summary>
        /// Gets or sets the batch pause interval.
        /// </summary>
        /// <value>The batch pause interval.</value>
        public int BatchPauseInterval
        {
            get
            {
                int batchPauseInterval;
                int.TryParse(this.CustomProperties[Office365Keys.BatchPauseInterval], out batchPauseInterval);
                return batchPauseInterval;
            }

            set
            {
                this.CustomProperties[Office365Keys.BatchPauseInterval] = value.ToString();
            }
        }

        /// <summary>
        /// Gets or sets the host.
        /// </summary>
        /// <value>The host.</value>
        public string Host
        {
            get
            {
                return this.CustomProperties[Office365Keys.Host];
            }

            set
            {
                this.CustomProperties[Office365Keys.Host] = value;
            }
        }

        /// <summary>
        /// Gets or sets the port.
        /// </summary>
        /// <value>The port.</value>
        public int Port
        {
            get
            {
                int port;
                int.TryParse(this.CustomProperties[Office365Keys.Port], out port);
                return port;
            }

            set
            {
                this.CustomProperties[Office365Keys.Port] = value.ToString();
            }
        }

        /// <summary>
        /// Gets or sets the username.
        /// </summary>
        /// <value>The username.</value>
        public string Username
        {
            get
            {
                return this.CustomProperties[Office365Keys.Username];
            }

            set
            {
                this.CustomProperties[Office365Keys.Username] = value;
            }
        }

        /// <summary>
        /// Gets or sets the password.
        /// </summary>
        /// <value>The password.</value>
        public string Password
        {
            get
            {
                return this.CustomProperties[Office365Keys.Password];
            }

            set
            {
                this.CustomProperties[Office365Keys.Password] = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the smtp server will use SSL.
        /// </summary>
        public bool UseSSL
        {
            get
            {
                bool useSsl;
                bool.TryParse(this.CustomProperties[Office365Keys.UseSsl], out useSsl);
                return useSsl;
            }

            set
            {
                this.CustomProperties[Office365Keys.UseSsl] = value.ToString();
            }
        }

        /// <summary>
        /// Contains constants that can be used as keys for the Office 365 SMTP sender profile specific configurations.
        /// </summary>
        public static class Office365Keys
        {
            /// <summary>
            /// Key for the ProfileType value.
            /// </summary>
            public const string ProfileType = "profileType";

            /// <summary>
            /// Key for the DefaultSenderEmailAddress value.
            /// </summary>
            public const string DefaultSenderEmailAddress = "defaultSenderEmailAddress";

            /// <summary>
            /// Key for the DefaultSenderName value.
            /// </summary>
            public const string DefaultSenderName = "defaultSenderName";

            /// <summary>
            /// Key for the TenantId value.
            /// </summary>
            public const string TenantId = "tenantId";

            /// <summary>
            /// Key for the ClientId value.
            /// </summary>
            public const string ClientId = "clientId";

            /// <summary>
            /// Key for the ClientSecret value.
            /// </summary>
            public const string ClientSecret = "clientSecret";

            /// <summary>
            /// Key for the Scopes value.
            /// </summary>
            public const string Scopes = "scopes";

            /// <summary>
            /// Key for the SenderType value.
            /// </summary>
            public const string SenderType = "senderType";

            /// <summary>
            /// Key for the BatchSize value.
            /// </summary>
            public const string BatchSize = "batchSize";

            /// <summary>
            /// Key for the BatchPauseInterval value.
            /// </summary>
            public const string BatchPauseInterval = "batchPauseInterval";

            /// <summary>
            /// Key for the Password value.
            /// </summary>
            public const string Password = "password";

            /// <summary>
            /// Key for the Username value.
            /// </summary>
            public const string Username = "username";

            /// <summary>
            /// Key for the Host value.
            /// </summary>
            public const string Host = "host";

            /// <summary>
            /// Key for the Port value.
            /// </summary>
            public const string Port = "port";

            /// <summary>
            /// Key for the UseSsl value.
            /// </summary>
            public const string UseSsl = "useSSL";
        }

        /// <summary>
        /// The type that specifies usage of smtp profile
        /// </summary>
        public const string SmtpProfileType = "smtp";

        /// <summary>
        /// The profile name for the smtp email settings 
        /// </summary>
        public const string SystemConfigSmtpSettingsMigratedProfileName = "SystemConfigSmtpSettingsMigrated";
    }
}
