﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telerik.Sitefinity.Services.Notifications.Configuration;

namespace Progress.Sitefinity.Office365.Mail.Sender.Configuration
{
    /// <summary>
    /// A proxy class which declares properties required to configure a smtp client
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
        /// Gets or sets the tenat id.
        /// </summary>
        /// <value>The domain.</value>
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
        /// Gets or sets the client id.
        /// </summary>
        /// <value>The domain.</value>
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
        /// Gets or sets the client secret.
        /// </summary>
        /// <value>The domain.</value>
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
        /// Gets or sets the scopes.
        /// </summary>
        /// <value>The domain.</value>
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
        /// Contains constants that can be used as keys for the SMTP sender profile specific configurations.
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
            /// Key for the ClientSecret value.
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