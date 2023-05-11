using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Net.Mail;
using Telerik.Sitefinity.Abstractions;
using Telerik.Sitefinity.Configuration;
using Telerik.Sitefinity.Localization;
using Telerik.Sitefinity.Services.Notifications;
using Telerik.Sitefinity.Services.Notifications.Configuration;
using Telerik.Sitefinity.Web.Configuration;

namespace Progress.Sitefinity.Office365.Mail.Sender.Configuration
{
    /// <summary>
    /// Contains the settings for the Office 365 SMTP sender profile that can be used by the notification service.
    /// </summary>
    public class Office365SenderProfileElement : SmtpSenderProfileElement
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Office365SenderProfileElement"/> class.
        /// </summary>
        /// <param name="parent">The parent element.</param>
        /// <remarks>
        /// ConfigElementCollection generally needs to have a parent, however, sometimes it is necessary
        /// to create a collection in memory only which is then used later on in the context of a parent.
        /// Therefore, is the element is of ConfigElementCollection, exception for a non existing parent
        /// will not be thrown.
        /// </remarks>
        public Office365SenderProfileElement(ConfigElement parent)
            : base(parent)
        {
        }

        /// <summary>
        /// Gets or sets the Azure Active Directory tenant ID.
        /// </summary>
        /// <value>The tenant ID.</value>
        [ConfigurationProperty(Office365SenderProfileProxy.Office365Keys.TenantId, IsRequired = true)]
        [DescriptionResource(typeof(Office365ConfigDescription), "TenantID")]
        public virtual string TenantId
        {
            get
            {
                return (string)this[Office365SenderProfileProxy.Office365Keys.TenantId];
            }

            set
            {
                this[Office365SenderProfileProxy.Office365Keys.TenantId] = value;
            }
        }

        /// <summary>
        /// Gets or sets the application (client) ID.
        /// </summary>
        /// <value>The application (client) ID.</value>
        [ConfigurationProperty(Office365SenderProfileProxy.Office365Keys.ClientId, IsRequired = true)]
        [DescriptionResource(typeof(Office365ConfigDescription), "ClientID")]
        public virtual string ClientId
        {
            get
            {
                return (string)this[Office365SenderProfileProxy.Office365Keys.ClientId];
            }

            set
            {
                this[Office365SenderProfileProxy.Office365Keys.ClientId] = value;
            }
        }

        /// <summary>
        /// Gets or sets the client secret value.
        /// </summary>
        /// <value>The client secret value.</value>
        [ConfigurationProperty(Office365SenderProfileProxy.Office365Keys.ClientSecret, IsRequired = true)]
        [DescriptionResource(typeof(Office365ConfigDescription), "ClientSecret")]
        [SecretData]
        public virtual string ClientSecret
        {
            get
            {
                return (string)this[Office365SenderProfileProxy.Office365Keys.ClientSecret];
            }

            set
            {
                this[Office365SenderProfileProxy.Office365Keys.ClientSecret] = value;
            }
        }

        /// <summary>
        /// Gets or sets the Microsoft Graph scopes.
        /// </summary>
        /// <value>The Microsoft Graph scopes.</value>
        [ConfigurationProperty(Office365SenderProfileProxy.Office365Keys.Scopes, DefaultValue = "https://graph.microsoft.com/.default", IsRequired = true)]
        [DescriptionResource(typeof(Office365ConfigDescription), "Scopes")]
        public virtual string Scopes
        {
            get
            {
                return (string)this[Office365SenderProfileProxy.Office365Keys.Scopes];
            }

            set
            {
                this[Office365SenderProfileProxy.Office365Keys.Scopes] = value;
            }
        }

        [Browsable(false)]
        public override string Host => base.Host;

        [Browsable(false)]
        public override int Port => base.Port;

        [Browsable(false)]
        public override bool UseAuthentication => base.UseAuthentication;

        [Browsable(false)]
        public override string Username => base.Username;

        [Browsable(false)]
        public override string Password => base.Password;

        [Browsable(false)]
        public override string Domain => base.Domain;

        [Browsable(false)]
        public override bool UseSSL => base.UseSSL;

        [Browsable(false)]
        public override SmtpDeliveryMethod DeliveryMethod => base.DeliveryMethod;

        [Browsable(false)]
        public override string PickupDirectoryLocation => base.PickupDirectoryLocation;

        [Browsable(false)]
        public override int Timeout => base.Timeout;

        [Browsable(false)]
        public override string EmailSubjectEncoding => base.EmailSubjectEncoding;

        [Browsable(false)]
        public override string EmailBodyEncoding => base.EmailBodyEncoding;

        /// <inheritdoc />
        public override void Initialize(IDictionary<string, string> items)
        {
            base.Initialize(items);

            foreach (var setValue in this.SetValueDelegates)
                setValue(items);
        }

        /// <inheritdoc />
        public override Dictionary<string, string> ToDictionary()
        {
            var dict = base.ToDictionary();

            dict.Add(Office365SenderProfileProxy.Office365Keys.ClientId, this.ClientId);
            dict.Add(Office365SenderProfileProxy.Office365Keys.ClientSecret, this.ClientSecret);
            dict.Add(Office365SenderProfileProxy.Office365Keys.TenantId, this.TenantId);
            dict.Add(Office365SenderProfileProxy.Office365Keys.Scopes, this.Scopes);

            return dict;
        }


        private bool SetTenantId(IDictionary<string, string> items)
        {
            bool valueChanged = false;
            string tenantId;
            if (items.TryGetValue(Office365SenderProfileProxy.Office365Keys.TenantId, out tenantId))
                if (this.TenantId != tenantId)
                {
                    this.TenantId = tenantId;
                    valueChanged = true;
                }

            return valueChanged;
        }

        private bool SetClientId(IDictionary<string, string> items)
        {
            bool valueChanged = false;
            string clientId;
            if (items.TryGetValue(Office365SenderProfileProxy.Office365Keys.ClientId, out clientId))
                if (this.ClientId != clientId)
                {
                    this.ClientId = clientId;
                    valueChanged = true;
                }

            return valueChanged;
        }

        private bool SetClientSecret(IDictionary<string, string> items)
        {
            bool valueChanged = false;
            string clientSecret;
            if (items.TryGetValue(Office365SenderProfileProxy.Office365Keys.ClientSecret, out clientSecret))
                if (this.ClientSecret != clientSecret)
                {
                    this.ClientSecret = clientSecret;
                    valueChanged = true;
                }

            return valueChanged;
        }

        private bool SetScopes(IDictionary<string, string> items)
        {
            bool valueChanged = false;
            string scopes;
            if (items.TryGetValue(Office365SenderProfileProxy.Office365Keys.Scopes, out scopes))
            {
                scopes = scopes.Replace(" ", string.Empty);
                if (this.Scopes != scopes)
                {
                    this.Scopes = scopes;
                    valueChanged = true;
                }
            }
            else
            {
                throw new ArgumentException(string.Format("The '{0}' parameter must be specified for the office365 sender profile.", Office365SenderProfileProxy.Office365Keys.DefaultSenderEmailAddress));
            }

            return valueChanged;
        }

        private IEnumerable<SetValueDelegate> SetValueDelegates
        {
            get
            {
                if (this.setValueDelegates == null)
                {
                    this.setValueDelegates = new List<SetValueDelegate>
                    {
                        new SetValueDelegate(this.SetTenantId),
                        new SetValueDelegate(this.SetClientId),
                        new SetValueDelegate(this.SetClientSecret),
                        new SetValueDelegate(this.SetScopes),
                    };
                }

                return this.setValueDelegates;
            }
        }

        private delegate bool SetValueDelegate(IDictionary<string, string> items);

        private IEnumerable<SetValueDelegate> setValueDelegates;
    }
}
