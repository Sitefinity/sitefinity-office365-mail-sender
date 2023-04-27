﻿using System;
using System.Collections.Generic;
using System.Configuration;
using Telerik.Sitefinity.Abstractions;
using Telerik.Sitefinity.Configuration;
using Telerik.Sitefinity.Localization;
using Telerik.Sitefinity.Services.Notifications;
using Telerik.Sitefinity.Services.Notifications.Configuration;
using Telerik.Sitefinity.Web.Configuration;

namespace Progress.Sitefinity.Office365.Mail.Sender.Configuration
{
    /// <summary>
    /// Contains the settings for the Smtp sender profile that can be used by the notification service.
    /// </summary>
    public class Office365SenderProfileElement : SenderProfileElement
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
        /// Gets or sets the default sender email address.
        /// </summary>
        /// <value>The default sender email address.</value>
        [ConfigurationProperty(Office365SenderProfileProxy.Office365Keys.DefaultSenderEmailAddress, IsRequired = true)]
        [DescriptionResource(typeof(ConfigDescriptions), "DefaultSenderEmailAddress")]
        public virtual string DefaultSenderEmailAddress
        {
            get
            {
                return (string)this[Office365SenderProfileProxy.Office365Keys.DefaultSenderEmailAddress];
            }

            set
            {
                this[Office365SenderProfileProxy.Office365Keys.DefaultSenderEmailAddress] = value;
            }
        }

        /// <summary>
        /// Gets or sets the default sender name.
        /// </summary>
        /// <value>The the default sender name.</value>
        [ConfigurationProperty(Office365SenderProfileProxy.Office365Keys.DefaultSenderName)]
        [DescriptionResource(typeof(ConfigDescriptions), "DefaultSenderName")]
        public virtual string DefaultSenderName
        {
            get
            {
                return (string)this[Office365SenderProfileProxy.Office365Keys.DefaultSenderName];
            }

            set
            {
                this[Office365SenderProfileProxy.Office365Keys.DefaultSenderName] = value;
            }
        }

        /// <summary>
        /// Gets or sets the client id.
        /// </summary>
        /// <value>The default sender email address.</value>
        [ConfigurationProperty(Office365SenderProfileProxy.Office365Keys.TenantId, IsRequired = true)]
        [DescriptionResource(typeof(ConfigDescriptions), "DefaultSenderEmailAddress")]
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
        /// Gets or sets the client id.
        /// </summary>
        /// <value>The default sender email address.</value>
        [ConfigurationProperty(Office365SenderProfileProxy.Office365Keys.ClientId, IsRequired = true)]
        [DescriptionResource(typeof(ConfigDescriptions), "DefaultSenderEmailAddress")]
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
        /// Gets or sets the client id.
        /// </summary>
        /// <value>The default sender email address.</value>
        [ConfigurationProperty(Office365SenderProfileProxy.Office365Keys.ClientSecret, IsRequired = true)]
        [DescriptionResource(typeof(ConfigDescriptions), "DefaultSenderEmailAddress")]
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
        /// Gets or sets the scopes.
        /// </summary>
        /// <value>The default sender email address.</value>
        [ConfigurationProperty(Office365SenderProfileProxy.Office365Keys.Scopes, DefaultValue = "https://graph.microsoft.com/.default", IsRequired = true)]
        [DescriptionResource(typeof(ConfigDescriptions), "DefaultSenderEmailAddress")]
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

        /// <summary>
        /// Gets or sets the size of the groups of the emails that will be sent.
        /// </summary>
        [ConfigurationProperty(Office365SenderProfileProxy.Office365Keys.BatchSize, DefaultValue = 100)]
        [DescriptionResource(typeof(NotificationsResources), "BatchSize")]
        public virtual int BatchSize
        {
            get
            {
                return (int)this[Office365SenderProfileProxy.Office365Keys.BatchSize];
            }

            set
            {
                this[Office365SenderProfileProxy.Office365Keys.BatchSize] = value;
            }
        }

        /// <summary>
        /// Gets or sets the pause in seconds between sending next group of the emails.
        /// </summary>
        [ConfigurationProperty(Office365SenderProfileProxy.Office365Keys.BatchPauseInterval, DefaultValue = 0)]
        [DescriptionResource(typeof(NotificationsResources), "BatchPauseInterval")]
        public virtual int BatchPauseInterval
        {
            get
            {
                return (int)this[Office365SenderProfileProxy.Office365Keys.BatchPauseInterval];
            }

            set
            {
                this[Office365SenderProfileProxy.Office365Keys.BatchPauseInterval] = value;
            }
        }

        /// <inheritdoc />
        public override void Initialize(IDictionary<string, string> items)
        {
            foreach (var setValue in this.SetValueDelegates)
                setValue(items);
        }

        /// <inheritdoc />
        /// <inheritdoc />
        public override Dictionary<string, string> ToDictionary()
        {
            var dict = new Dictionary<string, string>();

            dict.Add(Office365SenderProfileProxy.Office365Keys.DefaultSenderEmailAddress, this.DefaultSenderEmailAddress);
            dict.Add(Office365SenderProfileProxy.Office365Keys.DefaultSenderName, this.DefaultSenderName);
            dict.Add(Office365SenderProfileProxy.Office365Keys.ClientId, this.ClientId);
            dict.Add(Office365SenderProfileProxy.Office365Keys.ClientSecret, this.ClientSecret);
            dict.Add(Office365SenderProfileProxy.Office365Keys.TenantId, this.TenantId);
            dict.Add(Office365SenderProfileProxy.Office365Keys.SenderType, this.SenderType);
            return dict;
        }

        private bool SetDefaultSenderEmail(IDictionary<string, string> items)
        {
            bool valueChanged = false;
            string senderEmail;
            if (items.TryGetValue(Office365SenderProfileProxy.Office365Keys.DefaultSenderEmailAddress, out senderEmail) &&
                !senderEmail.IsNullOrWhitespace())
            {
                if (this.DefaultSenderEmailAddress != senderEmail)
                {
                    this.DefaultSenderEmailAddress = senderEmail;
                    valueChanged = true;
                }
            }
            else
            {
                throw new ArgumentException(string.Format("The '{0}' parameter must be specified for the office365 sender profile.", Office365SenderProfileProxy.Office365Keys.DefaultSenderEmailAddress));
            }

            return valueChanged;
        }

        private bool SetDefaultSenderName(IDictionary<string, string> items)
        {
            bool valueChanged = false;
            string defaultSenderName;
            if (items.TryGetValue(Office365SenderProfileProxy.Office365Keys.DefaultSenderName, out defaultSenderName))
            {
                if (this.DefaultSenderName != defaultSenderName)
                {
                    this.DefaultSenderName = defaultSenderName;
                    valueChanged = true;
                }
            }

            return valueChanged;
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

        private bool SetSenderType(IDictionary<string, string> items)
        {
            bool valueChanged = false;
            string senderType;
            if (items.TryGetValue(Office365SenderProfileProxy.Office365Keys.SenderType, out senderType))
            {
                if (this.SenderType != senderType)
                {
                    this.SenderType = senderType;
                    valueChanged = true;
                }
            }

            return valueChanged;
        }

        private bool SetBatchSize(IDictionary<string, string> items)
        {
            bool valueChanged = false;
            string batchSizeString;
            int batchSize;
            if (items.TryGetValue(Office365SenderProfileProxy.Office365Keys.BatchSize, out batchSizeString) &&
                int.TryParse(batchSizeString, out batchSize))
            {
                if (this.BatchSize != batchSize)
                {
                    this.BatchSize = batchSize;
                    valueChanged = true;
                }
            }

            return valueChanged;
        }

        private bool SetBatchPauseInterval(IDictionary<string, string> items)
        {
            bool valueChanged = false;
            string batchPauseIntervalString;
            int batchPauseInterval;
            if (items.TryGetValue(Office365SenderProfileProxy.Office365Keys.BatchPauseInterval, out batchPauseIntervalString) &&
                int.TryParse(batchPauseIntervalString, out batchPauseInterval))
            {
                if (this.BatchPauseInterval != batchPauseInterval)
                {
                    this.BatchPauseInterval = batchPauseInterval;
                    valueChanged = true;
                }
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
                        new SetValueDelegate(this.SetDefaultSenderEmail),
                        new SetValueDelegate(this.SetDefaultSenderName),
                        new SetValueDelegate(this.SetTenantId),
                        new SetValueDelegate(this.SetClientId),
                        new SetValueDelegate(this.SetClientSecret),
                        new SetValueDelegate(this.SetScopes),
                        new SetValueDelegate(this.SetSenderType),
                        new SetValueDelegate(this.SetBatchSize),
                        new SetValueDelegate(this.SetBatchPauseInterval)
                    };
                }

                return this.setValueDelegates;
            }
        }

        private delegate bool SetValueDelegate(IDictionary<string, string> items);

        private IEnumerable<SetValueDelegate> setValueDelegates;
    }
}