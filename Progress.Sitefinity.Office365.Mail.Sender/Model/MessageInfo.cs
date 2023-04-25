using System;
using System.Collections.Generic;
using System.Globalization;
using Telerik.Sitefinity.Services.Notifications;
using Telerik.Sitefinity.Services.Notifications.Composition;

namespace Progress.Sitefinity.Office365.Mail.Sender.Model
{
    /// <summary>
    /// Message Info class
    /// </summary>
    public class MessageInfo : IMessageInfo
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MessageInfo" /> class.
        /// </summary>
        /// <param name="messageTemplate">The message template.</param>
        /// <param name="senderEmail">The sender email.</param>
        /// <param name="senderName">Name of the sender.</param>
        public MessageInfo(IMessageTemplateRequest messageTemplate, string senderEmail, string senderName)
        {
            this.Subject = messageTemplate.Subject;
            this.IsHtml = true;
            this.SenderEmailAddress = senderEmail;
            this.SenderName = senderName;
            this.BodyHtml = messageTemplate.BodyHtml;
            this.SystemMessage = messageTemplate.SystemMessage;
            this.SystemUrl = messageTemplate.SystemUrl;
            this.PlainTextVersion = messageTemplate.PlainTextVersion;
            this.TemplateSenderEmailAddress = messageTemplate.TemplateSenderEmailAddress;
            this.TemplateSenderName = messageTemplate.TemplateSenderName ?? messageTemplate.TemplateSenderEmailAddress;
            this.CustomMessageHeaders = new Dictionary<string, string>();
        }

        /// <inheritdoc />
        public string GetTitle(CultureInfo culture = null)
        {
            return this.Subject;
        }

        /// <inheritdoc />
        public string Subject
        {
            get;
            set;
        }

        /// <inheritdoc />
        public string BodyHtml
        {
            get;
            set;
        }

        /// <inheritdoc />
        public string SystemMessage
        {
            get;
            set;
        }

        /// <inheritdoc />
        public string SystemUrl
        {
            get;
            set;
        }

        /// <inheritdoc />
        public string PlainTextVersion
        {
            get;
            set;
        }

        /// <inheritdoc />
        public bool IsHtml
        {
            get;
            set;
        }

        /// <inheritdoc />
        public string SenderEmailAddress
        {
            get;
            set;
        }

        /// <inheritdoc />
        public string SenderName
        {
            get;
            set;
        }

        /// <inheritdoc />
        public IDictionary<string, string> CustomMessageHeaders
        {
            get;
            set;
        }

        /// <inheritdoc />
        public string ResolveKey
        {
            get;
            set;
        }

        /// <inheritdoc />
        public DateTime? LastModified
        {
            get;
            set;
        }

        /// <inheritdoc />
        public string ModuleName
        {
            get;
            set;
        }

        /// <inheritdoc />
        public string TemplateSenderEmailAddress
        {
            get;
            set;
        }

        /// <inheritdoc />
        public string TemplateSenderName
        {
            get;
            set;
        }

        /// <inheritdoc />
        public Guid? LastModifiedById { get; set; }

        /// <inheritdoc />
        public string LastModifiedByProvider { get; set; }

        /// <inheritdoc />
        public string AdditionalMessageData { get; set; }
    }
}
