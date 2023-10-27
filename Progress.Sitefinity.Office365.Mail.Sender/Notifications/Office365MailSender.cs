using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Graph.Users.Item.SendMail;
using Progress.Sitefinity.Office365.Mail.Sender.Configuration;
using Progress.Sitefinity.Office365.Mail.Sender.Model;
using Telerik.Sitefinity.Abstractions;
using Telerik.Sitefinity.Services.Notifications;
using Telerik.Sitefinity.Services.Notifications.Composition;
using Telerik.Sitefinity.Services.Notifications.Model;

namespace Progress.Sitefinity.Office365.MailSender.Notifications
{
    /// <summary>
    /// Office365 mail sender
    /// </summary>
    public class Office365MailSender : Sender, IBatchSender
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Office365MailSender" /> class.
        /// </summary>
        /// <param name="senderProfile">The sender profile.</param>
        public Office365MailSender(Office365SenderProfileElement senderProfile)
        {
            this.profile = senderProfile;
        }

        /// <summary>
        /// Gets the size of the batch.
        /// </summary>
        /// <value>The size of the batch.</value>
        public int BatchSize
        {
            get
            {
                return this.profile.BatchSize;
            }
        }

        /// <summary>
        /// Gets the interval in seconds between batches.
        /// </summary>
        /// <value>The pause interval.</value>
        public int PauseInterval
        {
            get
            {
                return this.profile.BatchPauseInterval;
            }
        }

        /// <inheritdoc />
        public override void Dispose()
        {
        }

        /// <inheritdoc />
        public SendResult SendMessage(IMessageJobRequest messageJob, IEnumerable<ISubscriberResponse> subscribers)
        {
            SendResult result = new SendResult();
            foreach (var subscriber in subscribers)
            {
                // format the template with the subscriber object
                var senderEmailAddress =
                    !string.IsNullOrEmpty(messageJob.SenderEmailAddress) ? messageJob.SenderEmailAddress :
                    !string.IsNullOrEmpty(messageJob.MessageTemplate.TemplateSenderEmailAddress) ? messageJob.MessageTemplate.TemplateSenderEmailAddress :
                    this.profile.DefaultSenderEmailAddress;

                var senderName = 
                    !string.IsNullOrEmpty(messageJob.SenderName) ?  messageJob.SenderName :
                    !string.IsNullOrEmpty(messageJob.MessageTemplate.TemplateSenderName) ? messageJob.MessageTemplate.TemplateSenderName : 
                    this.profile.DefaultSenderName;

                MessageInfo subscriberMessage = new MessageInfo(messageJob.MessageTemplate, senderEmailAddress, senderName);

                var subscriberResult = this.SendMessage(subscriberMessage, subscriber);

                var notifiableSubscriber = subscriber as INotifiable;
                if (notifiableSubscriber != null)
                {
                    notifiableSubscriber.Result = subscriberResult.Type;
                    if (subscriberResult.Type == SendResultType.Success)
                    {
                        notifiableSubscriber.IsNotified = true;
                    }
                }

                if ((subscriberResult.Type == SendResultType.Failed || subscriberResult.Type == SendResultType.FailedRecipient) && result.Type != SendResultType.Failed)
                    result = subscriberResult;
            }

            return result;
        }

        /// <inheritdoc />
        public override SendResult SendMessage(IMessageInfo messageInfo, ISubscriberRequest subscriber)
        {
            var requestBody = new SendMailPostRequestBody
            {
                Message = new Message()
                {
                    Subject = messageInfo.Subject,
                    Body = new ItemBody
                    {
                        ContentType = messageInfo.IsHtml ? BodyType.Html : BodyType.Text,
                        Content = messageInfo.IsHtml ? messageInfo.BodyHtml : messageInfo.PlainTextVersion
                    },
                    ToRecipients = new List<Recipient>
                    {
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = subscriber.Email
                            }
                        }
                    },
                },
                SaveToSentItems = true
            };

            try
            {
                var result = Task.Run<Task>(async () => await this.GraphClient.Users[messageInfo.SenderEmailAddress].SendMail.PostAsync(requestBody)).Result;
            }
            catch (ODataError odataError)
            {
                string message = string.Format("{0}{1}", odataError.Error.Message, odataError.Error.Code);
                LogException(odataError);
                throw odataError;
            }
            catch(Exception e)
            {
                LogException(e);
                throw e;
            }

            return SendResult.ReturnSuccess();
        }

        /// <summary>
        /// Gets the configured instance of the <see cref="GraphServiceClient"/>.
        /// </summary>
        /// <returns>An instance of the <see cref="GraphServiceClient"/>.</returns>
        public GraphServiceClient GraphClient
        {
            get
            {
                if (this.graphClient == null)
                {
                    this.graphClient = this.CreateGraphClient(this.profile);
                }

                return this.graphClient;
            }
        }

        /// <summary>
        /// Create Microsoft Graph Client using OAuth
        /// </summary>
        /// <param name="profile">the profile</param>
        /// <returns>Return microsoft graph client</returns>
        private GraphServiceClient CreateGraphClient(Office365SenderProfileElement profile)
        {
            try
            {
                var scopes = profile.Scopes.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                ClientSecretCredential credential = new ClientSecretCredential(profile.TenantId, profile.ClientId, profile.ClientSecret);
                GraphServiceClient graphClient = new GraphServiceClient(credential, scopes);

                return graphClient;
            }
            catch (Exception e)
            {
                LogException(e);
                throw e;
            }
        }

        private void LogException(Exception exception)
        {
            Log.Write(exception, ConfigurationPolicy.ErrorLog);
        }

        private readonly Office365SenderProfileElement profile;
        private GraphServiceClient graphClient;
    }
}
