using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Graph.Users.Item.SendMail;
using Progress.Sitefinity.Office365.Mail.Sender.Model;
using Telerik.Sitefinity.Services.Notifications;
using Telerik.Sitefinity.Services.Notifications.Composition;
using Telerik.Sitefinity.Services.Notifications.Configuration;
using Telerik.Sitefinity.Services.Notifications.Model;

namespace Progress.Sitefinity.Office365.MailSender
{
    /// <summary>
    /// Office365 mail sender
    /// </summary>
    public class Office365MailSender : Sender, IBatchSender
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MicrosoftGraphSender" /> class.
        /// </summary>
        /// <param name="senderProfile">The sender profile.</param>
        public Office365MailSender(SmtpSenderProfileElement senderProfile)
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

        public override void Dispose()
        {
        }

        public SendResult SendMessage(IMessageJobRequest messageJob, IEnumerable<ISubscriberResponse> subscribers)
        {
            SendResult result = new SendResult();
            foreach (var subscriber in subscribers)
            {
                // format the template with the subscriber object
                MessageInfo subscriberMessage = new MessageInfo(messageJob.MessageTemplate, messageJob.SenderEmailAddress, messageJob.SenderName);

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

        public override SendResult SendMessage(IMessageInfo messageInfo, ISubscriberRequest subscriber)
        {
            var requestBody = new SendMailPostRequestBody
            {
                Message = new Message()
                {
                    Subject = "Meet for lunch TEST EMAIL?",
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = "The new cafeteria is open TEST Content."
                    },
                    ToRecipients = new List<Recipient>
                    {
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = ""
                            }
                        }
                    },
                },
                SaveToSentItems = true
            };

            try
            {
                var result = Task.Run<Task>(async () => await this.GraphClient.Users[""].SendMail.PostAsync(requestBody)).Result;
            }
            catch (ODataError odataError)
            {
                string message = string.Format("{0}{1}", odataError.Error.Message, odataError.Error.Code);
                throw odataError;
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
        private GraphServiceClient CreateGraphClient(SmtpSenderProfileElement profile)
        {
            try
            {
                var scopes = new string[]
                {
                   "https://graph.microsoft.com/.default"
                };
                ClientSecretCredential credential = new ClientSecretCredential("", "", "");
                GraphServiceClient graphClient = new GraphServiceClient(credential, scopes);

                return graphClient;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private readonly SmtpSenderProfileElement profile;
        private GraphServiceClient graphClient;
    }
}
