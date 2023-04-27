using System;
using System.ComponentModel;
using Progress.Sitefinity.Office365.Mail.Sender.Configuration;
using Progress.Sitefinity.Office365.MailSender.Notifications;
using Telerik.Sitefinity.Abstractions;
using Telerik.Sitefinity.Configuration;
using Telerik.Sitefinity.Services.Notifications.Configuration;

namespace Progress.Sitefinity.Office365.Mail.Sender
{
    /// <summary>
    /// Contains the application startup event handlers registering the required components for the translations module of Sitefinity.
    /// </summary>
    public static class Startup
    {
        /// <summary>
        /// Called before the Asp.Net application is started. Subscribes for the logging and exception handling configuration related events.
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public static void OnPreApplicationStart()
        {
            Bootstrapper.Bootstrapped += Startup.BootstrappedBootstrapping;
        }

        private static void BootstrappedBootstrapping(object sender, EventArgs e)
        {
            var configSection = Config.Get<NotificationsConfig>();
            var profiles = configSection.Profiles;

            SenderProfileElement element;
            if (!profiles.TryGetValue(Office365ProfileName, out element)) 
            {
                profiles.Add(new Office365SenderProfileElement(profiles)
                {
                    ProfileName = Office365ProfileName,
                    SenderType = typeof(Office365MailSender).FullName,
                    BatchSize = 100,
                    BatchPauseInterval = 0
                });
            }
        }

        private const string Office365ProfileName = "Office365";
    }
}
