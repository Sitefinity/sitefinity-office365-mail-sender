using System;
using System.ComponentModel;
using Progress.Sitefinity.Office365.Mail.Sender.Configuration;
using Progress.Sitefinity.Office365.MailSender.Notifications;
using Telerik.Sitefinity.Abstractions;
using Telerik.Sitefinity.Configuration;
using Telerik.Sitefinity.Data;
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
            Bootstrapper.Initialized += (new EventHandler<ExecutedEventArgs>(Bootstrapper_Initialized));
        }

        private static void Bootstrapper_Initialized(object sender, ExecutedEventArgs e)
        {
            if (e.CommandName == "Bootstrapped")
            {
                using (new ElevatedConfigModeRegion())
                {
                    var configManager = ConfigManager.GetManager();
                    var notificationConfig = configManager.GetSection<NotificationsConfig>();

                    SenderProfileElement element;
                    if (!notificationConfig.Profiles.TryGetValue(Office365ProfileName, out element))
                    {
                        notificationConfig.Profiles.Add(new Office365SenderProfileElement(notificationConfig.Profiles)
                        {
                            ProfileName = Office365ProfileName,
                            SenderType = typeof(Office365MailSender).FullName,
                            BatchSize = 100,
                            BatchPauseInterval = 0
                        });

                        configManager.SaveSection(notificationConfig);
                    }
                }
            }
        }

        private const string Office365ProfileName = "Office365";
    }
}
