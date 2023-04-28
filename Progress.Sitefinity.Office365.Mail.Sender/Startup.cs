using System;
using System.ComponentModel;
using Progress.Sitefinity.Office365.Mail.Sender.Configuration;
using Progress.Sitefinity.Office365.MailSender.Notifications;
using Telerik.Sitefinity.Abstractions;
using Telerik.Sitefinity.Configuration;
using Telerik.Sitefinity.Data;
using Telerik.Sitefinity.Restriction;
using Telerik.Sitefinity.Security;
using Telerik.Sitefinity.Security.Claims;
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
            ////Bootstrapper.Bootstrapped += Startup.BootstrappedBootstrapping;
            Bootstrapper.Bootstrapped += Startup.BootstrappedBootstrapping2;
            ////Bootstrapper.Initialized += (new EventHandler<ExecutedEventArgs>(Bootstrapper_Initialized));
        }

        private static void Bootstrapper_Initialized(object sender, ExecutedEventArgs e)
        {

            if (e.CommandName == "Bootstrapped")
            {
                var userId = ClaimsManager.GetCurrentUserId();
                var user = UserManager.GetManager().GetUser(userId);

                using (new UnrestrictedModeRegion())
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

        private static void BootstrappedBootstrapping(object sender, EventArgs e)
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

        private static void BootstrappedBootstrapping2(object sender, EventArgs e)
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
