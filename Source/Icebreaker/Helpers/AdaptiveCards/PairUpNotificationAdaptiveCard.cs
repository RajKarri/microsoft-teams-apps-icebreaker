// <copyright file="PairUpNotificationAdaptiveCard.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Icebreaker.Helpers.AdaptiveCards
{
    using System;
    using System.Globalization;
    using global::AdaptiveCards;
    using global::AdaptiveCards.Templating;
    using Icebreaker.Properties;
    using Microsoft.Azure;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;

    /// <summary>
    /// Builder class for the pairup notification card
    /// </summary>
    public class PairUpNotificationAdaptiveCard : AdaptiveCardBase
    {
        /// <summary>
        /// Default marker string in the UPN that indicates a user is externally-authenticated
        /// </summary>
        private const string ExternallyAuthenticatedUpnMarker = "#ext#";

        private static readonly Lazy<AdaptiveCardTemplate> AdaptiveCardTemplate =
            new Lazy<AdaptiveCardTemplate>(() => CardTemplateHelper.GetAdaptiveCardTemplate(AdaptiveCardName.PairUpNotification));

        /// <summary>
        /// Creates the adaptive card for the team welcome message
        /// </summary>
        /// <param name="teamName">The team name</param>
        /// <param name="botInstaller">The name of the person that installed the bot</param>
        /// <returns>The welcome team adaptive card</returns>
        // public static Attachment GetCard(string teamName, string botInstaller)
        public static Attachment GetCard(string teamName, TeamsChannelAccount sender, TeamsChannelAccount recipient, string botInstaller)
        {
            // Set alignment of text based on default locale.
            var textAlignment = CultureInfo.CurrentCulture.TextInfo.IsRightToLeft ? AdaptiveHorizontalAlignment.Right.ToString() : AdaptiveHorizontalAlignment.Left.ToString();

            // Guest users may not have their given name specified in AAD, so fall back to the full name if needed
            var senderGivenName = string.IsNullOrEmpty(sender.GivenName) ? sender.Name : sender.GivenName;
            var recipientGivenName = string.IsNullOrEmpty(recipient.GivenName) ? recipient.Name : recipient.GivenName;

            // To start a chat with a guest user, use their external email, not the UPN
            var recipientUpn = !IsGuestUser(recipient) ? recipient.UserPrincipalName : recipient.Email;

            string teamIntroPart1;
            string teamIntroPart2;
            string teamIntroPart3;

            if (string.IsNullOrEmpty(botInstaller))
            {
                teamIntroPart1 = string.Format(Resources.InstallMessageUnknownInstallerPart1, teamName);
                teamIntroPart2 = Resources.InstallMessageUnknownInstallerPart2;
                teamIntroPart3 = Resources.InstallMessageUnknownInstallerPart3;
            }
            else
            {
                teamIntroPart1 = string.Format(Resources.InstallMessageKnownInstallerPart1, botInstaller, teamName);
                teamIntroPart2 = Resources.InstallMessageKnownInstallerPart2;
                teamIntroPart3 = Resources.InstallMessageKnownInstallerPart3;
            }

            var baseDomain = CloudConfigurationManager.GetSetting("AppBaseDomain");
            var tourTitle = Resources.WelcomeTourTitle;
            var appId = CloudConfigurationManager.GetSetting("ManifestAppId");

            var welcomeData = new
            {
                textAlignment,
                teamIntroPart1,
                teamIntroPart2,
                teamIntroPart3,
                welcomeCardImageUrl = $"https://{baseDomain}/Content/welcome-card-image.png",
                tourUrl = GetTourFullUrl(appId, GetTourUrl(baseDomain), tourTitle),
                salutationText = Resources.SalutationTitleText,
                tourButtonText = Resources.TakeATourButtonText,
                personUpn = recipientUpn,
            };

            return GetCard(AdaptiveCardTemplate.Value, welcomeData);
        }

        /// <summary>
        /// Checks whether or not the account is a guest user.
        /// </summary>
        /// <param name="account">The <see cref="TeamsChannelAccount"/> user to check.</param>
        /// <returns>True if the account is a guest user, false otherwise.</returns>
        private static bool IsGuestUser(TeamsChannelAccount account)
        {
            return account.UserPrincipalName.IndexOf(ExternallyAuthenticatedUpnMarker, StringComparison.InvariantCultureIgnoreCase) >= 0;
        }
    }
}
