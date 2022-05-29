// <copyright file="AdaptiveCardCreator.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard
{
    using System;
    using AdaptiveCards;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;

    /// <summary>
    /// Adaptive Card Creator service.
    /// </summary>
    public class AdaptiveCardCreator
    {
        /// <summary>
        /// Creates an adaptive card.
        /// </summary>
        /// <param name="notificationDataEntity">Notification data entity.</param>
        /// <returns>An adaptive card.</returns>
        public virtual AdaptiveCard CreateAdaptiveCard(NotificationDataEntity notificationDataEntity)
        {
            return this.CreateAdaptiveCard(
                notificationDataEntity.Title,
                notificationDataEntity.ImageLink,
                notificationDataEntity.Summary,
                notificationDataEntity.Author,
                notificationDataEntity.ButtonTitle,
                notificationDataEntity.ButtonLink);
        }

        /// <summary>
        /// Create an adaptive card instance.
        /// </summary>
        /// <param name="title">The adaptive card's title value.</param>
        /// <param name="imageUrl">The adaptive card's image URL.</param>
        /// <param name="summary">The adaptive card's summary value.</param>
        /// <param name="author">The adaptive card's author value.</param>
        /// <param name="buttonTitle">The adaptive card's button title value.</param>
        /// <param name="buttonUrl">The adaptive card's button url value.</param>
        /// <returns>The created adaptive card instance.</returns>
        public AdaptiveCard CreateAdaptiveCard(
            string title,
            string imageUrl,
            string summary,
            string author,
            string buttonTitle,
            string buttonUrl)
        {
            var version = new AdaptiveSchemaVersion(1, 0);
            AdaptiveCard card = new AdaptiveCard(version);

            string imgHeader = "https://m4zonqsiv74hs.blob.core.windows.net/companycommunicatormedia/ccheader.png?sp=r&st=2022-05-27T11:45:31Z&se=2025-08-01T19:45:31Z&spr=https&sv=2020-08-04&sr=b&sig=rHUD3XdvZZNHovD%2FA2Xpu6pFMRzX6YvYdfB73PcDzTE%3D" ;
            string imgFooter = "https://m4zonqsiv74hs.blob.core.windows.net/companycommunicatormedia/ccfooter.png?sp=r&st=2022-05-27T11:43:57Z&se=2025-08-01T19:43:57Z&spr=https&sv=2020-08-04&sr=b&sig=ih90y5vEIeIKiUE%2BAi9mPMStRCfHCROVddnVRgxK5O8%3D" ;
            

            card.Body.Add(new AdaptiveImage()
            {
                Url = new Uri(imgHeader, UriKind.RelativeOrAbsolute),
                Spacing = AdaptiveSpacing.Default,
                Size = AdaptiveImageSize.Stretch,
                AltText = string.Empty,
            });

            card.Body.Add(new AdaptiveTextBlock()
            {
                Text = title,
                Size = AdaptiveTextSize.ExtraLarge,
                HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                Weight = AdaptiveTextWeight.Bolder,
                Wrap = true,
            });

            if (!string.IsNullOrWhiteSpace(imageUrl))
            {
                card.Body.Add(new AdaptiveImage()
                {
                    Url = new Uri(imageUrl, UriKind.RelativeOrAbsolute),
                    Spacing = AdaptiveSpacing.Default,
                    Size = AdaptiveImageSize.Stretch,
                    AltText = string.Empty,
                });
            }

            if (!string.IsNullOrWhiteSpace(summary))
            {
                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = summary,
                    Wrap = true,
                });
            }

            if (!string.IsNullOrWhiteSpace(author))
            {
                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = author,
                    Size = AdaptiveTextSize.Small,
                    Weight = AdaptiveTextWeight.Lighter,
                    Wrap = true,
                });
            }

            if (!string.IsNullOrWhiteSpace(buttonTitle)
                && !string.IsNullOrWhiteSpace(buttonUrl))
            {
                card.Actions.Add(new AdaptiveOpenUrlAction()
                {
                    Title = buttonTitle,
                    Url = new Uri(buttonUrl, UriKind.RelativeOrAbsolute),
                });
            }

            card.Body.Add(new AdaptiveImage()
            {
                Url = new Uri(imgFooter, UriKind.RelativeOrAbsolute),
                Spacing = AdaptiveSpacing.Default,
                Size = AdaptiveImageSize.Stretch,
                AltText = string.Empty,
            });

            return card;
        }
    }
}
