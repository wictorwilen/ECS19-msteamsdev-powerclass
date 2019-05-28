using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace AuthenticationBot
{
    public class TeamsSigninActionWorkaroundMiddleware : IMiddleware
    {
        public async Task OnTurnAsync(ITurnContext turnContext, NextDelegate next, CancellationToken cancellationToken = new CancellationToken())
        {
            // hook up onSend pipeline
            turnContext.OnSendActivities(async (ctx, activities, nextSend) =>
            {
                foreach (var activity in activities)
                {
#pragma warning disable SA1503 // Braces should not be omitted
                    if (activity.ChannelId != "msteams") continue;
                    if (activity.Attachments == null) continue;
                    if (!activity.Attachments.Any()) continue;
                    if (activity.Attachments[0].ContentType != "application/vnd.microsoft.card.signin") continue;
                    if (!(activity.Attachments[0].Content is SigninCard card)) continue;
                    if (!(card.Buttons is IList<CardAction> buttons)) continue;
                    if (!buttons.Any()) continue;
#pragma warning restore SA1503 // Braces should not be omitted

                    // Modify button type to openUrl as signIn is not working in teams
                    buttons.First().Type = ActionTypes.OpenUrl;
                }

                // run full pipeline
                return await nextSend().ConfigureAwait(false);
            });

            await next(cancellationToken);
        }
    }
}
