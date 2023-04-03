// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityTypes, CardFactory } = require('botbuilder');
const { ComponentDialog } = require('botbuilder-dialogs');

class LogoutDialog extends ComponentDialog {
    constructor(id, connectionName) {
        super(id);
        this.connectionName = connectionName;
    }

    async onBeginDialog(innerDc, options) {
        console.log("=== console in logout dialog on begin method ===")
        const result = await this.interrupt(innerDc);
        if (result) {
            return result;
        }

        return await super.onBeginDialog(innerDc, options);
    }

    async onContinueDialog(innerDc) {
        console.log("=== console in logout dialog on continue method ===")
        const result = await this.interrupt(innerDc);
        if (result) {
            return result;
        }

        return await super.onContinueDialog(innerDc);
    }

    async interrupt(innerDc) {
        console.log("=== console in logout dialog, interrupt method ===")
        if (innerDc.context.activity.type === ActivityTypes.Message) {
            const text = innerDc.context.activity.text.toLowerCase();
            if (text.toLowerCase() === 'sign out' || text.toLowerCase() === 'sign out ') {
                const userTokenClient = innerDc.context.turnState.get(innerDc.context.adapter.UserTokenClientKey);

                const { activity } = innerDc.context;
                await userTokenClient.signOutUser(activity.from.id, this.connectionName, activity.channelId);

                await innerDc.context.sendActivity('You have been signed out successfully.');
                return await innerDc.cancelAllDialogs();
            }else if (text.toLowerCase() === 'help' || text.toLowerCase() === 'help '){
                let adaptiveCard = {
                    "type": "AdaptiveCard",
                    "body": [
                        {
                            "type": "TextBlock",
                            "spacing": "Medium",
                            "size": "Default",
                            "weight": "Bolder",
                            "text": "Help:",
                            "wrap": true,
                            "maxLines": 0
                        },
                        {
                            "type": "TextBlock",
                            "text": "Supervity is citizen automation platform to create, share and use digital skills.",
                            "wrap": true,
                            "size": "Default",
                            "fontType": "Default",
                            "maxLines": 0,
                            "spacing": "Small",
                            "separator": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "Supervity comes with a Skill Hub that acts as a global repository of skills created by Subject Matter Experts (SMEs) & Star Creators across popular software applications.The Supervity app for Microsoft Teams allows users to do the following tasks:",
                            "wrap": true,
                            "spacing": "Small"
                        },
                        {
                            "type": "TextBlock",
                            "text": "- Search for the skills in Supervity Skill Hub \r- Select and trigger the digital skill \r- Takes care of routine work heavily dominated by repetitive tasks \r- Enables faster adoption of software applications by automating tasks in minutes",
                            "wrap": true,
                            "spacing": "Small"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Some helpful commands:",
                            "wrap": true,
                            "spacing": "Medium"
                        },
                        {
                            "type": "TextBlock",
                            "text": "- Type **Sign in** to connect your Supervity and Microsoft Teams accounts \r- Type **Sign out** to disconnect your Supervity and Microsoft Teams accounts \r- Type **Help** to see this message again",
                            "wrap": true,
                            "spacing": "Small"
                        },
                        {
                            "type": "TextBlock",
                            "text": "New to Supervity? Learn more at [Techforce.ai](https://www.techforce.ai)",
                            "wrap": true,
                            "spacing": "Medium"
                        }
                    ],
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.5"
                }
                const userCard = await CardFactory.adaptiveCard(adaptiveCard);
                await innerDc.context.sendActivity({ attachments: [userCard], attachmentLayout: 'carousel' });
                // await innerDc.context.sendActivity('This is help dialog.');
                return await innerDc.cancelAllDialogs();
            }
        }
    }
}

module.exports.LogoutDialog = LogoutDialog;