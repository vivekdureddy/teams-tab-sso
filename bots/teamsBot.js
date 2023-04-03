// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { DialogBot } = require('./dialogBot');
const { ActivityTypes, CardFactory } = require('botbuilder');

class TeamsBot extends DialogBot {
    /**
     *
     * @param {ConversationState} conversationState
     * @param {UserState} userState
     * @param {Dialog} dialog
     */
    constructor(conversationState, userState, dialog) {
        super(conversationState, userState, dialog);

        this.onMembersAdded(async (context, next) => {
            console.log(" === console in Teams Bot on members added method ===",context)
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    let adaptiveCard = {
                        "type": "AdaptiveCard",
                        "body": [
                            {
                                "type": "TextBlock",
                                "spacing": "Medium",
                                "size": "Default",
                                "weight": "Bolder",
                                "text": "Welcome to Supervity Bot!",
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
                    await context.sendActivity({ attachments: [userCard], attachmentLayout: 'carousel' });
                    // await context.sendActivity('Welcome to Supervity Bot. Please type \'login\' to sign-in. Type \'logout\' to sign-out.');
                }
            }

            await next();
        });
    }

    // async onTokenResponseEvent(context) {
    // console.log('Running dialog with Token Response Event Activity.');

    // // Run the Dialog with the new Token Response Event Activity.
    // await this.dialog.run(context, this.dialogState);
    // }

    async handleTeamsSigninVerifyState(context, query) {
        console.log('Running dialog with signin/verifystate from an Invoke Activity.');
        await this.dialog.run(context, this.dialogState);
    }
    async handleTeamsSigninTokenExchange(context, query) { console.log('Running dialog with signin/tokenExchange from an Invoke Activity.'); await this.dialog.run(context, this.dialogState); }
}


module.exports.TeamsBot = TeamsBot;
