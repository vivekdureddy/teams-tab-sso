// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ConfirmPrompt, ChoicePrompt, ChoiceFactory, TextPrompt, DialogSet, DialogTurnStatus, OAuthPrompt, WaterfallDialog, ComponentDialog } = require('botbuilder-dialogs');
const { LogoutDialog } = require('./logoutDialog');
const { SkillDialog } = require('./skillDialog');

const CONFIRM_PROMPT = 'ConfirmPrompt';
const CHOICE_PROMPT = 'ChoicePrompt';
const TEXT_PROMPT = 'TextPrompt';
const MAIN_DIALOG = 'MainDialog';
const MAIN_WATERFALL_DIALOG = 'MainWaterfallDialog';
const OAUTH_PROMPT = 'OAuthPrompt';
const { SimpleGraphClient } = require('../simpleGraphClient');
const { polyfills } = require('isomorphic-fetch');
const { CardFactory } = require('botbuilder-core');

class MainDialog extends SkillDialog {
    constructor() {
        super(MAIN_DIALOG, process.env.connectionName);

        this.addDialog(new OAuthPrompt(OAUTH_PROMPT, {
            connectionName: process.env.connectionName,
            text: 'Please Sign In',
            title: 'Sign In',
            timeout: 300000
        }));
        this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
        this.addDialog(new ChoicePrompt(CHOICE_PROMPT));
        this.addDialog(new TextPrompt(TEXT_PROMPT));
        this.addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
            this.promptStep.bind(this),
            this.loginStep.bind(this),
            // this.secondStep.bind(this),
            // this.thirdStep.bind(this)
            // this.ensureOAuth.bind(this),
            // this.displayToken.bind(this)
        ]));

        // this.addDialog(new WaterfallDialog(SKILL_DIALOG, [
        //     this.skillStep.bind(this)
        // ]));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;
    }

    /**
     * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {*} dialogContext
     */
    async run(context, accessor) {
        console.log("=== console in main dialog run method ===",context,"=== accessor ===",accessor,"=== id ===",this);
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);
        const skillDialog = new SkillDialog();
        dialogSet.add(skillDialog);
        const dialogContext = await dialogSet.createContext(context);
        const results = await dialogContext.continueDialog();
        console.log("results=======================================",results,context.activity)
        if(context.activity.text.toLowerCase() == "search a skill"){
            // await dialogContext.endDialog();
            await dialogContext.beginDialog(skillDialog.id);
        }else 
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
        // else if(results.status === DialogTurnStatus.empty){
        //     await context.sendActivity('Welcome to Supervity Bot. Please type \'login\' to sign-in. Type \'logout\' to sign-out.');
        // }
    }

    async promptStep(stepContext) {
        console.log("=== console in main dialog prompt step method ===")
        try {
            await stepContext.context.sendActivity("Initiating Login Process...");
            return await stepContext.beginDialog(OAUTH_PROMPT);
        } catch (err) {
            console.error(err);
            await stepContext.context.sendActivity("Error in fetching the token.");
            return await stepContext.endDialog();
        }
    }

    // async skillStep(stepContext) {
    //     console.log("=== console in main dialog skill step method ===")
    //     try {
    //         await stepContext.context.sendActivity("Hello");
    //         // await stepContext.endDialog();
    //         return await stepContext.beginDialog(MAIN_WATERFALL_DIALOG);
    //     } catch (err) {
    //         console.error(err);
    //     }
    // }

    async loginStep(stepContext) {
        console.log("=== console in main dialog login step method ===")
        // Get the token from the previous step. Note that we could also have gotten the
        // token directly from the prompt itself. There is an example of this in the next method.
        let tokenResponse = stepContext.result;
        if (!tokenResponse || !tokenResponse.token) {
            await stepContext.context.sendActivity('Login was not successful please try again.');
        } else {
            // const client = new SimpleGraphClient(tokenResponse.token);
            // const me = await client.getMe();
            // const title = me ? me.jobTitle : 'UnKnown';
            // await stepContext.context.sendActivity(`You're logged in as ${me.displayName} (${me.userPrincipalName}); your job title is: ${title}; your photo is: `);
            // const photoBase64 = await client.GetPhotoAsync(tokenResponse.token);
            // const card = CardFactory.thumbnailCard("", CardFactory.images([photoBase64]));
            // await stepContext.context.sendActivity({attachments: [card]});

            // return await stepContext.prompt(CONFIRM_PROMPT, 'Would you like to view your token?');
            let user_email;
            try {
                let parseToken = JSON.parse(atob(tokenResponse.token.split('.')[1]));
                user_email = parseToken.email;
                console.log("parsed token:",parseToken);
                await stepContext.context.sendActivity(`You have been successfully logged in as '${user_email}'.`);
                await stepContext.prompt(CHOICE_PROMPT, {
                    prompt: 'Please click on \'Search a skill\' to Search for a Skill / Click on \'Logout\' to sign-out.',
                    choices: ChoiceFactory.toChoices(['Search a Skill', 'Logout'])
                });
            } catch(err) {
                console.log("error in parse token:",err);
                return await stepContext.endDialog();
            }
            // return await stepContext.prompt(TEXT_PROMPT, 'Type the \'Skill\' that you want to Search for:');
        }
        return await stepContext.endDialog();
    }

    async secondStep(stepContext) {
        console.log("skill dialog second step:",stepContext)
        const skill_name = stepContext.result;
        await stepContext.context.sendActivity(`You have searched for '${skill_name}'.`);
        const response = await fetch(process.env.skillUrl+"?skillname="+skill_name);
        let skills = await response.json();
        skills = skills.data.results
        console.log("============================================================",skills)
        let adaptiveCard = [];
        if(skills.length){
            skills = skills.length > 9 ? skills.splice(0, 10) : skills;
            for(let i=0; i<skills.length; ++i){
                let card = {
                    contentType: "application/vnd.microsoft.card.adaptive",
                    content: {
                        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                        type: "AdaptiveCard",
                        version: "1.0",
                        body: [
                            {
                                type: "Image",
                                style: "Person",
                                url: skills[i].IMAGE,
                                size: "Large"
                            },
                            {
                                type: "TextBlock",
                                size: "Medium",
                                weight: "Bolder",
                                text: skills[i].SKILL_NAME,
                                wrap: true
                            },
                            {
                                type: "TextBlock",
                                text: skills[i].SKILL_DESCRIPTION,
                                wrap: true
                            }
                        ],
                        actions: [
                            {
                                type: "Action.Submit",
                                title: "Use Skill",
                                data: {
                                    msteams: {
                                        type: "messageBack",
                                        displayText: skills[i].SKILL_NAME,
                                        text: skills[i].SKILL_NAME
                                    }
                                }
                            }
                        ]
                    }
                }
                adaptiveCard.push(card);
            }
        }else{
            await stepContext.context.sendActivity(`There are no skills found with '${skill_name}' in our Skill Hub.`);
            return await stepContext.endDialog();
        }
        console.log("------------------------------adaptiveCard",adaptiveCard)
        await stepContext.context.sendActivity('Please find the below search results:');
        // const userCard = await CardFactory.adaptiveCard(adaptiveCard[0]);
        // const userCard1 = await CardFactory.adaptiveCard(adaptiveCard[1]);
        // const userCard2 = await CardFactory.adaptiveCard(adaptiveCard[2]);
        // console.log("------------------------------userCard",userCard)
        await stepContext.context.sendActivity({ attachments: adaptiveCard, attachmentLayout: 'carousel' });
        return await stepContext.prompt(TEXT_PROMPT, 'Click on \'Use Skill\' to trigger any skill from the above list.');
        // return await stepContext.endDialog();
    }

    async thirdStep(stepContext) {
        console.log("skill dialog third step:",stepContext)
        const result = stepContext.result;
        await stepContext.context.sendActivity(`You have selected '${result}'.`);
        return await stepContext.endDialog();
    }

    async ensureOAuth(stepContext) {
        console.log("=== console in main dialog ensureOAuth step method ===")
        await stepContext.context.sendActivity('Thank you.');
        const result = stepContext.result;
        if (result) {
            // Call the prompt again because we need the token. The reasons for this are:
            // 1. If the user is already logged in we do not need to store the token locally in the bot and worry
            // about refreshing it. We can always just call the prompt again to get the token.
            // 2. We never know how long it will take a user to respond. By the time the
            // user responds the token may have expired. The user would then be prompted to login again.
            //
            // There is no reason to store the token locally in the bot because we can always just call
            // the OAuth prompt to get the token or get a new token if needed.
            return await stepContext.beginDialog(OAUTH_PROMPT);
        }
        return await stepContext.endDialog();
    }

    async displayToken(stepContext) {
        console.log("=== console in main dialog display token method ===")
        const tokenResponse = stepContext.result;
        if (tokenResponse && tokenResponse.token) {
            await stepContext.context.sendActivity(`Here is your token ${tokenResponse.token}.`);
        }
        return await stepContext.endDialog();
    }
}

module.exports.MainDialog = MainDialog;
