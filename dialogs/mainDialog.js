// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ConfirmPrompt, DialogSet, DialogTurnStatus, OAuthPrompt, WaterfallDialog, ComponentDialog } = require('botbuilder-dialogs');
const { LogoutDialog } = require('./logoutDialog');
const { SkillDialog } = require('./skillDialog');

const CONFIRM_PROMPT = 'ConfirmPrompt';
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
        this.addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
            this.promptStep.bind(this),
            this.loginStep.bind(this),
            this.ensureOAuth.bind(this),
            this.displayToken.bind(this)
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
        console.log("dialogSet================================",dialogSet)
        dialogSet.add(this);
        const skillDialog = new SkillDialog();
        dialogSet.add(skillDialog);
        const dialogContext = await dialogSet.createContext(context);
        console.log("dialogContext===============================",dialogContext)
        const results = await dialogContext.continueDialog();
        console.log("results=======================================",results,context.activity)
        if(context.activity.text == "skill"){
            // await dialogContext.endDialog();
            await dialogContext.beginDialog(skillDialog.id);
        }else 
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    async promptStep(stepContext) {
        console.log("=== console in main dialog prompt step method ===")
        try {
            await stepContext.context.sendActivity("Initiating Login Process:")
            return await stepContext.beginDialog(OAUTH_PROMPT);
        } catch (err) {
            console.error(err);
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
        const tokenResponse = stepContext.result;
        console.log("token========================",tokenResponse)
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
            return await stepContext.prompt(CONFIRM_PROMPT, 'Would you like to view your token?');
        }
        return await stepContext.endDialog();
    }

    async ensureOAuth(stepContext) {
        console.log("=== console in main dialog ensureOAuth step method ===")
        await stepContext.context.sendActivity('Thank you.');

        const result = stepContext.result;
        console.log("---------------------------------")
        if (result) {
            console.log("-----------------------------------------------------------------------------------------------------------------------------------")
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
            await stepContext.context.sendActivity(`Here is your token ${tokenResponse.token}`);
        }
        return await stepContext.endDialog();
    }
}

module.exports.MainDialog = MainDialog;
