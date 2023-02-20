// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ConfirmPrompt, TextPrompt, DialogSet, DialogTurnStatus, OAuthPrompt, WaterfallDialog, ComponentDialog } = require('botbuilder-dialogs');
const { TeamsActivityHandler , MessageFactory } = require('botbuilder');

const { LogoutDialog } = require('./logoutDialog');

const SKILL_DIALOG = 'SkillDialog';
const CONFIRM_PROMPT = 'ConfirmPrompt';
const TEXT_PROMPT = 'TextPrompt';
const { SimpleGraphClient } = require('../simpleGraphClient');
const { polyfills } = require('isomorphic-fetch');
const { CardFactory } = require('botbuilder-core');

class SkillDialog extends LogoutDialog {
    constructor() {
        super(SKILL_DIALOG, process.env.connectionName);

        this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
        this.addDialog(new TextPrompt(TEXT_PROMPT));
        this.addDialog(new WaterfallDialog(SKILL_DIALOG, [
            this.firstStep.bind(this),
            this.secondStep.bind(this),
            this.thirdStep.bind(this)
        ]));

        this.initialDialogId = SKILL_DIALOG;
    }

    async firstStep(stepContext) {
        // return await stepContext.context.sendActivity(MessageFactory.text('Hello, Type the skill name that you want to Search for:'));
        // return await stepContext.prompt(CONFIRM_PROMPT, 'Hello, Type the skill name that you want to Search for:');
        return await stepContext.prompt(TEXT_PROMPT, 'Hello, Type the skill name that you want to Search for:');
    }

    async secondStep(stepContext) {
        console.log("skill dialog second step:",stepContext)
        const skill_name = stepContext.result;
        await stepContext.context.sendActivity(`You have searched for '${skill_name}'`);
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
        await stepContext.context.sendActivity('Please find below:');
        // const userCard = await CardFactory.adaptiveCard(adaptiveCard[0]);
        // const userCard1 = await CardFactory.adaptiveCard(adaptiveCard[1]);
        // const userCard2 = await CardFactory.adaptiveCard(adaptiveCard[2]);
        // console.log("------------------------------userCard",userCard)
        await stepContext.context.sendActivity({ attachments: adaptiveCard, attachmentLayout: 'carousel' });
        return await stepContext.prompt(TEXT_PROMPT, 'Please select any one skill to trigger, from above:');
        // return await stepContext.endDialog();
    }

    async thirdStep(stepContext) {
        console.log("skill dialog third step:",stepContext)
        const result = stepContext.result;
        await stepContext.context.sendActivity(`You have searched for ${result}`);
        return await stepContext.endDialog();
    }
}

module.exports.SkillDialog = SkillDialog;
