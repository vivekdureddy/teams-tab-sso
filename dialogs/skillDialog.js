// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ConfirmPrompt, TextPrompt, ChoicePrompt, ChoiceFactory, DialogSet, DialogTurnStatus, OAuthPrompt, WaterfallDialog, ComponentDialog } = require('botbuilder-dialogs');
const { TeamsActivityHandler, ActivityTypes, MessageFactory } = require('botbuilder');
const axios = require('axios');
const { LogoutDialog } = require('./logoutDialog');

const SKILL_DIALOG = 'SkillDialog';
const CONFIRM_PROMPT = 'ConfirmPrompt';
const CHOICE_PROMPT = 'ChoicePrompt';
const TEXT_PROMPT = 'TextPrompt';
const OAUTH_PROMPT = 'OAuthPrompt';
const { SimpleGraphClient } = require('../simpleGraphClient');
const { polyfills } = require('isomorphic-fetch');
const { CardFactory } = require('botbuilder-core');

class SkillDialog extends LogoutDialog {
    constructor() {
        super(SKILL_DIALOG, process.env.connectionName);
        this.skill_name = '';
        this.user_email = '';
        this.orgId = 0;
        this.device_id = 0;
        this.addDialog(new OAuthPrompt(OAUTH_PROMPT, {
            connectionName: process.env.connectionName,
            text: 'Please Sign In',
            title: 'Sign In',
            timeout: 300000
        }));
        this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
        this.addDialog(new ChoicePrompt(CHOICE_PROMPT));
        this.addDialog(new TextPrompt(TEXT_PROMPT));
        this.addDialog(new WaterfallDialog(SKILL_DIALOG, [
            this.firstStep.bind(this),
            this.fetchEmail.bind(this),
            this.secondStep.bind(this),
            this.thirdStep.bind(this)
        ]));

        this.initialDialogId = SKILL_DIALOG;
    }

    async firstStep(stepContext) {
        // return await stepContext.context.sendActivity(MessageFactory.text('Hello, Type the skill name that you want to Search for:'));
        // return await stepContext.prompt(CONFIRM_PROMPT, 'Hello, Type the skill name that you want to Search for:');
        return await stepContext.prompt(TEXT_PROMPT, 'Type the Skill name that you want to Search for:');
    }

    async fetchEmail(stepContext) {
        console.log("=== console in skill dialog fetch email method ===",this.skill_name);
        let skill_name = stepContext.result;
        this.skill_name = skill_name;
        try {
            return await stepContext.beginDialog(OAUTH_PROMPT);
        } catch (err) {
            console.error(err);
            await stepContext.context.sendActivity("Error in fetching the token.");
            return await stepContext.endDialog();
        }
    }

    async secondStep(stepContext) {
        console.log("skill dialog second step:",stepContext);
        let skill_name = this.skill_name;
        try{
            let tokenResponse = stepContext.result;
            let parseToken = JSON.parse(atob(tokenResponse.token.split('.')[1]));
            console.log("=============================",parseToken);
            let user_email = parseToken.email;
            this.user_email = user_email;
            // this.orgId = parseToken.orgId;
            await stepContext.context.sendActivity(`You have searched for '${skill_name}'.`);
            const response = await fetch(`${process.env.skillHubUrl}/botapi/draftSkills/retrieveDraft?skillname=${skill_name}&email=${user_email}`);
            let skills = await response.json();
            skills = skills.data.results;
            console.log("============================================================",skills,skill_name,user_email)
            let adaptiveCard = [];
            if(skills.length){
                skills = skills.length > 9 ? skills.splice(0, 10) : skills;
                console.log("-----------------------------------------------------",skills[0])
                this.device_id = skills[0].deviceId;
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
                                            text: `${skills[i].id.toString()}_${skills[i].deviceId.toString()}`
                                        }
                                    }
                                },
                                {
                                    type: "Action.Submit",
                                    title: "Logout",
                                    data: {
                                        msteams: {
                                            type: "messageBack",
                                            displayText: "Logout",
                                            text: "logout"
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
                await stepContext.prompt(CHOICE_PROMPT, {
                    prompt: 'Please click on \'Search a skill\' to Search for a Skill / Click on \'Logout\' to sign-out.',
                    choices: ChoiceFactory.toChoices(['Search a Skill', 'Logout'])
                });
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
        } catch(err) {
            console.log("error in parse token:",err);
            return await stepContext.endDialog();
        }
    }

    async thirdStep(stepContext) {
        console.log("skill dialog third step:",typeof this.device_id);
        const result = stepContext.result;
        let { data } = await axios.post(`${process.env.skillHubUrl}/botapi/draftSkills/extension/execute`,{
            skillId: parseInt(result.split("_")[0]),
            deviceId: parseInt(result.split("_")[1]), 
            orgId: 87, 
            bulk: true
        },{
            headers: {
                "X-API-TOKEN": '8fd28b05d243fe1c753646f8',
                "X-API-ORG": 87
            }
        })
        console.log("===============================================",data)
        // await stepContext.context.sendActivity(`Please wait while we trigger the skill.`);
        await stepContext.context.sendActivities([
            { type: ActivityTypes.Message, text: 'Please wait while we trigger the skill.' },
            { type: 'delay', value: 6000 }
        ]);
        await stepContext.context.sendActivity(`Skill Triggered Successfully.`);
        await stepContext.context.sendActivities([
            { type: 'delay', value: 2000 }
        ]);
        await stepContext.prompt(CHOICE_PROMPT, {
            prompt: 'Please click on \'Search a skill\' to Search for a Skill / Click on \'Logout\' to sign-out.',
            choices: ChoiceFactory.toChoices(['Search a Skill', 'Logout'])
        });
        return await stepContext.endDialog();
    }
}

module.exports.SkillDialog = SkillDialog;
