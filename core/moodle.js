// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityTypes, TokenResponse, TurnContext, MessageFactory, CardFactory } = require('botbuilder');
const { SimpleGraphClient } = require('./../services/simple-graph-client');
const { SimpleMoodleClient } = require('./../services/simple-moodle-client');
const { Listcard } = require('./listcard');

// Moodle instance url
const MOODLE_URL = process.env.moodleUrl || false;
if(!MOODLE_URL) {
    console.error("Please set up environment variable 'moodleUrl'");
    process.exit(1);
}

class Moodle {
    /**
     * Displays available questions for users.
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     * @param {TokenResponse} tokenResponse A response that includes a user token.
     * @param {string} intent Question intent.
     * @param {any} entities Intent data.
     * @param {string} userLanguage User language in which answer should be returned.
     */
    static async callMoodleWebservice(turnContext, tokenResponse, intent, entities = {}, userLanguage = DEFAULT_LANGUAGE, bot = false) {
        if (!turnContext) {
            throw new Error('Moodle.callMoodleWebservice(): `turnContext` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('Moodle.callMoodleWebservice(): `tokenResponse` cannot be undefined.');
        }
        if (!intent) {
            throw new Error('Moodle.callMoodleWebservice(): `intent` cannot be undefined.');
        }

        const client = new SimpleGraphClient(tokenResponse.token);
        const me = await client.getMe();
        if(me.exception){
            console.log('Moodle.callMoodleWebservice(): error occured -> ' + data.message);
            await turnContext.sendActivity(bot.translator[userLanguage].__("Sorry, the answer to this question is not available for now"));
        }
        // Get message from Moodle.
        const moodleclient = new SimpleMoodleClient(MOODLE_URL, me.mail, tokenResponse.token);
        entities = JSON.stringify(entities);
        const data = await moodleclient.get_moodle_reply(intent, entities);
        if(data.language && bot){
            let language = data.language.split('_');
            let languagechanged = await bot.changeLanguage(turnContext, language[0], userLanguage);
        }
        if(data.exception){
            console.log('Moodle.callMoodleWebservice(): error occured -> ' + data.message);
            await turnContext.sendActivity(bot.translator[userLanguage].__("Sorry, the answer to this question is not available for now"));
        } else if((data.listItems && data.listItems.length > 0) || data.message != ''){
            await this.sendMoodleReply(turnContext, data);
        }else{
            await turnContext.sendActivity(bot.translator[userLanguage].__("Sorry, I do not understand"));
        }
    }

     /**
     * Sends message for user
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation.
     * @param  data Message data.
     * @param  {boolean} activityFeed If set true, message will be showed in Teams activity feed.
     */
    static async sendMoodleReply(turnContext, data, activityFeed = false) {
        if(data.listItems && data.listItems.length > 0){
            let cardListItems = [];
            for(let item of data.listItems){
                let action = null;
                if(item.action || item.url){
                    item.actionType = item.actionType || 'openUrl';
                    item.action = item.action || item.url;
                    action = {
                        type: item.actionType,
                        value: item.action
                    }
                }
                cardListItems.push(
                    Listcard.createListCardItem(
                    'resultItem',
                    item.title,
                    item.subtitle,
                    item.icon,
                    action
                    )
                );
            }
            let listCard = Listcard.createListCard(data.listTitle, cardListItems);
            let messageWithCard = MessageFactory.list([listCard], data.message);
            if(activityFeed){
                messageWithCard.channelData = {notification: {alert: true}};
            }
            await turnContext.sendActivity(messageWithCard);
        }else{
            let proactivenotification = { type: ActivityTypes.Message };
            proactivenotification.text = data.message;
            if(activityFeed){
                proactivenotification.channelData = {notification: {alert: true}};
            }
            await turnContext.sendActivity(proactivenotification);
        }
    }

    /**
     * Sends proactive notification for user
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation.
     * @param  data Message data.
     */
    static async sendProactiveNotification(turnContext, data, usertranslator) {
        if(data.listItems && data.listItems.length == 1){
            let item = data.listItems[0];
            item.actionType = item.actionType || 'openUrl';
            item.action = item.action || item.url;
            let actions = null;
            if(item.action){
                actions = CardFactory.actions([
                    {
                        type: item.actionType,
                        title: usertranslator.__("View"),
                        value: item.action
                    }
                ]);
            }
            let icon = null;
            if(item.icon){
                icon = [{ url: item.icon }];
            }
            const messageCard = CardFactory.thumbnailCard(item.title, icon, actions, {text: data.message});
            await turnContext.sendActivity({ attachments: [messageCard], channelData: { notification: { alert: true } } });
        } else {
            await this.sendMoodleReply(turnContext, data, true);
        }
    }
}

exports.Moodle = Moodle;