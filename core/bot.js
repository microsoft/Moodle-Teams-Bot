// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActionTypes, ActivityTypes, CardFactory } = require('botbuilder');
const { DialogContext, DialogSet, WaterfallDialog, WaterfallStepContext } = require('botbuilder-dialogs');
const { Login, LOGIN_PROMPT } = require('./login');
const { Moodle } = require('./moodle');
const { LuisRecognizer } = require('botbuilder-ai');

const OAUTH_CONNECTION = process.env.oAuthConnection || false;
if(!OAUTH_CONNECTION) {
    console.error("Please set up environment variable 'oAuthConnection'");
    process.exit(1);
}

// Used to create the BotStatePropertyAccessor for storing the user's language preference.
const LANGUAGE_PREFERENCE = 'language_preference';

var DEFAULT_AVAILABLE_LANGUAGES = process.env.availableLanguages || 'en,es';

DEFAULT_AVAILABLE_LANGUAGES = DEFAULT_AVAILABLE_LANGUAGES.split(',').map(function(item) {
    return item.toLowerCase().trim();
});

//Default bot language
const DEFAULT_LANGUAGE = process.env.defaultLanguage || 'en';

//Setting up LUIS instance for each available language
const LUIS_INSTANCES = {};
for (let lang of DEFAULT_AVAILABLE_LANGUAGES){
    LUIS_INSTANCES[lang] = {
        APPLICATION_ID : process.env["luisApplicationId_"+lang]|| false,
        ENDPOINT : process.env["luisEndpoint_"+lang] || false,
        ENDPOINT_KEY : process.env["luisEndpointKey_"+lang] || false
    }
}
if(!LUIS_INSTANCES[DEFAULT_LANGUAGE].APPLICATION_ID) {
    console.error(`Please set up environment variable 'luisApplicationId_${ DEFAULT_LANGUAGE }'`);
    process.exit(1);
}
if(!LUIS_INSTANCES[DEFAULT_LANGUAGE].ENDPOINT) {
    console.error(`Please set up environment variable 'luisEndpoint_${ DEFAULT_LANGUAGE }'`);
    process.exit(1);
}
if(!LUIS_INSTANCES[DEFAULT_LANGUAGE].ENDPOINT_KEY) {
    console.error(`Please set up environment variable 'luisEndpointKey_${ DEFAULT_LANGUAGE }'`);
    process.exit(1);
}

class MoodleWsBot {
    /*
     * @param {ConversationState} conversationState The state that will contain the DialogState BotStatePropertyAccessor.
     */
    constructor(conversationState, userState, translator) {
        this.conversationState = conversationState;
        this.userState = userState;
        this.translator = translator;
        // Create property for languge selection
        this.languagePreferenceProperty = this.userState.createProperty(LANGUAGE_PREFERENCE);
        // Add the LUIS recognizer for each available language.
        this.luisRecognizer = {};
        for (let lang of DEFAULT_AVAILABLE_LANGUAGES){
            if(LUIS_INSTANCES[lang].APPLICATION_ID){
                this.luisRecognizer[lang] = new LuisRecognizer({
                    applicationId: LUIS_INSTANCES[lang].APPLICATION_ID,
                    endpoint: LUIS_INSTANCES[lang].ENDPOINT,
                    endpointKey: LUIS_INSTANCES[lang].ENDPOINT_KEY
                });
            }
        }
        // DialogState property accessor. Used to keep persist DialogState when using DialogSet.
        this.dialogState = conversationState.createProperty('dialogState');
        this.commandState = conversationState.createProperty('commandState');

        // Create a DialogSet that contains the OAuthPrompt.
        this.dialogs = new DialogSet(this.dialogState);

        // Add an OAuthPrompt with the connection name as specified on the Bot's settings blade in Azure.
        this.dialogs.add(Login.prompt(OAUTH_CONNECTION, this.translator[DEFAULT_LANGUAGE]));

        this._graphDialogId = 'graphDialog';

        // Logs in the user and calls proceeding dialogs, if login is successful.
        this.dialogs.add(new WaterfallDialog(this._graphDialogId, [
            this.promptStep.bind(this),
            this.processStep.bind(this)
        ]));
    };

    /**
     * This controls what happens when an activity get sent to the bot.
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     */
    async onTurn(turnContext) {
        const dc = await this.dialogs.createContext(turnContext);
        switch (turnContext.activity.type) {
            case ActivityTypes.Message:
                await this.processInput(dc);
                break;
            case ActivityTypes.Event:
            case ActivityTypes.Invoke:
                // Sanity check the Activity type and channel Id.
                if (turnContext.activity.type === ActivityTypes.Invoke && turnContext.activity.channelId !== 'msteams') {
                    throw new Error('The Invoke type is only valid on the MS Teams channel.');
                };
                await dc.continueDialog();
                if (!turnContext.responded) {
                    await dc.beginDialog(this._graphDialogId);
                };
                break;
            case ActivityTypes.ConversationUpdate:
                await this.sendWelcomeMessage(turnContext);
                break;
            default:
                await turnContext.sendActivity(`[${ turnContext.activity.type }]-type activity detected.`);
        }
        await this.conversationState.saveChanges(turnContext);
        await this.userState.saveChanges(turnContext);
    };

    /**
     * Creates a Hero Card that is sent as a welcome message to the user.
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     */
    async sendWelcomeMessage(turnContext) {
        const userLanguage = await this.languagePreferenceProperty.get(turnContext, DEFAULT_LANGUAGE);
        const activity = turnContext.activity;
        if (activity && activity.membersAdded) {
            const heroCard = CardFactory.heroCard(
                this.translator[userLanguage].__('Hello!'),
                undefined,
                CardFactory.actions([
                    {
                        type: ActionTypes.ImBack,
                        title: this.translator[userLanguage].__('Help'),
                        value: 'help'
                    }
                ]),
                {text: this.translator[userLanguage].__("I am Moodle Assistant, a bot that answers questions about your assignments and courses. <br/><br/> If you are curious about what I can do, just type 'help' or click on the button below and I will give you the list of questions I can answer!")}
            );

            for (const idx in activity.membersAdded) {
                if (activity.membersAdded[idx].id !== activity.recipient.id) {
                    await turnContext.sendActivity({ attachments: [heroCard] });
                }
            }
        }
    }

     /**
     * Creates a Thumbnail Card that is sent as a feedback message to the user.
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     */
    async sendFeedbackMessage(turnContext) {
        const userLanguage = await this.languagePreferenceProperty.get(turnContext, DEFAULT_LANGUAGE);
        const feedbackCard = CardFactory.thumbnailCard('', undefined, CardFactory.actions([
                {
                    type: 'openUrl',
                    title: this.translator[userLanguage].__('Give feedback'),
                    value: 'https://microsoftteams.uservoice.com/forums/916759-moodle'
                }
            ]), {text: this.translator[userLanguage].__('Please give us feedback by clicking on the button below.')}
        );
        await turnContext.sendActivity({ attachments: [feedbackCard] });
    }

    /**
     * Checks and changes User State language if needed
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     */
    async changeLanguage(turnContext, language, userLanguage) {
        if (isLanguageChangeRequested(language, userLanguage)) {
            await this.languagePreferenceProperty.set(turnContext, language);
            await this.userState.saveChanges(turnContext);
            return true;
        }else{
            return false;
        }
    }

    /**
     * Processes input and route to the appropriate step.
     * @param {DialogContext} dc DialogContext
     */
    async processInput(dc) {
        const userLanguage = await this.languagePreferenceProperty.get(dc.context, DEFAULT_LANGUAGE);
        if(dc.context.activity.channelData.team != undefined){
            await dc.context.sendActivity(this.translator[userLanguage].__('The answer to your query can not be displayed in team conversation. Please ask me the same question in personal chat.'));
        }else{
            switch (dc.context.activity.text.toLowerCase()) {
                case 'signoff':
                case 'logoff':
                case 'signout':
                case 'logout':
                    const botAdapter = dc.context.adapter;
                    await botAdapter.signOutUser(dc.context, OAUTH_CONNECTION);
                    await dc.context.sendActivity(this.translator[userLanguage].__('You are now signed out.'));
                    break;
                default:
                    // The waterfall dialog to handle the input.
                    await dc.continueDialog();
                    if (!dc.context.responded) {
                        await dc.beginDialog(this._graphDialogId);
                    }
            }
        }
    };

    /**
     * WaterfallDialogStep for storing commands and beginning the OAuthPrompt.
     * Saves the user's message as the command to execute if the message is not
     * a magic code.
     * @param {WaterfallStepContext} step WaterfallStepContext
     */
    async promptStep(step) {
        const activity = step.context.activity;

        if (activity.type === ActivityTypes.Message && !(/\d{6}/).test(activity.text)) {
            await this.commandState.set(step.context, activity.text);
            await this.conversationState.saveChanges(step.context);
        }
        return await step.beginDialog(LOGIN_PROMPT);
    }

    /**
     * WaterfallDialogStep to process the command sent by the user.
     * @param {WaterfallStepContext} step WaterfallStepContext
     */
    async processStep(step) {
        const tokenResponse = step.result;
        const userLanguage = await this.languagePreferenceProperty.get(step.context, DEFAULT_LANGUAGE);
        // If the user is authenticated the bot can use the token to make API calls.
        if (tokenResponse !== undefined) {
            let topIntent = false;
            let results = null;
            try{
                let luis = this.luisRecognizer[userLanguage] || this.luisRecognizer[DEFAULT_LANGUAGE]
                results = await this.luisRecognizer[userLanguage].recognize(step.context);
                topIntent = LuisRecognizer.topIntent(results);
            }catch(err){
                console.error(`\n [onTurnError]: LUIS does not work. Only basic commands available. -> ${ err }`);
            }
            if(topIntent){
                switch (topIntent){
                    case 'share-feedback':
                        await this.sendFeedbackMessage(step.context, userLanguage);
                        break;
                    default:
                        await Moodle.callMoodleWebservice(step.context, tokenResponse, topIntent, results.entities, userLanguage, this);
                }
            }else{
                let parts = await this.commandState.get(step.context);
                if (!parts) {
                    parts = step.context.activity.text;
                }
                parts = parts.split(' ');
                let command = parts[0].toLowerCase();
                command = command.trim();
                if (command === 'help') {
                    await Moodle.callMoodleWebservice(step.context, tokenResponse, 'get-help', null, userLanguage, this);
                } else if (command === 'feedback') {
                    await this.sendFeedbackMessage(step.context);
                } else {
                    await step.context.sendActivity(this.translator[userLanguage].__("Sorry, I do not understand"));
                }
            }
        } else {
            // Ask the user to try logging in later as they are not logged in.
            await step.context.sendActivity(this.translator[userLanguage].__("We couldn't log you in. Please try again later."));
        }
        return await step.endDialog();
    };

    async cacheBotData(storage, data){
        let userObjectId = data.from.aadObjectId;
        let userId = data.from.id;
        let serviceUrl = data.serviceUrl;
        let team = null;
        if(data.channelData != undefined && data.channelData.team != undefined){
            team = data.channelData.team.id;
        }
        let tenantId = null;
        if(data.channelData != undefined && data.channelData.tenant != undefined){
            tenantId = data.channelData.tenant.id;
        }
        try {
            let storeItems = await storage.read(["botCache"])
            var botCache = storeItems["botCache"];

            if (typeof (botCache) != 'undefined') {
                let store = false;
                if (typeof (storeItems["botCache"].usersList[userObjectId]) == 'undefined') {
                    storeItems["botCache"].usersList[userObjectId] = userId;
                    store = true;
                }
                if(storeItems["botCache"].serviceUrl != serviceUrl){
                    storeItems["botCache"].serviceUrl = serviceUrl;
                    store = true;
                }
                if(storeItems["botCache"].tenant == 'undefined' && tenantId != null){
                    storeItems["botCache"].tenant = tenantId;
                    store = true;
                }
                if(team != null){
                    let teamslist = storeItems["botCache"].teamsList;
                    teamslist.push(team);
                    // Leaving only unique array values
                    storeItems["botCache"].teamsList = [...new Set(storeItems["botCache"].teamsList)];
                }
                if(store){
                    try {
                        await storage.write(storeItems)
                    } catch (err) {
                        console.log(`Write failed: ${err}`);
                    }
                }
             } else {
                let botObject = data.recipient;
                let channelId = data.channelId;
                let teamsList = [team];
                let usersList = {};
                usersList[userObjectId] = userId;
                storeItems["botCache"] = { teamsList: teamsList, usersList: usersList, botObject: botObject,
                            tenant : tenantId, serviceUrl : serviceUrl, channelId : channelId, "eTag": "*" }
                try {
                    await storage.write(storeItems)
                } catch (err) {
                    console.log(`Write failed: ${err}`);
                }
            }
        } catch (err) {
            console.log(`Read rejected. ${err}`);
        };
    }

    async getBotCacheData(storage, property){
        let result = null;
        let storeItems = await storage.read(["botCache"]);
        if(storeItems["botCache"] != undefined){
            result = storeItems["botCache"][property];
        }
        return result;
    }

    async getUserFromTeams(userObjectId, teams, connectorClient, storage){
        let userid = null;
        let storeItems = await storage.read(["botCache"])
        let botCache = storeItems["botCache"];
        let store = false;
        for(let team of teams){
            if(userid != null){
                break;
            }
            let members = await connectorClient.conversations.getConversationMembersWithHttpOperationResponse(team);
            members = JSON.parse(members.bodyAsText);
            for(let member of members){
                if (typeof (botCache) != 'undefined') {
                    if (typeof (storeItems["botCache"].usersList[member.objectId]) == 'undefined') {
                        storeItems["botCache"].usersList[member.objectId] = member.id;
                        store = true;
                    }
                }
                if(userObjectId == member.objectId){
                    userid = member.id;
                    break;
                }
            }
        }
        if(store){
            try {
                await storage.write(storeItems)
            } catch (err) {
                console.log(`Write failed of userslist: ${err}`);
            }
        }
        return userid;
    }

    async processProactiveMessage(req, res, adapter, memoryStorage){
        try{
            parseRequest(req).then(async (data) => {
                let userId = null;
                const authHeader = req.headers.authorization || '';
                data.serviceUrl = await this.getBotCacheData(memoryStorage, 'serviceUrl');
                if(data.serviceUrl == null){
                    res.send('Bot cache empty');
                    res.status(404);
                    res.end();
                }else{
                    data.channelId = await this.getBotCacheData(memoryStorage, 'channelId');
                    adapter.authenticateRequest(data, authHeader).then(async() => {
                        const connectorClient = await adapter.createConnectorClient(data.serviceUrl);
                        let usersList = await this.getBotCacheData(memoryStorage, 'usersList');
                        if(usersList[data.user] == undefined){
                            let teamslist = await this.getBotCacheData(memoryStorage, 'teamsList');
                            userId = await this.getUserFromTeams(data.user, teamslist, connectorClient, memoryStorage);
                        }else{
                            userId = usersList[data.user];
                        }
                        if(userId != null){
                            let tenantId = await this.getBotCacheData(memoryStorage, 'tenant');
                            const botparam = await this.getBotCacheData(memoryStorage, 'botObject');
                            const tenant = { id: tenantId };
                            const user =  { id: userId };
                            const parameters = { bot: botparam, members: [user], channelData: {tenant: tenant}};
                            const newConversation = await connectorClient.conversations.createConversation(parameters);
                            const newReference = {
                                user: user,
                                bot: botparam,
                                conversation:
                                { conversationType: 'personal',
                                    id: newConversation.id },
                                channelId: data.channelId,
                                serviceUrl: data.serviceUrl
                            }
                            adapter.continueConversation(newReference, async (ctx) => {
                                const userLanguage = await this.languagePreferenceProperty.get(ctx, DEFAULT_LANGUAGE);
                                await Moodle.sendProactiveNotification(ctx, data, this.translator[userLanguage]);
                                res.send('Message sent');
                                res.status(200);
                                res.end();
                            });
                        }else{
                            res.send('User not found');
                            res.status(404);
                            res.end();
                        }
                    }, (reason) => {
                        res.send(reason);
                        res.status(401);
                        res.end();
                    });
                }
            });
        }catch(err){
            res.send(err);
            res.status(500);
            res.end();
        }
    }
};
// Check if language changes are requested
function isLanguageChangeRequested(newLanguage, currentLanguage) {
    if (!newLanguage) {
        return false;
    }
    const cleanedUpLanguage = newLanguage.toLowerCase().trim();

    if (DEFAULT_AVAILABLE_LANGUAGES.indexOf(cleanedUpLanguage) == -1) {
        return false;
    }
    return cleanedUpLanguage !== currentLanguage;
}

//sed to parse proactive notification content
function parseRequest(req) {
    return new Promise((resolve, reject) => {
            let requestData = '';
            req.on('data', (chunk) => {
                requestData += chunk;
            });
            req.on('end', () => {
                try {
                    req.body = JSON.parse(requestData);
                    resolve(req.body);
                }
                catch (err) {
                    reject(err);
                }
            });
    });
}

exports.MoodleWsBot = MoodleWsBot;
exports.DEFAULT_AVAILABLE_LANGUAGES = DEFAULT_AVAILABLE_LANGUAGES;