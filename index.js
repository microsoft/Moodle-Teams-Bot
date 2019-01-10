// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// Import required packages
const restify = require('restify');
const  Translator = require('i18n-nodejs');

// Import required bot services.
const { BotFrameworkAdapter, ConversationState, MemoryStorage, UserState } = require('botbuilder');
const { MoodleWsBot, DEFAULT_AVAILABLE_LANGUAGES, DEFAULT_LANGUAGE } = require('./core/bot');

if(!process.env.microsoftAppID) {
    console.error("Please set up environment variable 'microsoftAppID'");
    process.exit(1);
}
if(!process.env.microsoftAppPassword) {
    console.error("Please set up environment variable 'microsoftAppPassword'");
    process.exit(1);
}

let translator = [];
for (let lang of DEFAULT_AVAILABLE_LANGUAGES){
    translator[lang] = new Translator(lang, "./../../lang/translations.json");
}

// Create bot adapter.
const adapter = new BotFrameworkAdapter({
    appId: process.env.microsoftAppID,
    appPassword: process.env.microsoftAppPassword,
    openIdMetadata: process.env.BotOpenIdMetadata
});

// Catch-all for any unhandled errors in bot.
adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError]: ${ error }`);
    // Send a message to the user
    if(error.code == "InvalidAuthenticationToken"){
        await context.sendActivity(translator[DEFAULT_LANGUAGE].__("Your session has timed out."));
    }else{
        await context.sendActivity(translator[DEFAULT_LANGUAGE].__("Oops. Something went wrong!"));
    }
    // Clear out state
    conversationState.clear(context);
};

// Define a state store for your bot.
const memoryStorage = new MemoryStorage();
let conversationState = new ConversationState(memoryStorage);
let userState = new UserState(memoryStorage);

// Create the main dialog.
const bot = new MoodleWsBot(conversationState, userState, translator);

// Create HTTP server
let server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function() {
    console.log(`\n${ server.name } listening to ${ server.url }`);
});

// Listen for incoming activities and route them to bot main dialog.
server.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async (context) => {
        bot.cacheBotData(memoryStorage, context.activity);
        // route to main dialog.
        await bot.onTurn(context);
    });
});
// Listen for proactive notifications and send them for users.
server.post('/api/webhook', async (req, res) => {
    bot.processProactiveMessage(req, res, adapter, memoryStorage);
});

