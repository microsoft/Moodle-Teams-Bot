// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

let moodleClient = require("moodle-client");
let rp  = require("request-promise");
let promiseRetry = require("promise-retry");


/**
 * This class is a wrapper for the Moodle webservices.
 */
class SimpleMoodleClient {
    constructor(url, email, token) {
        if (!url || !url.trim()) {
            throw new Error('MoodleClient: Invalid url received.');
        }
        if (!email || !email.trim()) {
            throw new Error('MoodleClient: Invalid email received.');
        }
        if (!token || !token.trim()) {
            throw new Error('MoodleClient: Invalid token received.');
        }
        let options = {
            uri: `${ url }/local/o365/token.php`,
            qs: {
                'username': email,
                'service': 'o365_webservices',
            },
            headers: {
                'Authorization': 'Bearer '+ token
            },
            json: true
        };

        this.moodleWS = promiseRetry(function (retry) {
            return rp(options)
            .catch(retry);
        })
        .then(function (response) {
            return moodleClient.init({
                wwwroot: url,
                token: response.token,
                service: 'o365_webservices'
            });
        }, function (err) {
            console.log("Unable to initialize the Moodle client: " + err);
        });
    }

    async get_moodle_reply(intent, entities = null) {
        return await this.moodleWS.then(function(client){
            return promiseRetry(function (retry) {
                return client.call({
                    wsfunction: "local_o365_get_bot_message",
                    args: {
                        intent: intent,
                        entities: entities
                    }
                })
                .catch(retry);
            })
            .then(function (response) {
                return response;
            }, function (err) {
                console.log("Unable to get data from Moodle WS: " + err);
                return {exception: 1, message: 'WS not working'};
            });
        });
    }
}

exports.SimpleMoodleClient = SimpleMoodleClient;