// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { Client } = require('@microsoft/microsoft-graph-client');
let promiseRetry = require("promise-retry");

/**
 * This class is a wrapper for the Microsoft Graph API.
 * See: https://developer.microsoft.com/en-us/graph for more information.
 */
class SimpleGraphClient {
    constructor(token) {
        if (!token || !token.trim()) {
            throw new Error('SimpleGraphClient: Invalid token received.');
        }

        this._token = token;

        // Get an Authenticated Microsoft Graph client using the token issued to the user.
        this.graphClient = Client.init({
            authProvider: (done) => {
                done(null, this._token); // First parameter takes an error if you can't get an access token.
            }
        });
    }

    /**
     * Collects information about the user in the bot.
     */
    async getMe() {
        let client = await this.graphClient;
        return promiseRetry(function (retry) {
            return client.api('/me')
                .get()
                .catch(retry);
            })
            .then(function (response) {
                return response;
            }, function (err) {
                console.log("Unable to get data from Graph API: " + err);
                return {exception: 1, message: 'Graph API not working'};
            });
    }
}

exports.SimpleGraphClient = SimpleGraphClient;
