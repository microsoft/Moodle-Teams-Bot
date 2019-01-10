// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { OAuthPrompt } = require('botbuilder-dialogs');

// DialogId for the OAuthPrompt.
const LOGIN_PROMPT = 'loginPrompt';

class Login {
    /**
     * Prompts the user to log in using the OAuth provider specified by the connection name.
     * @param {string} connectionName The connectionName from Azure when the OAuth provider is created.
     */
    static prompt(connectionName, translator) {
        const loginPrompt = new OAuthPrompt(LOGIN_PROMPT,
            {
                connectionName: connectionName,
                text: translator.__('Please login'),
                title: 'Login',
                timeout: 30000 // User has 5 minutes to login.
            });
        return loginPrompt;
    }
}

exports.Login = Login;
exports.LOGIN_PROMPT = LOGIN_PROMPT;