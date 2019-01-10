// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

class Listcard {
    static createListCard(title = '', items = [], buttons = []){
        let data = {
            contentType : "application/vnd.microsoft.teams.card.list",
            content : {
                title : title,
                items : items,
                buttons : buttons
            }
        }
        return data;
    }

    static createListCardItem(cardtype, title = '', subtitle = '', icon = null, action = null, other = null){
        let data = {
            type: cardtype,
            icon: icon,
            title: title,
            subtitle: subtitle,
            tap: action
        };
        for(let param in other){
            data[param] = other[param];
        }
        return data;
    }
}

exports.Listcard = Listcard;