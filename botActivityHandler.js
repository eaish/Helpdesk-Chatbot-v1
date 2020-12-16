// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    TurnContext,
    MessageFactory,
    TeamsActivityHandler,
    CardFactory,
    ActionTypes
} = require('botbuilder');

class BotActivityHandler extends TeamsActivityHandler {
    constructor() {
        super();
        /* Conversation Bot */
        /*  Teams bots are Microsoft Bot Framework bots.
            If a bot receives a message activity, the turn handler sees that incoming activity
            and sends it to the onMessage activity handler.
            Learn more: https://aka.ms/teams-bot-basics.

            NOTE:   Ensure the bot endpoint that services incoming conversational bot queries is
                    registered with Bot Framework.
                    Learn more: https://aka.ms/teams-register-bot. 
        */
        // Registers an activity event handler for the message event, emitted for every incoming message activity.
        this.onMessage(async (context, next) => {
            TurnContext.removeRecipientMention(context.activity);

            /*remove punctuation:*/
            context.activity.text = context.activity.text.replace(/[!"#$%&'()*+,-./:;<=>?@[\]^_`{|}~]/g,'');
            context.activity.text = context.activity.text.replace(/\s{2,}/g,' ');

            switch (context.activity.text.trim().toLowerCase()) {

            case 'hello': case 'hi': case 'good morning': case 'good afternoon': case 'good evening': case 'hi there': case 'greetings':
                await this.mentionActivityAsync(context);
                break;
            
            case 'how do i connect to the wifi': 
            case 'i cant connect to the wifi': 
            case 'how to connect to wifi': 
            case 'i dont know how to connect to the wifi': 
            case 'help me connect to the wifi':
                await this.wifiQuestionAsync(context);
                break;
            
            case 'where do i find my classes': 
            case 'where are my classes': 
            case 'i cant find my classes': 
            case 'how do i access my classes':
                await this.brightspaceQuestionAsync(context);
                break;

            case 'i forgot my password':
            case 'i cant remember my password':
            case 'ive forgotten my password':
            case 'how do i reset my password':
                await this.passwordResetAsync(context);
                break;
            
            case 'thank you': case 'thanks': case 'thank you so much': case 'thanks so much':
                await this.thankYouAsync(context);
                break;

            default:
                // By default for unknown activity directs user to helpdesk
                const replyActivity = MessageFactory.text('Sorry, I don\'t know how to answer that. Please contact helpdesk@uvic.ca for further assistance.');
                await context.sendActivity(replyActivity);
                break;
            }
            await next();
        });
        /* Conversation Bot */
    }

    /* says hello and mentions user*/
    async mentionActivityAsync(context) {
        const TextEncoder = require('html-entities').XmlEntities;

        const mention = {
            mentioned: context.activity.from,
            text: `<at>${ new TextEncoder().encode(context.activity.from.name) }</at>`,
            type: 'mention'
        };

        const replyActivity = MessageFactory.text(`Hi ${ mention.text }. What can I help you with today?`);
        replyActivity.entities = [mention];
        
        await context.sendActivity(replyActivity);
    }
    
    /*directs user to brightspace*/
    async brightspaceQuestionAsync(context) {

        const replyActivity = MessageFactory.text('You can access your courses at https://bright.uvic.ca.');
        
        await context.sendActivity(replyActivity);
    }

    /*opens articles on how to connect to wifi*/
    async wifiQuestionAsync(context) {

        const card = CardFactory.heroCard(
            'Select your operating system:', null,
            [   
            {
                type: 'openURL',
                title: 'Windows 10',
                value: 'https://www.uvic.ca/systems/support/internettelephone/wireless/uvic-win10.php',
            },
            {
                type: 'openURL',
                title: 'macOS',
                value: 'https://www.uvic.ca/systems/support/internettelephone/wireless/uvic-defaultosx.php',
            },
            {
                type: 'openURL',
                title: 'Android 8',
                value: 'https://www.uvic.ca/systems/support/internettelephone/wireless/uvic-android8.php',
            }
            ]);
        
        await context.sendActivity({ attachments: [card] });

    }

    async passwordResetAsync(context) {
        const replyActivity = MessageFactory.text('You can reset your Netlink password [here](https://www.uvic.ca/netlink/recover/identifyIssue).');
        
        await context.sendActivity(replyActivity);

    }

    /*responds to thank you*/
    async thankYouAsync(context) {

        const replyActivity = MessageFactory.text('I\'m always happy to help! :)');
        
        await context.sendActivity(replyActivity);

    }



}

module.exports.BotActivityHandler = BotActivityHandler;

