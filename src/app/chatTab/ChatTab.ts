import { PreventIframe } from "express-msteams-host";
import * as Util from "util";
const TextEncoder = Util.TextEncoder;
import { BotDeclaration } from 'express-msteams-host';
import {
    StatePropertyAccessor,
    CardFactory,
    TurnContext,
    MemoryStorage,
    ConversationState,
    ActivityTypes,
    ActivityHandler,
    TeamsActivityHandler,
    ConversationReference,
    MessageFactory,
    ActionTypes,
    AttachmentLayoutTypes
} from 'botbuilder';

import getPersonas from "../database";


/**
 * Used as place holder for the decorators
 */
@PreventIframe("/chatTab/index.html")
@PreventIframe("/chatTab/config.html")
@PreventIframe("/chatTab/remove.html")
@BotDeclaration(
    '/api/messages',
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)
export class ChatTab extends ActivityHandler {

    conversationReferences: object;

    constructor(conversationReferences) {
        super();

        this.conversationReferences = conversationReferences;
        this.onConversationUpdate(async (context: TurnContext, next: () => Promise<void>) => {
            console.log(context)
            //let persona = await getPersonas(context.activity.from.name);
            if (!conversationReferences.greeted) {
                conversationReferences.greeted = true;

                await context.sendActivity('Hi');
                await context.sendActivity('Seems like you are new around here..');
                await context.sendActivity('I am your first buddy here..Lets get started by knowing about each other');
                await context.sendActivity('Would you like to tell me your name?');
            }
            console.log(TurnContext.getConversationReference(context.activity));
            this.addConversationReference(TurnContext.getConversationReference(context.activity));
            await next();
        });

        this.onMessage(async (context: TurnContext, next: () => Promise<void>) => {
            console.log(context);
            const botMessageText: String = context.activity.text.trim().toLowerCase();
            console.log(botMessageText);

            if (botMessageText.endsWith("</at> mentionme")) {
                await this.handleMessageMentionMeChannelConversation(context);
            }

            switch (botMessageText) {
                case "mentionme":
                    await this.handleMessageMentionMeOneOnOne(context);
                    await next();
                    break;
                case "filltimesheet":
                    await this.handleTimesheetSubmission(context, conversationReferences);
                    conversationReferences.lastQuestionAsked = 'fillTimeSheets';
                    await next();
                    break;
                case "applyleave":
                    await this.handleMessageMentionMeOneOnOne(context);
                    await next();
                    break;
                case "help":
                    await this.sendHelpCard(context);
                    await next();
                    break;
                case "connect with mentor":
                    await this.handleConnectionWithMentor(context);
                    await next();
                    break;
                case "logout":
                    await context.sendActivity('Bye..See you !!');
                    conversationReferences.greetedWithName = false;
                    await next();
                    break;
                default:
                    if (botMessageText.includes("worked for") && conversationReferences.lastQuestionAsked === 'fillTimeSheets') {
                        await context.sendActivity('Updating the system with the provided details...');
                        await context.sendActivity('Update complete');
                    }
                    else {

                        var message = botMessageText.split(" ");
                        var name = message[message.length - 1];
                        let username = context.activity.from.name.toLowerCase();
                        if (username.includes(name) && (conversationReferences.greetedWithName === undefined || !conversationReferences.greetedWithName)) {
                            conversationReferences.greetedWithName = true;
                            await this.handleMessageMentionMeOneOnOne(context);
                        }
                        else {
                            await context.sendActivity(`I am sorry..I didn't quite get that`);
                        }
                        await next();
                    }
                    break;
            }

        })
    }



    //**************** Helper Functions *******************/

    addConversationReference(conversationalReference) {
        const conversationReference = conversationalReference;
        this.conversationReferences[conversationalReference.conversation.id] = conversationalReference;
    }

    private async handleMessageMentionMeOneOnOne(context: TurnContext): Promise<void> {
        const mention = {
            mentioned: context.activity.from,
            text: `<at>${new TextEncoder().encode(context.activity.from.name)}</at>`,
            type: "mention"
        };

        let persona = await getPersonas(context.activity.from.name);

        const replyActivity = MessageFactory.text(`Hi ${mention.text}`);
        replyActivity.entities = [mention];
        await context.sendActivity(replyActivity);
        await context.sendActivity(persona.greetingmessage);
        let AdaptiveCard = require('../cards/Image.json');

        AdaptiveCard.body[0].columns[0].items[1].url = persona.greetingimage;
        await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(AdaptiveCard)],
            attachmentLayout: AttachmentLayoutTypes.Carousel
        });
        await context.sendActivity('Like I said.. I am your first buddy here. Remember.. If you need my assistance then choose one of the existing options from the menu or type a direct message and send through.');
    }

    private async handleMessageMentionMeChannelConversation(context: TurnContext): Promise<void> {
        const mention = {
            mentioned: context.activity.from,
            text: `<at>${new TextEncoder().encode(context.activity.from.name)}</at>`,
            type: "mention"
        };

        const replyActivity = MessageFactory.text(`Hi ${mention.text}!`);
        replyActivity.entities = [mention];
        const followupActivity = MessageFactory.text(`*We are in a channel conversation*`);
        await context.sendActivities([replyActivity, followupActivity]);
    }

    async sendHelpCard(context) {
        console.log("in");
        let AdaptiveCard = require('../cards/Image.json');

        AdaptiveCard.body[0].columns[0].items[1].url = 'https://vcarebuddy.blob.core.windows.net/$web/help.gif';
        await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(AdaptiveCard)],
            attachmentLayout: AttachmentLayoutTypes.Carousel
        });
        const card = CardFactory.heroCard(
            'I am here to help..',
            'Let me know what do you need assistance with?',
            [''],
            [
                {
                    type: ActionTypes.PostBack,
                    title: 'Connect With Mentor',
                    value: 'connectmentor'
                },
                {
                    type: ActionTypes.PostBack,
                    title: 'Project Related Help',
                    value: 'projecthelp'
                },
                {
                    type: ActionTypes.PostBack,
                    title: 'Others',
                    value: 'otherhelp'
                }
            ]
        );

        await context.sendActivity({ attachments: [card] });
    }

    async handleConnectionWithMentor(context) {
        await context.sendActivity('Swipe through the mentors to connect with them..');
        let mentorCard1 = require('../cards/mentor1.json');
        let mentorCard2 = require('../cards/mentor2.json');
        let mentorCard3 = require('../cards/mentor3.json');


        await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(mentorCard1), CardFactory.adaptiveCard(mentorCard2), CardFactory.adaptiveCard(mentorCard3)],
            attachmentLayout: AttachmentLayoutTypes.Carousel
        });
    }

    async handleTimesheetSubmission(context, conversationReferences) {
        await context.sendActivity('Tell me the details in this format "Worked for #TimeInHours on #Project On #DateWorked"')

    }

}
