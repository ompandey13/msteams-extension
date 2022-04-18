import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import { default as axios } from "axios";
import {
  AdaptiveCardInvokeResponse,
  AdaptiveCardInvokeValue,
  CardFactory,
  TeamsActivityHandler,
  TurnContext,
} from "botbuilder";
import * as querystring from "querystring";
// import rawWelcomeCard from "./adaptiveCards/welcome.json"
import rawLearnCard from "./adaptiveCards/learn.json";

export interface DataInterface {
  likeCount: number;
}

interface IContact {
  id: number;

  // name
  full_name: string | null;

  // emails and phones (phone numbers are in e164 format)
  phone_number: string;
  emails: string[];
  secondary_phone_numbers: string[];

  // misc
  information?: string;
  picture_url?: string;
  company_name?: string;

  // extracts
  extracts: IContactExtract[];

  // owner
  owner_id: number;

  // date
  created_at: string;
  updated_at: string;
}
interface IContactExtract {
  id: number;

  // name
  first_name?: string;
  last_name?: string;
  full_name?: string;

  // emails and phones
  phone_number: string;
  emails: string[];
  secondary_phone_numbers: string[];

  // misc
  information?: string;
  picture_url?: string;
  company_name?: string;
  external_id: string;
  external_link?: string;

  // sources info
  source_app_id: string;
  source_id: string;
  source_app_name: string;

  // date
  created_at: string;
  updated_at: string;
}

function fetchCRM(contact: IContact) {
  const items = [];
  if (contact?.extracts[1]?.source_app_name) {
    const item = {
      title: contact.extracts[1].source_app_name,
      value: `[View in ${contact.extracts[1].source_app_name}](${contact.extracts[1].external_link})`,
    };
    items.push(item);
  }
  return items;
}

let contact: IContact;

contact = {
  id: 123,
  full_name: "Albert Flores",
  phone_number: "+34768776683",
  secondary_phone_numbers: [],
  emails: ["m.morrison@teleworm.es"],
  extracts: [
    {
      id: 1234,
      external_id: "1234",
      source_app_name: "Hubspot",
      external_link: "https://google.com",
      source_app_id: "23455",
      source_id: "1233",
      phone_number: "",
      emails: [],
      secondary_phone_numbers: [],
      created_at: "",
      updated_at: "",
    },
  ],
  owner_id: 123,
  created_at: "",
  updated_at: "",
  company_name: "Sony",
};

const rawWelcomeCard = greenJson(contact);

export function greenJson(contact: IContact) {
  const contactName = contact.full_name
    ? `${contact.full_name}`
    : "Aircall Contact";
  const greenJson = {
    $schema: "https://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.0",
    body: [
      {
        type: "Container",
        items: [
          {
            type: "ColumnSet",
            columns: [
              {
                type: "Column",
                verticalContentAlignment: "center",
                items: [
                  {
                    type: "Image",
                    url: "https://www.colorhexa.com/08b48c.png",
                    width: "400px",
                    height: "3px",
                  },
                ],
                width: "stretch",
              },
            ],
          },
        ],
      },
      {
        type: "ColumnSet",
        columns: [
          {
            type: "Column",
            width: "auto",
            items: [
              {
                type: "Image",
                url: "https://connectorsdemo.azurewebsites.net/images/MSC12_Oscar_002.jpg",
                size: "small",
                style: "person",
                altText: `${contactName}\'s Profile Picture`,
              },
            ],
          },
          {
            type: "Column",
            width: "stretch",
            items: [
              {
                type: "TextBlock",
                text: contactName,
                weight: "bolder",
                wrap: true,
              },
              {
                type: "TextBlock",
                spacing: "none",
                text: contact.company_name,
                isSubtle: true,
                wrap: true,
              },
            ],
          },
        ],
      },
      {
        type: "Container",
        items: [
          {
            type: "FactSet",
            facts: [
              {
                title: "Mobile:",
                value: contact.phone_number,
              },
              {
                title: "Work:",
                value: "+49111000460",
              },
              {
                title: "Email:",
                value: contact.emails[0],
              },
              ...fetchCRM(contact),
            ],
          },
        ],
      },
      {
        type: "Container",
        items: [
          {
            type: "ColumnSet",
            columns: [
              {
                type: "Column",
                width: "stretch",
                items: [
                  {
                    type: "TextBlock",
                    text: "This contact was last edited on 24 Jan 2020",
                    isSubtle: true,
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
    actions: [
      {
        type: "Action.OpenUrl",
        title: "Call Mobile",
        url: "tel:+32 50 58 05 67",
      },
      {
        type: "Action.OpenUrl",
        title: "Call Work",
        url: "tel:+32 50 58 05 67",
      },
      {
        type: "Action.OpenUrl",
        title: "VCard",
        url: `http://localhost:3001/vcard?email=${contact.emails[0]}`,
      },
    ],
  };
  return greenJson;
}

export class TeamsBot extends TeamsActivityHandler {
  // record the likeCount
  likeCountObj: { likeCount: number };

  constructor() {
    super();

    this.likeCountObj = { likeCount: 0 };

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");

      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      );
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Trigger command by IM text
      switch (txt) {
        case "welcome": {
          const card =
            AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          });
          break;
        }
        case "learn": {
          this.likeCountObj.likeCount = 0;
          const card = AdaptiveCards.declare<DataInterface>(
            rawLearnCard
          ).render(this.likeCountObj);
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          });
          break;
        }
        /**
         * case "yourCommand": {
         *   await context.sendActivity(`Add your response here!`);
         *   break;
         * }
         */
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card =
            AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          });
          break;
        }
      }
      await next();
    });
  }

  // Invoked when an action is taken on an Adaptive Card. The Adaptive Card sends an event to the Bot and this
  // method handles that event.
  async onAdaptiveCardInvoke(
    context: TurnContext,
    invokeValue: AdaptiveCardInvokeValue
  ): Promise<AdaptiveCardInvokeResponse> {
    // The verb "userlike" is sent from the Adaptive Card defined in adaptiveCards/learn.json
    if (invokeValue.action.verb === "userlike") {
      this.likeCountObj.likeCount++;
      const card = AdaptiveCards.declare<DataInterface>(rawLearnCard).render(
        this.likeCountObj
      );
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(card)],
      });
      return { statusCode: 200, type: undefined, value: undefined };
    } else if (invokeValue.action.verb === "vCard") {
      // this.likeCountObj.likeCount++;
      // const card = AdaptiveCards.declare<DataInterface>(rawLearnCard).render(
      //   this.likeCountObj
      // );
      // await context.updateActivity({
      //   type: "message",
      //   id: context.activity.replyToId,
      //   attachments: [CardFactory.adaptiveCard(card)],
      // });
      let vCardsJS = require("vcards-js");

      // This is your vCard instance, that
      // represents a single contact file
      const vCard = vCardsJS();

      // Set contact properties
      vCard.firstName = "James";
      vCard.middleName = "Daniel";
      vCard.lastName = "Smith";
      vCard.organization = "GeeksforGeeks";
      vCard.title = "Technical Writer";
      vCard.email = "james@example.com";
      vCard.cellPhone = "+1 (123) 456-789";
      vCard.version = "3.0";
      // Save contact to VCF file
      vCard.saveToFile(`contact.vcf`);
      const attachment = {
        name: "vCard",
        contentType: "image/jpg",
        // content: vCard.getFormattedString(),
        contentUrl: vCard.getFormattedString(),
      };

      // const temp = attachment;

      // console.log({ temp });

      await context.sendActivity({
        // type: "message",
        text: "Attachment",
        attachments: [attachment],
        // attachmentLayout: "list",
      });

      return { statusCode: 200, type: undefined, value: undefined };
    }
  }

  // Messaging extension Code
  // Action.
  public async handleTeamsMessagingExtensionSubmitAction(
    context: TurnContext,
    action: any
  ): Promise<any> {
    switch (action.commandId) {
      case "createCard":
        return createCardCommand(context, action);
      case "shareMessage":
        return shareMessageCommand(context, action);
      default:
        throw new Error("NotImplemented");
    }
  }

  // Search.
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: any
  ): Promise<any> {
    const searchQuery = query.parameters[0].value;
    const response = await axios.get(
      `http://registry.npmjs.com/-/v1/search?${querystring.stringify({
        text: searchQuery,
        size: 8,
      })}`
    );

    const attachments = [];
    response.data.objects.forEach((obj) => {
      const heroCard = CardFactory.heroCard(obj.package.name);
      const preview = CardFactory.heroCard(obj.package.name);
      preview.content.tap = {
        type: "invoke",
        value: { name: obj.package.name, description: obj.package.description },
      };
      const attachment = { ...heroCard, preview };
      attachments.push(attachment);
    });

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }

  public async handleTeamsMessagingExtensionSelectItem(
    context: TurnContext,
    obj: any
  ): Promise<any> {
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [CardFactory.heroCard(obj.name, obj.description)],
      },
    };
  }

  // Link Unfurling.
  public async handleTeamsAppBasedLinkQuery(
    context: TurnContext,
    query: any
  ): Promise<any> {
    const attachment = CardFactory.thumbnailCard(
      "Image Preview Card",
      query.url,
      [query.url]
    );

    const result = {
      attachmentLayout: "list",
      type: "result",
      attachments: [attachment],
    };

    const response = {
      composeExtension: result,
    };
    return response;
  }
}

async function createCardCommand(
  context: TurnContext,
  action: any
): Promise<any> {
  // The user has chosen to create a card by choosing the 'Create Card' context menu command.
  const data = action.data;
  const heroCard = CardFactory.heroCard(data.title, data.text);
  heroCard.content.subtitle = data.subTitle;
  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
}

async function shareMessageCommand(
  context: TurnContext,
  action: any
): Promise<any> {
  // The user has chosen to share a message by choosing the 'Share Message' context menu command.
  let userName = "unknown";
  if (
    action.messagePayload &&
    action.messagePayload.from &&
    action.messagePayload.from.user &&
    action.messagePayload.from.user.displayName
  ) {
    userName = action.messagePayload.from.user.displayName;
  }

  // This Messaging Extension example allows the user to check a box to include an image with the
  // shared message.  This demonstrates sending custom parameters along with the message payload.
  let images = [];
  const includeImage = action.data.includeImage;
  if (includeImage === "true") {
    images = [
      "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQtB3AwMUeNoq4gUBGe6Ocj8kyh3bXa9ZbV7u1fVKQoyKFHdkqU",
    ];
  }
  const heroCard = CardFactory.heroCard(
    `${userName} originally sent this message:`,
    action.messagePayload.body.content,
    images
  );

  if (
    action.messagePayload &&
    action.messagePayload.attachment &&
    action.messagePayload.attachments.length > 0
  ) {
    // This sample does not add the MessagePayload Attachments.  This is left as an
    // exercise for the user.
    heroCard.content.subtitle = `(${action.messagePayload.attachments.length} Attachments not included)`;
  }

  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
}
