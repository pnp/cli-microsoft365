import type * as ACData from 'adaptivecards-templating';
import { Logger } from '../../../cli/Logger.js';
import GlobalOptions from '../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../request.js';
import AnonymousCommand from '../../base/AnonymousCommand.js';
import commands from '../commands.js';
import { optionsUtils } from '../../../utils/optionsUtils.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  actionUrl?: string;
  card?: string;
  cardData?: string;
  description?: string;
  imageUrl?: string;
  title?: string;
  url: string;
}

class AdaptiveCardSendCommand extends AnonymousCommand {
  public get name(): string {
    return commands.SEND;
  }

  public get description(): string {
    return 'Sends adaptive card to the specified URL';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        actionUrl: typeof args.options.actionUrl !== 'undefined',
        card: typeof args.options.card !== 'undefined',
        cardData: typeof args.options.cardData !== 'undefined',
        description: typeof args.options.description !== 'undefined',
        imageUrl: typeof args.options.imageUrl !== 'undefined',
        title: typeof args.options.title !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '-i, --imageUrl [imageUrl]'
      },
      {
        option: '-a, --actionUrl [actionUrl]'
      },
      {
        option: '--card [card]'
      },
      {
        option: '--cardData [cardData]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.card) {
          try {
            JSON.parse(args.options.card);
          }
          catch (e) {
            return `Error while parsing the card: ${e}`;
          }
        }

        if (args.options.cardData) {
          try {
            JSON.parse(args.options.cardData);
          }
          catch (e) {
            return `Error while parsing card data: ${e}`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['title', 'card'] }
    );
  }

  public allowUnknownOptions(): boolean {
    return true;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const unknownOptions = optionsUtils.getUnknownOptions(args.options, this.options);
    const unknownOptionNames: string[] = Object.getOwnPropertyNames(unknownOptions);
    const card: any = await this.getCard(args, unknownOptionNames, unknownOptions);

    const requestOptions: CliRequestOptions = {
      url: args.options.url,
      headers: {
        'content-type': 'application/json',
        'x-anonymous': true
      },
      data: {
        type: 'message',
        attachments: [{
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: card
        }]
      },
      responseType: 'json'
    };

    try {
      const res = await request.post<string | number | undefined>(requestOptions);

      if (res) {
        // when sending card to Teams succeeds, the body contains 1 which we
        // can safely ignore
        if (typeof res === 'string') {
          // when sending the webhook to Teams fails, the response is 200
          // but the body contains a string similar to 'Webhook message delivery
          // failed with error: Microsoft Teams endpoint returned HTTP error 400
          // with ContextId MS-CV=Qn6afVIGzEq...' which we should treat as
          // a failure
          if (res.indexOf('failed') > -1) {
            throw res;
          }

          await logger.log(res);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getCard(args: CommandArgs, unknownOptionNames: string[], unknownOptions: any): Promise<any> {
    // use custom card
    if (args.options.card) {
      let card: any = JSON.parse(args.options.card);
      const cardData: any = this.getCardData(args, unknownOptionNames, unknownOptions);

      if (cardData) {
        // lazy-load adaptive cards templating SDK
        const ACData = await import('adaptivecards-templating');
        const template: ACData.Template = new ACData.Template(card);

        // Create a data binding context, and set its $root property to the
        // data object to bind the template to
        const context: ACData.IEvaluationContext = {
          $root: cardData
        };

        // expand the template - this generates the final Adaptive Card
        card = template.expand(context);
      }

      return card;
    }

    // use predefined card
    const card: any = {
      type: "AdaptiveCard",
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      version: "1.2",
      body: []
    };

    if (args.options.title) {
      card.body.push({
        type: "TextBlock",
        size: "Medium",
        weight: "Bolder",
        text: args.options.title
      });
    }

    if (args.options.imageUrl) {
      card.body.push({
        type: "Image",
        url: args.options.imageUrl,
        size: "Stretch"
      });
    }

    if (args.options.description) {
      card.body.push({
        type: "TextBlock",
        text: args.options.description,
        wrap: true
      });
    }

    if (unknownOptionNames.length > 0) {
      card.body.push({
        type: "FactSet",
        facts: unknownOptionNames.map(o => {
          return {
            title: `${o}:`,
            value: unknownOptions[o]
          };
        })
      });
    }

    if (args.options.actionUrl) {
      card.actions = [
        {
          type: "Action.OpenUrl",
          title: "View",
          url: args.options.actionUrl
        }
      ];
    }

    return card;
  }

  private getCardData(args: CommandArgs, unknownOptionNames: string[], unknownOptions: any): any {
    if (args.options.cardData) {
      return JSON.parse(args.options.cardData);
    }

    if (unknownOptionNames.length > 0) {
      return unknownOptions;
    }

    if (!args.options.title &&
      !args.options.description &&
      !args.options.imageUrl &&
      !args.options.actionUrl) {
      return undefined;
    }

    const cardData: any = {};

    if (args.options.title) {
      cardData.title = args.options.title;
    }
    if (args.options.description) {
      cardData.description = args.options.description;
    }
    if (args.options.imageUrl) {
      cardData.imageUrl = args.options.imageUrl;
    }
    if (args.options.actionUrl) {
      cardData.actionUrl = args.options.actionUrl;
    }

    return cardData;
  }
}

export default new AdaptiveCardSendCommand();