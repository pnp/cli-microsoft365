import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type?: string;
  userId?: string;
  userName?: string;
}

class TeamsChatListCommand extends GraphCommand {
  public supportedTypes = ['oneOnOne', 'group', 'meeting'];
  public get name(): string {
    return commands.CHAT_LIST;
  }

  public get description(): string {
    return 'Lists all chat conversations';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'topic', 'chatType'];
  }

  constructor() {
    super();


    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        type: args.options.type
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --type [type]',
        autocomplete: this.supportedTypes
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.type !== undefined && this.supportedTypes.indexOf(args.options.type) === -1) {
          return `${args.options.type} is not a valid chatType. Accepted values are ${this.supportedTypes.join(', ')}`;
        }

        if (args.options.userId && args.options.userName) {
          return `You can only specify either 'userId' or 'userName'`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);

    if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
      throw `The option 'userId' or 'userName' is required when obtaining chats using app only permissions`;
    }
    else if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName)) {
      throw `The options 'userId' or 'userName' cannot be used when obtaining chats using delegated permissions`;
    }

    let requestUrl = `${this.resource}/v1.0/${!isAppOnlyAccessToken ? 'me' : `users/${args.options.userId || args.options.userName}`}/chats`;

    if (args.options.type) {
      requestUrl += `?$filter=chatType eq '${args.options.type}'`;
    }

    try {
      const items = await odata.getAllItems(requestUrl);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsChatListCommand();