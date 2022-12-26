import auth from '../../../../Auth';
import { accessToken } from '../../../../utils/accessToken';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { validation } from '../../../../utils/validation';

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
    const isAppOnlyAuth: boolean = accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken);

    if (isAppOnlyAuth && !args.options.userId && !args.options.userName) {
      throw `The option 'userId' or 'userName' is required when obtaining chats using app only permissions`;
    }
    else if (!isAppOnlyAuth && (args.options.userId || args.options.userName)) {
      throw `The options 'userId' or 'userName' cannot be used when obtaining chats using delegated permissions`;
    }

    let requestUrl = `${this.resource}/v1.0/${!isAppOnlyAuth ? 'me' : `users/${args.options.userId || args.options.userName}`}/chats`;

    if (args.options.type) {
      requestUrl += `?$filter=chatType eq '${args.options.type}'`;
    }

    try {
      const items = await odata.getAllItems(requestUrl);
      logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsChatListCommand();