import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import GraphCommand from "../../../base/GraphCommand";
import commands from '../../commands';
import { validation } from '../../../../utils/validation';
import { accessToken } from '../../../../utils/accessToken';
import { Event } from '@microsoft/microsoft-graph-types';
import { aadUser } from '../../../../utils/aadUser';
import { formatting } from '../../../../utils/formatting';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
  email?: string;
  joinUrl: string;
}

class TeamsMeetingGetCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_GET;
  }

  public get description(): string {
    return 'Get specified meeting details';
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
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        email: typeof args.options.email !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --userId [userId]'
      },
      {
        option: '-n, --userName [userName]'
      },
      {
        option: '--email [email]'
      },
      {
        option: '-j, --joinUrl <joinUrl>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid Guid`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken);
    if (isAppOnlyAccessToken) {
      if (!args.options.userId && !args.options.userName && !args.options.email) {
        this.handleError(`The option 'userId', 'userName' or 'email' is required when retrieving meetings using app only permissions`);
      }
    }
    else {
      if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName || args.options.email)) {
        this.handleError(`The options 'userId', 'userName' and 'email' cannot be used when retrieving meetings using delegated permissions`);
      }
    }

    if (this.verbose) {
      logger.logToStderr(`Retrieving meeting for ${isAppOnlyAccessToken ? 'specific user' : 'currently logged in user'}`);
    }

    try {
      let requestUrl = `${this.resource}/v1.0/`;

      if (isAppOnlyAccessToken) {
        requestUrl += 'users/';

        const userId = await this.getUserId(args.options);
        requestUrl += userId;
      }
      else {
        requestUrl += `me`;
      }

      requestUrl += `/onlineMeetings?$filter=JoinWebUrl eq '${formatting.encodeQueryParameter(args.options.joinUrl)}'`;

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: Event[] }>(requestOptions);

      if (res.value.length > 0) {
        logger.log(res.value[0]);
      }
      else {
        throw `The specified meeting was not found`;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserId(options: Options): Promise<string> {
    if (options.userId) {
      return options.userId;
    }

    if (options.userName) {
      return aadUser.getUserIdByUpn(options.userName);
    }

    return aadUser.getUserIdByEmail(options.email!);
  }
}

module.exports = new TeamsMeetingGetCommand();