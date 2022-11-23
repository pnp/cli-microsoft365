import auth, { Auth } from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import Command from '../../../../Command';
import { Logger } from '../../../../cli/Logger';
import { AxiosRequestConfig } from 'axios';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from "../../../base/GraphCommand";
import commands from '../../commands';
import { Meeting } from '../Meeting';
import { validation } from '../../../../utils/validation';
import * as AadUserGetCommand from '../../../aad/commands/user/user-get';
import { Options as AadUserGetCommandOptions } from '../../../aad/commands/user/user-get';

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

  private async getUserId(email: string): Promise<string> {
    const options: AadUserGetCommandOptions = {
      email: email,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await Cli.executeCommandWithOutput(AadUserGetCommand as Command, { options: { ...options, _: [] } });
    const getUserOutput = JSON.parse(output.stdout);
    return getUserOutput.id;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAuth: boolean | undefined = Auth.isAppOnlyAuth(auth.service.accessTokens[this.resource].accessToken);
    if (this.verbose) {
      logger.logToStderr(`Retrieving meeting for ${isAppOnlyAuth ? 'specific user' : 'currently logged in user'}`);
    }

    let requestUrl = `${this.resource}/v1.0/`;
    if (isAppOnlyAuth) {
      if (!args.options.userId && !args.options.userName && !args.options.email) {
        this.handleError(`The option 'userId', 'userName' or 'email' is required when retrieving meetings using app only permissions`);
      }

      requestUrl += 'users/';
      if (args.options.userId) {
        requestUrl += args.options.userId;
      }
      else if (args.options.userName) {
        requestUrl += args.options.userName;
      }
      else if (args.options.email) {
        const userId = await this.getUserId(args.options.email);
        requestUrl += userId;
      }
    }
    else {
      if (args.options.userId || args.options.userName || args.options.email) {
        this.handleError(`The options 'userId', 'userName' and 'email' cannot be used when retrieving meetings using delegated permissions`);
      }

      requestUrl += `me`;
    }

    requestUrl += `/onlineMeetings?$filter=JoinWebUrl eq '${encodeURIComponent(args.options.joinUrl)}'`;

    try {
      const requestOptions: AxiosRequestConfig = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: Meeting[] }>(requestOptions);

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
}

module.exports = new TeamsMeetingGetCommand();