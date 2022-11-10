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
    this.#initOptionSets();
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

  #initOptionSets(): void {
    this.optionSets.push(
      ['userId', 'userName', 'email']
    );
  }

  private async getUserId(args: CommandArgs): Promise<string> {
    if (args.options.userId) {
      return args.options.userId;
    }

    const options: AadUserGetCommandOptions = {
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    if (args.options.email) {
      options.email = args.options.email;
    }
    else {
      options.userName = args.options.userName;
    }

    const output = await Cli.executeCommandWithOutput(AadUserGetCommand as Command, { options: { ...options, _: [] } });
    const getUserOutput = JSON.parse(output.stdout);
    return getUserOutput.id;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const userId: string = await this.getUserId(args);

      const requestOptions: AxiosRequestConfig = {
        url: `${this.resource}/v1.0/users/${userId}/onlineMeetings?$filter=JoinWebUrl eq '${encodeURIComponent(args.options.joinUrl)}'`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: Meeting[] }>(requestOptions);

      if (res.value && res.value.length > 0) {
        logger.log(res.value[0]);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsMeetingGetCommand();