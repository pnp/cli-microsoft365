import { Cli } from '../../../../cli/Cli';
import { CommandOutput } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import * as AadUserGetCommand from '../../../aad/commands/user/user-get';
import { Options as AadUserGetCommandOptions } from '../../../aad/commands/user/user-get';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  teamId?: string;
  userId?: string;
  userName?: string;
}

class TeamsAppInstallCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_INSTALL;
  }

  public get description(): string {
    return 'Installs a Microsoft Teams team app from the catalog in the specified team or for the specified user';
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
        teamId: typeof args.options.teamId !== 'undefined',
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--id <id>' },
      { option: '--teamId [teamId]' },
      { option: '--userId [userId]' },
      { option: '--userName [userName]' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.teamId &&
          !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.userId &&
          !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['teamId', 'userId', 'userName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await this.validateUser(args, logger);

      let url: string = `${this.resource}/v1.0`;
      if (args.options.teamId) {
        url += `/teams/${formatting.encodeQueryParameter(args.options.teamId)}/installedApps`;
      }
      else {
        url += `/users/${formatting.encodeQueryParameter((args.options.userId ?? args.options.userName) as string)}/teamwork/installedApps`;
      }

      const requestOptions: any = {
        url: url,
        headers: {
          'content-type': 'application/json;odata=nometadata',
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          'teamsApp@odata.bind': `${this.resource}/v1.0/appCatalogs/teamsApps/${args.options.id}`
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  // we need this method, because passing an invalid user ID to the API
  // won't cause an error
  private validateUser(args: CommandArgs, logger: Logger): Promise<boolean> {
    if (!args.options.userId) {
      return Promise.resolve(true);
    }

    if (this.verbose) {
      logger.logToStderr(`Checking if user ${args.options.userId} exists...`);
    }

    const options: AadUserGetCommandOptions = {
      id: args.options.userId,
      output: 'json',
      debug: args.options.debug,
      verbose: args.options.verbose
    };

    return Cli
      .executeCommandWithOutput(AadUserGetCommand as Command, { options: { ...options, _: [] } })
      .then((res: CommandOutput) => {
        if (this.verbose) {
          logger.logToStderr(res.stderr);
        }

        return true;
      }, (err: { error: CommandError, stderr: string }) => {
        if (this.verbose) {
          logger.logToStderr(err.stderr);
        }

        return Promise.reject(`User with ID ${args.options.userId} not found. Original error: ${err.error.message}`);
      });
  }
}

module.exports = new TeamsAppInstallCommand();
