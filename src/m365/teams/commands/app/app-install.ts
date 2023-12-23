import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import aadUserGetCommand, { Options as AadUserGetCommandOptions } from '../../../aad/commands/user/user-get.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
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
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        teamId: typeof args.options.teamId !== 'undefined',
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --id [id]' },
      { option: '-n, --name [name]' },
      { option: '--teamId [teamId]' },
      { option: '--userId [userId]' },
      { option: '--userName [userName]' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
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
    this.optionSets.push({ options: ['id', 'name'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await this.validateUser(args, logger);
      const appId: string = await this.getAppId(args.options);

      let url: string = `${this.resource}/v1.0`;
      if (args.options.teamId) {
        url += `/teams/${formatting.encodeQueryParameter(args.options.teamId)}/installedApps`;
      }
      else {
        url += `/users/${formatting.encodeQueryParameter((args.options.userId ?? args.options.userName) as string)}/teamwork/installedApps`;
      }

      const requestOptions: CliRequestOptions = {
        url: url,
        headers: {
          'content-type': 'application/json;odata=nometadata',
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          'teamsApp@odata.bind': `${this.resource}/v1.0/appCatalogs/teamsApps/${appId}`
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
  private async validateUser(args: CommandArgs, logger: Logger): Promise<boolean> {
    if (!args.options.userId) {
      return true;
    }

    if (this.verbose) {
      await logger.logToStderr(`Checking if user ${args.options.userId} exists...`);
    }

    const options: AadUserGetCommandOptions = {
      id: args.options.userId,
      output: 'json',
      debug: args.options.debug,
      verbose: args.options.verbose
    };
    try {
      const res = await cli.executeCommandWithOutput(aadUserGetCommand as Command, { options: { ...options, _: [] } });

      if (this.verbose) {
        await logger.logToStderr(res.stderr);
      }

      return true;
    }
    catch (err: any) {
      if (this.verbose) {
        await logger.logToStderr(err.stderr);
      }

      throw `User with ID ${args.options.userId} not found. Original error: ${err.error.message}`;
    }
  }

  private async getAppId(options: Options): Promise<string> {
    if (options.id) {
      return options.id;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/appCatalogs/teamsApps?$filter=displayName eq '${formatting.encodeQueryParameter(options.name as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: { id: string; }[] }>(requestOptions);
    const app: { id: string; } | undefined = response.value[0];

    if (!app) {
      throw `The specified Teams app does not exist`;
    }

    if (response.value.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
      const result = await cli.handleMultipleResultsFound<{ id: string; }>(`Multiple Teams apps with name ${options.name} found.`, resultAsKeyValuePair);
      return result.id;
    }

    return app.id;
  }
}

export default new TeamsAppInstallCommand();
