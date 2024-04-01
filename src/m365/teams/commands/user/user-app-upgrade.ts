import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  userId?: string;
  userName?: string;
}

class TeamsUserAppUpgradeCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_APP_UPGRADE;
  }

  public get description(): string {
    return 'Upgrade an app in the personal scope of the specified user';
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
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id [id]'
      },
      {
        option: '--name [name]'
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
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid userName`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['id', 'name']
      },
      {
        options: ['userId', 'userName']
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Upgrading app ${args.options.id || args.options.name} for user ${args.options.userId || args.options.userName}`);
      }

      const installedAppId: string = await this.getInstalledAppId(args.options, logger);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(args.options.userId || args.options.userName!)}/teamwork/installedApps/${installedAppId}/upgrade`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getInstalledAppId(options: Options, logger: Logger): Promise<string> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving app ID`);
    }

    if (options.id) {
      return options.id;
    }

    const installedApps = await odata.getAllItems<{ id: string }>(`${this.resource}/v1.0/users/${formatting.encodeQueryParameter(options.userId || options.userName!)}/teamwork/installedApps?$expand=teamsAppDefinition&$filter=teamsAppDefinition/displayName eq '${formatting.encodeQueryParameter(options.name!)}'&$select=id`);

    if (installedApps.length === 1) {
      return installedApps[0].id;
    }

    if (installedApps.length === 0) {
      throw `The specified Teams app ${options.name!} does not exist or is not installed for the user`;
    }

    const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', installedApps);
    const result: { id: string } = (await cli.handleMultipleResultsFound(`Multiple installed Teams apps with name '${options.name}' found.`, resultAsKeyValuePair)) as { id: string };
    return result.id;
  }
}

export default new TeamsUserAppUpgradeCommand();