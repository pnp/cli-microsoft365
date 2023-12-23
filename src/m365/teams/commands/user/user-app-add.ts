import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
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

class TeamsUserAppAddCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_APP_ADD;
  }

  public get description(): string {
    return 'Install an app in the personal scope of the specified user';
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
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

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
    this.optionSets.push({ options: ['id', 'name'] });
    this.optionSets.push({ options: ['userId', 'userName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const appId: string = await this.getAppId(args);
    const userId: string = (args.options.userId ?? args.options.userName) as string;
    const endpoint: string = `${this.resource}/v1.0`;

    if (this.verbose) {
      await logger.logToStderr(`Removing app with ID ${appId} for user ${args.options.userId}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${endpoint}/users/${formatting.encodeQueryParameter(userId)}/teamwork/installedApps`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        'teamsApp@odata.bind': `${endpoint}/appCatalogs/teamsApps/${appId}`
      }
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/users/${args.options.userId}/teamwork/installedApps?$expand=teamsAppDefinition&$filter=teamsAppDefinition/displayName eq '${formatting.encodeQueryParameter(args.options.name as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: { id: string; }[] }>(requestOptions);

    if (response.value.length === 1) {
      return response.value[0].id;
    }

    if (response.value.length === 0) {
      throw `The specified Teams app does not exist`;
    }

    const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
    const result: { id: string } = (await cli.handleMultipleResultsFound(`Multiple Teams apps with name '${args.options.name}' found.`, resultAsKeyValuePair)) as { id: string };
    return result.id;
  }
}

export default new TeamsUserAppAddCommand();
