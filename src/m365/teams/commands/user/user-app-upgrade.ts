import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { formatting } from '../../../../utils/formatting.js';
import { cli } from '../../../../cli/cli.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  userId?: string;
  userName?: string;
  name?: string;
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

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
      },
      {
        option: '--userId <userId>'
      },
      {
        option: '--userName <userName>'
      },
      {
        option: '--name <name>'
      }
    );
  }


  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (!validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const appId: string = await this.getAppId(args);
    const endpoint: string = `${this.resource}/v1.0`;
    const username = (args.options.userId) ? args.options.userId : args.options.userName;

    const requestOptions: CliRequestOptions = {
      url: `${endpoint}/users/${username}/teamwork/installedApps/${appId}/upgrade`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
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

export default new TeamsUserAppUpgradeCommand();