import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import aadCommands from '../../aadCommands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
  objectId?: string;
}

class EntraEnterpriseAppAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ENTERPRISEAPP_ADD;
  }

  public get description(): string {
    return 'Creates an enterprise application (or service principal) for a registered Entra app';
  }

  public alias(): string[] | undefined {
    return [aadCommands.SP_ADD, commands.SP_ADD];
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
        id: (!(!args.options.id)).toString(),
        displayName: (!(!args.options.displayName)).toString(),
        objectId: (!(!args.options.objectId)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --displayName [displayName]'
      },
      {
        option: '--objectId [objectId]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.objectId && !validation.isValidGuid(args.options.objectId)) {
          return `${args.options.objectId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'displayName', 'objectId'] });
  }

  private async getAppId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    let spMatchQuery: string = '';
    if (args.options.displayName) {
      spMatchQuery = `displayName eq '${formatting.encodeQueryParameter(args.options.displayName)}'`;
    }
    else if (args.options.objectId) {
      spMatchQuery = `id eq '${formatting.encodeQueryParameter(args.options.objectId)}'`;
    }

    const appIdRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/applications?$filter=${spMatchQuery}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: { appId: string; }[] }>(appIdRequestOptions);

    const spItem: { appId: string } | undefined = response.value[0];

    if (!spItem) {
      throw `The specified Entra app doesn't exist`;
    }

    if (response.value.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('appId', response.value);
      const result = await cli.handleMultipleResultsFound<{ appId: string }>(`Multiple Entra apps with name '${args.options.displayName}' found.`, resultAsKeyValuePair);
      return result.appId;
    }

    return spItem.appId;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.SP_ADD, commands.SP_ADD);

    try {
      const appId = await this.getAppId(args);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/servicePrincipals`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        data: {
          appId: appId
        },
        responseType: 'json'
      };

      const res = await request.post(requestOptions);

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraEnterpriseAppAddCommand();