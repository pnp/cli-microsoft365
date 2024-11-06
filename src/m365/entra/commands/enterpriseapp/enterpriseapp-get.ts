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
  displayName?: string;
  objectId?: string;
}

class EntraEnterpriseAppGetCommand extends GraphCommand {
  public get name(): string {
    return commands.ENTERPRISEAPP_GET;
  }

  public get description(): string {
    return 'Gets information about an Enterprise Application';
  }

  public alias(): string[] | undefined {
    return [commands.SP_GET];
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

  private async getSpId(args: CommandArgs): Promise<string> {
    if (args.options.objectId) {
      return args.options.objectId;
    }

    let spMatchQuery: string = '';
    if (args.options.displayName) {
      spMatchQuery = `displayName eq '${formatting.encodeQueryParameter(args.options.displayName)}'`;
    }
    else if (args.options.id) {
      spMatchQuery = `appId eq '${formatting.encodeQueryParameter(args.options.id)}'`;
    }

    const idRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/servicePrincipals?$filter=${spMatchQuery}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: { id: string; }[] }>(idRequestOptions);

    const spItem: { id: string } | undefined = response.value[0];

    if (!spItem) {
      throw `The specified Entra app does not exist`;
    }

    if (response.value.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
      const result = await cli.handleMultipleResultsFound<{ id: string }>(`Multiple Entra apps with name '${args.options.displayName}' found.`, resultAsKeyValuePair);
      return result.id;
    }

    return spItem.id;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving enterprise application information...`);
    }

    try {
      const id = await this.getSpId(args);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/servicePrincipals/${id}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraEnterpriseAppGetCommand();