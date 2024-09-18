import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
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
  force?: boolean;
}

class EntraEnterpriseAppRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ENTERPRISEAPP_REMOVE;
  }

  public get description(): string {
    return 'Deletes an enterprise application (or service principal)';
  }

  public alias(): string[] | undefined {
    return [commands.SP_REMOVE];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        displayName: typeof args.options.displayName !== 'undefined',
        objectId: typeof args.options.objectId !== 'undefined',
        force: !!args.options.force
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
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `The option 'id' with value '${args.options.id}' is not a valid GUID.`;
        }

        if (args.options.objectId && !validation.isValidGuid(args.options.objectId)) {
          return `The option 'objectId' with value '${args.options.objectId}' is not a valid GUID.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'displayName', 'objectId'] });
  }

  #initTypes(): void {
    this.types.string.push('id', 'displayName', 'objectId');
    this.types.boolean.push('force');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeEnterpriseApplication = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing enterprise application ${args.options.id || args.options.displayName || args.options.objectId}...`);
      }

      try {
        let url = `${this.resource}/v1.0`;

        if (args.options.id) {
          url += `/servicePrincipals(appId='${args.options.id}')`;
        }
        else {
          const id = await this.getSpId(args.options);
          url += `/servicePrincipals/${id}`;
        }

        const requestOptions: CliRequestOptions = {
          url: url,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeEnterpriseApplication();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove enterprise application '${args.options.id || args.options.displayName || args.options.objectId}'?` });

      if (result) {
        await removeEnterpriseApplication();
      }
    }
  }

  private async getSpId(options: Options): Promise<string> {
    if (options.objectId) {
      return options.objectId;
    }

    const spItemsResponse = await odata.getAllItems<{ id: string }>(`${this.resource}/v1.0/servicePrincipals?$filter=displayName eq '${formatting.encodeQueryParameter(options.displayName!)}'&$select=id`);

    if (spItemsResponse.length === 0) {
      throw `The specified enterprise application does not exist.`;
    }

    if (spItemsResponse.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', spItemsResponse);
      const result = await cli.handleMultipleResultsFound<{ id: string }>(`Multiple enterprise applications with name '${options.displayName}' found.`, resultAsKeyValuePair);
      return result.id;
    }

    const spItem = spItemsResponse[0];

    return spItem.id;
  }
}

export default new EntraEnterpriseAppRemoveCommand();