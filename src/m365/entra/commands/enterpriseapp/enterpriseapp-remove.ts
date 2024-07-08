import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import aadCommands from '../../aadCommands.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  appDisplayName?: string;
  appObjectId?: string;
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
    return [aadCommands.SP_REMOVE, commands.SP_REMOVE];
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
        appId: (!(!args.options.appId)).toString(),
        appDisplayName: (!(!args.options.appDisplayName)).toString(),
        appObjectId: (!(!args.options.appObjectId)).toString(),
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --appId [appId]'
      },
      {
        option: '-n, --appDisplayName [appDisplayName]'
      },
      {
        option: '--appObjectId [appObjectId]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
          return `${args.options.appId} is not a valid appId GUID`;
        }

        if (args.options.appObjectId && !validation.isValidGuid(args.options.appObjectId)) {
          return `${args.options.appObjectId} is not a valid objectId GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['appId', 'appDisplayName', 'appObjectId'] });
  }

  #initTypes(): void {
    this.types.string.push('appId', 'appDisplayName', 'appObjectId');
  }

  private async getSpId(args: CommandArgs): Promise<string> {
    if (args.options.appObjectId) {
      return args.options.appObjectId;
    }

    let spMatchQuery: string = '';
    if (args.options.appDisplayName) {
      spMatchQuery = `displayName eq '${formatting.encodeQueryParameter(args.options.appDisplayName)}'`;
    }
    else if (args.options.appId) {
      spMatchQuery = `appId eq '${formatting.encodeQueryParameter(args.options.appId)}'`;
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
      const result = await cli.handleMultipleResultsFound<{ id: string }>(`Multiple Entra apps with name '${args.options.appDisplayName}' found.`, resultAsKeyValuePair);
      return result.id;
    }

    return spItem.id;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.SP_REMOVE, commands.SP_REMOVE);

    const removeEnterpriseApplication = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing an enterprise application ${args.options.appId || args.options.appDisplayName || args.options.appObjectId}...`);
      }

      try {
        const id = await this.getSpId(args);

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/servicePrincipals/${id}`,
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
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the enterprise application ${args.options.appId || args.options.appDisplayName || args.options.appObjectId}?` });

      if (result) {
        await removeEnterpriseApplication();
      }
    }
  }
}

export default new EntraEnterpriseAppRemoveCommand();