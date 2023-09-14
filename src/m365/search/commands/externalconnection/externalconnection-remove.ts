import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  force?: boolean;
}

class SearchExternalConnectionRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.EXTERNALCONNECTION_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific External Connection from Microsoft Search';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--id [id]' },
      { option: '--name [name]' },
      { option: '-f, --force' }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'name'] });
  }

  private async getExternalConnectionId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/external/connections?$filter=name eq '${formatting.encodeQueryParameter(args.options.name as string)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res: { value: { id: string }[] } = await request.get<{ value: { id: string }[] }>(requestOptions);

    if (res.value.length === 1) {
      return res.value[0].id;
    }

    if (res.value.length === 0) {
      throw `The specified connection does not exist in Microsoft Search`;
    }

    const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', res.value);
    const result = await Cli.handleMultipleResultsFound<{ id: string }>(`Multiple external connections with name ${args.options.name} found.`, resultAsKeyValuePair);
    return result.id;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeExternalConnection(args);
    }
    else {
      const result = await Cli.promptForConfirmation({ message: `Are you sure you want to remove the external connection '${args.options.id || args.options.name}'?` });

      if (result) {
        await this.removeExternalConnection(args);
      }
    }
  }

  private async removeExternalConnection(args: CommandArgs): Promise<void> {
    try {
      const externalConnectionId: string = await this.getExternalConnectionId(args);
      const requestOptions: any = {
        url: `${this.resource}/v1.0/external/connections/${formatting.encodeQueryParameter(externalConnectionId)}`,
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
  }
}

export default new SearchExternalConnectionRemoveCommand();