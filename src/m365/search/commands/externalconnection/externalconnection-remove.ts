import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  confirm?: boolean;
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
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--id [id]' },
      { option: '--name [name]' },
      { option: '--confirm' }
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

    throw `Multiple external connections with name ${args.options.name} found. Please disambiguate (IDs): ${res.value.map(x => x.id).join(', ')}`;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.confirm) {
      await this.removeExternalConnection(args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the external connection '${args.options.id || args.options.name}'?`
      });

      if (result.continue) {
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

module.exports = new SearchExternalConnectionRemoveCommand();