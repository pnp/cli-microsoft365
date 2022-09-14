import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
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
    this.optionSets.push(['id', 'name']);
  }

  private getExternalConnectionId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return Promise.resolve(args.options.id);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/external/connections?$filter=name eq '${encodeURIComponent(args.options.name as string)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: { id: string }[] }>(requestOptions)
      .then((res: { value: { id: string }[] }): Promise<string> => {
        if (res.value.length === 1) {
          return Promise.resolve(res.value[0].id);
        }

        if (res.value.length === 0) {
          return Promise.reject(`The specified connection does not exist in Microsoft Search`);
        }

        return Promise.reject(`Multiple external connections with name ${args.options.name} found. Please disambiguate (IDs): ${res.value.map(x => x.id).join(', ')}`);
      });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeExternalConnection: () => void = (): void => {
      this
        .getExternalConnectionId(args)
        .then((externalConnectionId: string): Promise<void> => {
          const requestOptions: any = {
            url: `${this.resource}/v1.0/external/connections/${encodeURIComponent(externalConnectionId)}`,
            headers: {
              accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
          };

          return request
            .delete(requestOptions);
        })
        .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
    };

    if (args.options.confirm) {
      removeExternalConnection();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the external connection '${args.options.id || args.options.name}'?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeExternalConnection();
        }
      });
    }
  }
}

module.exports = new SearchExternalConnectionRemoveCommand();