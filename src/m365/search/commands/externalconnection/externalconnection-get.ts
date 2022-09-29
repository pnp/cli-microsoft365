import { Logger } from '../../../../cli';
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
}

class SearchExternalConnectionGetCommand extends GraphCommand {
  public get name(): string {
    return commands.EXTERNALCONNECTION_GET;
  }

  public get description(): string {
    return 'Get a specific external connection for use in Microsoft Search';
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
        name: typeof args.options.name !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'name']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let url: string = `${this.resource}/v1.0/external/connections`;
    if (args.options.id) {
      url += `/${encodeURIComponent(args.options.id as string)}`;
    }
    else {
      url += `?$filter=name eq '${encodeURIComponent(args.options.name as string)}'`;
    }

    const requestOptions: any = {
      url: url,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      let res = await request.get<any>(requestOptions);

      if (args.options.name) {
        if (res.value.length === 0) {
          throw `External connection with name '${args.options.name}' not found`;
        }

        res = res.value[0];
      }

      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SearchExternalConnectionGetCommand();