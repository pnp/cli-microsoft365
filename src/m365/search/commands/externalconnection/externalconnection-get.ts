import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
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

    request
      .get(requestOptions)
      .then((res: any): Promise<void> => {
        if (args.options.name) {      
          if (res.value.length === 0) {
            return Promise.reject(`External connection with name '${args.options.name}' not found`);
          }
          
          res = res.value[0];
        }

        return Promise.resolve(res);
      })
      .then(res => {
        logger.log(res);
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public optionSets(): string[][] | undefined {
    return [
      ['id', 'name']
    ];
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new SearchExternalConnectionGetCommand();