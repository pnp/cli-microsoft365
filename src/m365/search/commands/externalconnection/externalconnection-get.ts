import { Logger } from '../../../../cli';
import { ExternalConnectors } from '@microsoft/microsoft-graph-types';
import GraphCommand from '../../../base/GraphCommand';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import { CommandOption } from '../../../../Command';
import request from '../../../../request';

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
    return 'Retrieves the specified external connections';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let endpoint: string = `${this.resource}/v1.0/external/connections`;
    
    if (args.options.id) {
      endpoint += `/${encodeURIComponent(args.options.id)}`;
    }
    else {
      endpoint += `?$filter=name eq '${encodeURIComponent(args.options.name as string)}'`;
    }

    const requestOptions: any = {
      url: endpoint,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request.get<ExternalConnectors.ExternalConnection>(requestOptions)
      .then((res: ExternalConnectors.ExternalConnection): void => {
        logger.log(res);
        cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, logger, cb);
      });
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

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.name && !args.options.id) {
      return 'Specify either name or id';
    }

    if (args.options.name && args.options.id) {
      return 'Specify either name or id but not both';
    }

    return true;
  }
}

module.exports = new SearchExternalConnectionGetCommand();