import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  status?: string;
  owner?: string;
  pageNumber?: string;
  pageSize?: string;
}

class GraphSchemaExtensionListCommand extends GraphCommand {
  public get name(): string {
    return commands.SCHEMAEXTENSION_LIST;
  }

  public get description(): string {
    return 'Get a list of schemaExtension objects created in the current tenant, that can be InDevelopment, Available, or Deprecated.';
  }
  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.status = typeof args.options.status !== 'undefined';
    telemetryProps.owner = typeof args.options.owner !== 'undefined';
    telemetryProps.pageNumber = typeof args.options.pageNumber !== 'undefined';
    telemetryProps.pageSize = typeof args.options.pageSize !== 'undefined';
    return telemetryProps;
  }
  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const filter: string = this.getFilter(args.options);
    let url = `${this.resource}/v1.0/schemaExtensions?$select=*${(filter.length > 0 ? '&' + filter : '')}`;

    if (args.options.pageNumber && Number(args.options.pageNumber) > 0) {
      const rowLimit: string = `&$top=${Number(args.options.pageSize ? args.options.pageSize : 10) * Number(args.options.pageNumber + 1)}`;
      url += rowLimit;
    }
    const requestOptions: any = {
      url: url,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };
    request.get(requestOptions)
      .then((res: any): void => {
        if (res.value && res.value.length > 0) {
          const size = args.options.pageSize ? parseInt(args.options.pageSize) : parseInt(res.value.length);
          const result = res.value.slice(-size);
          if (args.options.output !== 'json' && result.length > 1) {
            logger.log(result.map((x: any) => ({
              id: x.id,
              description: x.description,
              targetTypes: x.targetTypes,
              status: x.status,
              owner: x.owner,
              properties: JSON.stringify(x.properties)
            })));
          } else {
            logger.log(result);
          }
          if (this.verbose) {
            logger.logToStderr(chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }


  private getFilter(options: any): string {
    const filters: any = {};
    const filterOptions: string[] = [
      'status',
      'owner'
    ];

    Object.keys(options).forEach(key => {
      if (filterOptions.indexOf(key) !== -1) {
        filters[key] = options[key].replace(/'/g, `''`);
      }
    });
    let filter: string = Object.keys(filters).map(key => `${key} eq '${filters[key]}'`).join(' and ');
    if (filter.length > 0) {
      filter = '$filter=' + filter;
    }

    return filter;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-s, --status [status]',
        autocomplete: ['Available', 'InDevelopment', 'Deprecated']
      },
      {
        option: '--owner [owner]'
      },
      {
        option: '-p, --pageSize [pageSize]'
      },
      {
        option: '-n, --pageNumber [pageNumber]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.owner && !Utils.isValidGuid(args.options.owner)) {
      return `${args.options.owner} is not a valid GUID`;
    }
    if (args.options.pageNumber && parseInt(args.options.pageNumber) < 1) {
      return 'pageNumber must be a positive number';
    }
    if (args.options.pageSize && parseInt(args.options.pageSize) < 1) {
      return 'pageSize must be a positive number';
    }
    if (args.options.status &&
      ['Available', 'InDevelopment', 'Deprecated'].indexOf(args.options.status) === -1) {
      return `${args.options.status} is not a valid status value. Allowed values are Available|InDevelopment|Deprecated`;
    }
    return true;
  }
}
module.exports = new GraphSchemaExtensionListCommand();