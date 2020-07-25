import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
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
      json: true
    };
    request.get(requestOptions)
      .then((res: any): void => {
        if (res.value && res.value.length > 0) {
          const size = args.options.pageSize ? parseInt(args.options.pageSize) : parseInt(res.value.length);
          const result = res.value.slice(-size);
          if(args.options.output !== 'json' && result.length > 1) {
            cmd.log(result.map((x: any) => ({
              id: x.id, 
              description: x.description,
              targetTypes: x.targetTypes,
              status: x.status,
              owner: x.owner,
              properties: JSON.stringify(x.properties)
            })));
          } else {
            cmd.log(result);
          }
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
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
        description: 'The status to filter on. Available values are Available, InDevelopment, Deprecated',
        autocomplete: ['Available', 'InDevelopment', 'Deprecated']
      },
      {
        option: '--owner [owner]',
        description: 'The id of the owner to filter on'
      },
      {
        option: '-p, --pageSize [pageSize]',
        description: 'Number of objects to return'
      },
      {
        option: '-n, --pageNumber [pageNumber]',
        description: 'Page number to return if pageSize is specified (first page is indexed as value of 0)'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
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
    };
  }
}
module.exports = new GraphSchemaExtensionListCommand();