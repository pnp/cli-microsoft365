import config from '../../../../config';
import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import Utils from '../../../../Utils';
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  status?: string;
  owner?: string;
  pageNumber?: string;
  pageSize?: string;
}

class GraphSchemaExtensionList extends GraphCommand {
  public get name(): string {
    return `${commands.SCHEMAEXTENSION_LIST}`;
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
    ((): Promise<any> => {
      if (args.options.pageNumber && Number(args.options.pageNumber) > 0) {
        const rowLimit: string = `&$top=${Number(args.options.pageSize ? args.options.pageSize : 10) * Number(args.options.pageNumber)}`;
        url += rowLimit;
        const requestOptions: any = {
          url: url,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          },
          json: true
        };
        return request.get(requestOptions);
      } else {
        return Promise.resolve();
      }
  })().then((res:any):void =>{
    const requestOptions: any = {
      url: args.options.pageNumber && Number(args.options.pageNumber) > 0 ?res[`@odata.nextLink`] :url,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      json: true
    };
    request.get(requestOptions)
      .then((res: any): void => {
        cmd.log(res);
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  });
  }

  private getFilter(options: any): string {
    const filters: any = {};
    const excludeOptions: string[] = [
      'pageNumber',
      'pageSize',
      'debug',
      'verbose',
      'output'
    ];

    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        filters[key] = encodeURIComponent(options[key].replace(/'/g, `''`));
      }
    });
    let filter: string = Object.keys(filters).map(key => `eq(${key}, '${filters[key]}')`).join(' and ');
    if (filter.length > 0) {
      filter = '$filter=' + filter;
    }

    return filter;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-s, --status [status]',
        description: `The status to filter on`
      },
      {
        option: '--owner [owner]',
        description: `The id of the owner to filter on`
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
      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:

    Get a list of schemaExtension objects created in the current tenant, that can be InDevelopment, Available, or Deprecated.
      ${chalk.grey(config.delimiter)} ${this.name}

    Get a list of schemaExtension objects created in the current tenant, with owner 617720dc-85fc-45d7-a187-cee75eaf239e
      ${chalk.grey(config.delimiter)} ${this.name} --owner 617720dc-85fc-45d7-a187-cee75eaf239e

  Additional information:
    pageNumber is specified as a 0-based index. A value of 2 returns the third page of items. 
    
      More information: https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/schemaextension_list
    `
    );
  }
}
module.exports = new GraphSchemaExtensionList();