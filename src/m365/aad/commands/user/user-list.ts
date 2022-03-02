import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  properties?: string;
  deleted?: boolean;
}

class AadUserListCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_LIST;
  }

  public get description(): string {
    return 'Lists users matching specified criteria';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.properties = args.options.properties;
    telemetryProps.deleted = typeof args.options.deleted !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const properties: string[] = args.options.properties ?
      args.options.properties.split(',').map(p => p.trim()) :
      ['userPrincipalName', 'displayName'];
    const filter: string = this.getFilter(args.options);
    const endpoint: string = args.options.deleted ? 'directory/deletedItems/microsoft.graph.user' : 'users';
    const url: string = `${this.resource}/v1.0/${endpoint}?$select=${properties.join(',')}${(filter.length > 0 ? '&' + filter : '')}&$top=100`;

    odata
      .getAllItems<User>(url, logger)
      .then((users): void => {
        logger.log(users);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getFilter(options: any): string {
    const filters: any = {};
    const excludeOptions: string[] = [
      'properties',
      'p',
      'deleted',
      'd',
      'debug',
      'verbose',
      'output',
      'o',
      'query',
      '_'
    ];

    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        filters[key] = encodeURIComponent(options[key].replace(/'/g, `''`));
      }
    });
    let filter: string = Object.keys(filters).map(key => `startsWith(${key}, '${filters[key]}')`).join(' and ');
    if (filter.length > 0) {
      filter = '$filter=' + filter;
    }

    return filter;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '-p, --properties [properties]' },
      { option: '-d, --deleted' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new AadUserListCommand();
