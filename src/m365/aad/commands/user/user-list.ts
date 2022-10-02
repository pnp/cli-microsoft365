import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
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

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        properties: args.options.properties,
        deleted: typeof args.options.deleted !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-p, --properties [properties]' },
      { option: '-d, --deleted' }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let filter: string = '';
      const properties: string[] = args.options.properties ?
        args.options.properties.split(',').map(p => p.trim()) :
        ['userPrincipalName', 'displayName'];
      try {
        filter = this.getFilter(args.options);
      }
      catch (ex: any) {
        throw ex;
      }
      const endpoint: string = args.options.deleted ? 'directory/deletedItems/microsoft.graph.user' : 'users';
      const url: string = `${this.resource}/v1.0/${endpoint}?$select=${properties.join(',')}${(filter.length > 0 ? '&' + filter : '')}&$top=100`;
      const users = await odata.getAllItems<User>(url);
      logger.log(users);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
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
        if (typeof options[key] === 'boolean') {
          throw `Specify value for the ${key} property`;
        }

        filters[key] = encodeURIComponent(options[key].toString().replace(/'/g, `''`));
      }
    });
    let filter: string = Object.keys(filters).map(key => `startsWith(${key}, '${filters[key]}')`).join(' and ');
    if (filter.length > 0) {
      filter = '$filter=' + filter;
    }

    return filter;
  }
}

module.exports = new AadUserListCommand();
