import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  properties?: string;
}

class EntraUserListCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_LIST;
  }

  public get description(): string {
    return 'Lists users matching specified criteria';
  }

  public alias(): string[] | undefined {
    return [aadCommands.USER_LIST];
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
        properties: args.options.properties
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-p, --properties [properties]' }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    this.showDeprecationWarning(logger, aadCommands.USER_LIST, commands.USER_LIST);

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

      const url: string = `${this.resource}/v1.0/users?$select=${properties.join(',')}${(filter.length > 0 ? '&' + filter : '')}&$top=100`;
      const users = await odata.getAllItems<User>(url);
      await logger.log(users);
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

        filters[key] = formatting.encodeQueryParameter(options[key].toString());
      }
    });
    let filter: string = Object.keys(filters).map(key => `startsWith(${key}, '${filters[key]}')`).join(' and ');
    if (filter.length > 0) {
      filter = '$filter=' + filter;
    }

    return filter;
  }
}

export default new EntraUserListCommand();
