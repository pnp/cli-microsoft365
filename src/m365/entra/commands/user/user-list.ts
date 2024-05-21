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
  type?: string;
  properties?: string;
}

class EntraUserListCommand extends GraphCommand {
  private static readonly allowedTypes: string[] = ['Member', 'Guest'];

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

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'mail', 'userPrincipalName'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        type: typeof args.options.type !== 'undefined',
        properties: typeof args.options.properties !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--type [type]',
        autocomplete: EntraUserListCommand.allowedTypes
      },
      {
        option: '-p, --properties [properties]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.type && !EntraUserListCommand.allowedTypes.some(t => t.toLowerCase() === args.options.type!.toLowerCase())) {
          return `'${args.options.type}' is not a valid value for option 'type'. Allowed values are: ${EntraUserListCommand.allowedTypes.join(',')}.`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('type', 'properties');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.USER_LIST, commands.USER_LIST);

    try {
      let url = `${this.resource}/v1.0/users`;

      if (args.options.properties) {
        const selectProperties = args.options.properties;
        const allSelectProperties = selectProperties.split(',');
        const propertiesWithSlash = allSelectProperties.filter(item => item.includes('/'));

        const fieldExpand = propertiesWithSlash
          .map(p => `${p.split('/')[0]}($select=${p.split('/')[1]})`)
          .join(',');

        const expandParam = fieldExpand.length > 0 ? `&$expand=${fieldExpand}` : '';
        const selectParam = allSelectProperties.filter(item => !item.includes('/'));

        url += `?$select=${selectParam}${expandParam}`;
      }

      let filter: string = '';
      try {
        filter = this.getFilter(args.options);
        if (filter.length > 0) {
          url += `${args.options.properties ? '&' : '?'}${filter}`;
        }
      }
      catch (ex: any) {
        throw ex;
      }

      const users = await odata.getAllItems<User>(url);
      await logger.log(users);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getFilter(options: Options): string {
    const filters: any = {};
    const excludeOptions: string[] = [
      'type',
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
      filter = `$filter=${filter}`;
    }

    if (options.type) {
      filter += filter.length > 0 ? ` and userType eq '${options.type}'` : `$filter=userType eq '${options.type}'`;
    }

    return filter;
  }
}

export default new EntraUserListCommand();
