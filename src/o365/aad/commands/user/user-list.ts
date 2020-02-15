import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption
} from '../../../../Command';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  properties?: string;
}

class AadUserListCommand extends GraphItemsListCommand<any> {
  public get name(): string {
    return `${commands.USER_LIST}`;
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
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const properties: string[] = args.options.properties ?
      args.options.properties.split(',').map(p => p.trim()) :
      ['userPrincipalName', 'displayName'];
    const filter: string = this.getFilter(args.options);
    const url: string = `${this.resource}/v1.0/users?$select=${properties.join(',')}${(filter.length > 0 ? '&' + filter : '')}&$top=100`;

    this
      .getAllItems(url, cmd, true)
      .then((): void => {
        cmd.log(this.items);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getFilter(options: any): string {
    const filters: any = {};
    const excludeOptions: string[] = [
      'properties',
      'debug',
      'verbose',
      'output'
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
      {
        option: '-p, --properties [properties]',
        description: 'Comma-separated list of properties to retrieve'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    Using the ${chalk.blue('--properties')} option, you can specify
    a comma-separated list of user properties to retrieve from the Microsoft
    Graph. If you don't specify any properties, the command will retrieve
    user's display name and account name.

    To filter the list of users, include additional options that match the user
    property that you want to filter with. For example
    ${chalk.blue('--displayName Patt')} will return all users whose displayName
    starts with ${chalk.grey('Patt')}. Multiple filters will be combined using
    the ${chalk.blue('and')} operator.

  Examples:

    List all users in the tenant
      ${this.name}

    List all users in the tenant. For each one return the display name and
    e-mail address
      ${this.name} --properties "displayName,mail"

    Show users whose display name starts with ${chalk.grey('Patt')}
      ${this.name} --displayName Patt

    Show all account managers whose display name starts with ${chalk.grey('Patt')}
      ${this.name} --displayName Patt --jobTitle 'Account manager'

  More information:

    Microsoft Graph User properties
      https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/user#properties
`);
  }
}

module.exports = new AadUserListCommand();
