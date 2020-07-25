import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption
} from '../../../../Command';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
          cmd.log(chalk.green('DONE'));
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
}

module.exports = new AadUserListCommand();
