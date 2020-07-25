import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  confirm?: boolean;
}

class SpoHubSiteUnregisterCommand extends SpoCommand {
  public get name(): string {
    return `${commands.HUBSITE_UNREGISTER}`;
  }

  public get description(): string {
    return 'Unregisters the specifies site collection as a hub site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = args.options.confirm || false;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const unregisterHubSite: () => void = (): void => {
      this
        .getRequestDigest(args.options.url)
        .then((res: ContextInfo): Promise<void> => {
          const requestOptions: any = {
            url: `${args.options.url}/_api/site/UnregisterHubSite`,
            headers: {
              'X-RequestDigest': res.FormDigestValue,
              accept: 'application/json;odata=nometadata'
            },
            json: true
          };

          return request.post(requestOptions);
        })
        .then((): void => {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }

          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      unregisterHubSite();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to unregister the site collection ${args.options.url} as a hub site?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          unregisterHubSite();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'URL of the site collection to unregister as a hub site'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming unregistering the hub site'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      return SpoCommand.isValidSharePointUrl(args.options.url);
    };
  }
}

module.exports = new SpoHubSiteUnregisterCommand();