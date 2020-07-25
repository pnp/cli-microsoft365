import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  hubSiteId: string;
}

class SpoHubSiteConnectCommand extends SpoCommand {
  public get name(): string {
    return `${commands.HUBSITE_CONNECT}`;
  }

  public get description(): string {
    return 'Connects the specified site collection to the given hub site';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getRequestDigest(args.options.url)
      .then((res: ContextInfo): Promise<void> => {
        const requestOptions: any = {
          url: `${args.options.url}/_api/site/JoinHubSite('${encodeURIComponent(args.options.hubSiteId)}')`,
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
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'The URL of the site collection to connect to the hub site'
      },
      {
        option: '-i, --hubSiteId <hubSiteId>',
        description: 'The ID of the hub site to which to connect the site collection'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!Utils.isValidGuid(args.options.hubSiteId)) {
        return `${args.options.hubSiteId} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new SpoHubSiteConnectCommand();