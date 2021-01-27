import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ContextInfo } from '../../spo';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getRequestDigest(args.options.url)
      .then((res: ContextInfo): Promise<void> => {
        const requestOptions: any = {
          url: `${args.options.url}/_api/site/JoinHubSite('${encodeURIComponent(args.options.hubSiteId)}')`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>'
      },
      {
        option: '-i, --hubSiteId <hubSiteId>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (!Utils.isValidGuid(args.options.hubSiteId)) {
      return `${args.options.hubSiteId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new SpoHubSiteConnectCommand();