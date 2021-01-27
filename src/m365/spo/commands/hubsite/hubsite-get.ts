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

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class SpoHubSiteGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.HUBSITE_GET}`;
  }

  public get description(): string {
    return 'Gets information about the specified hub site';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getSpoUrl(logger, this.debug)
      .then((spoUrl: string): Promise<any> => {
        const requestOptions: any = {
          url: `${spoUrl}/_api/hubsites/getbyid('${encodeURIComponent(args.options.id)}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })
      .then((res: any): void => {
        logger.log(res);

        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new SpoHubSiteGetCommand();