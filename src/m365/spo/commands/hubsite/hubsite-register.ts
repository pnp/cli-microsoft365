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
}

class SpoHubSiteRegisterCommand extends SpoCommand {
  public get name(): string {
    return `${commands.HUBSITE_REGISTER}`;
  }

  public get description(): string {
    return 'Registers the specified site collection as a hub site';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getRequestDigest(args.options.url)
      .then((res: ContextInfo): Promise<any> => {
        const requestOptions: any = {
          url: `${args.options.url}/_api/site/RegisterHubSite`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        cmd.log(res);

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
        description: 'URL of the site collection to register as a hub site'
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

module.exports = new SpoHubSiteRegisterCommand();