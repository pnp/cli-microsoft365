import * as chalk from 'chalk';
import { Cli, Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
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
  name: string;
  webUrl: string;
  confirm?: boolean;
}

class SpoPageRemoveCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_REMOVE}`;
  }

  public get description(): string {
    return 'Removes a modern page';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let requestDigest: string = '';
    let pageName: string = args.options.name;

    const removePage = () => {
      this
        .getRequestDigest(args.options.webUrl)
        .then((res: ContextInfo): Promise<void> => {
          requestDigest = res.FormDigestValue;

          if (!pageName.endsWith('.aspx')) {
            pageName += '.aspx';
          }

          if (this.verbose) {
            logger.logToStderr(`Removing page ${pageName}...`);
          }

          const requestOptions: any = {
            url: `${args.options
              .webUrl}/_api/web/getfilebyserverrelativeurl('${Utils.getServerRelativeSiteUrl(args.options.webUrl)}/sitepages/${pageName}')`,
            headers: {
              'X-RequestDigest': requestDigest,
              'X-HTTP-Method': 'DELETE',
              'content-type': 'application/json;odata=nometadata',
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
        },
          (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb)
        );
    };

    if (args.options.confirm) {
      removePage();
    }
    else {
      Cli.prompt(
        {
          type: 'confirm',
          name: 'continue',
          default: false,
          message: `Are you sure you want to remove the page '${args.options.name}'?`
        },
        (result: { continue: boolean }): void => {
          if (!result.continue) {
            cb();
          }
          else {
            removePage();
          }
        }
      );
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>'
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoPageRemoveCommand();
