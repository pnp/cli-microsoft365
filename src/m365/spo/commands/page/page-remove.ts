import request from '../../../../request';
import commands from '../../commands';
import { CommandOption, CommandValidate } from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
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
            cmd.log(`Removing page ${pageName}...`);
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
            json: true
          };

          return request.post(requestOptions);
        })
        .then((): void => {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }
          cb();
        },
          (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb)
        );
    };

    if (args.options.confirm) {
      removePage();
    }
    else {
      cmd.prompt(
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
        option: '-n, --name <name>',
        description: 'Name of the page to remove'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site from which the page should be removed'
      },
      {
        option: '--confirm',
        description: `Don't prompt before removing the page`
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoPageRemoveCommand();
