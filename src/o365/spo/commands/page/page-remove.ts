import request from '../../../../request';
import commands from '../../commands';
import { CommandOption, CommandValidate } from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
            cmd.log(vorpal.chalk.green('DONE'));
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
      if (!args.options.name) {
        return 'Required parameter name missing';
      }

      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    If you try to remove a page with that does not exist, you will get
    a ${chalk.grey('The file does not exist')} error.

    If you set the ${chalk.grey('--confirm')}  flag, you will not be prompted for confirmation
    before the page is actually removed.

  Examples:

    Remove a modern page.
      ${this.name} --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team

    Remove a modern page without a confirmation prompt.
      ${this.name} --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --confirm
    `
    );
  }
}

module.exports = new SpoPageRemoveCommand();
