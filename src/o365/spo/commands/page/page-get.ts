import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate, CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { PageItem } from './PageItem';
import { ClientSidePage } from './clientsidepages';
import Utils from '../../../../Utils';
//import { relative } from 'path';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
}

class SpoPageGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_GET}`;
  }

  public get description(): string {
    return 'Gets information about the specific modern page';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving information about the page...`);
    }

    let pageName: string = args.options.name;
    if (args.options.name.indexOf('.aspx') < 0) {
      pageName += '.aspx';
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${Utils.getServerRelativeSiteUrl(args.options.webUrl)}/SitePages/${encodeURIComponent(pageName)}')?$expand=ListItemAllFields/ClientSideApplicationId,ListItemAllFields/PageLayoutType,ListItemAllFields/CommentsDisabled`,
      headers: {
        'content-type': 'application/json;charset=utf-8',
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<PageItem>(requestOptions)
      .then((res: PageItem): void => {
        if (res.ListItemAllFields.ClientSideApplicationId !== 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
          cb(new CommandError(`Page ${args.options.name} is not a modern page.`));
          return;
        }

        const clientSidePage: ClientSidePage = ClientSidePage.fromHtml(res.ListItemAllFields.CanvasContent1);
        let numControls: number = 0;
        clientSidePage.sections.forEach(s => {
          s.columns.forEach(c => {
            numControls += c.controls.length;
          });
        });

        const page: any = {
          commentsDisabled: res.ListItemAllFields.CommentsDisabled,
          numSections: clientSidePage.sections.length,
          numControls: numControls,
          title: res.ListItemAllFields.Title
        };

        if (res.ListItemAllFields.PageLayoutType) {
          page.layoutType = res.ListItemAllFields.PageLayoutType;
        }

        cmd.log(page);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'Name of the page to retrieve'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to retrieve is located'
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

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    If the specified ${chalk.grey('name')} doesn't refer to an existing modern page, you will get
    a ${chalk.grey('File doesn\'t exists')} error.

  Examples:

    Get information about the modern page with name ${chalk.grey('home.aspx')}
      ${this.name} --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx
`);
  }
}

module.exports = new SpoPageGetCommand();
