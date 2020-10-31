import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import { CommandError, CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSidePage } from './clientsidepages';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.log(`Retrieving information about the page...`);
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
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
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

        let page: any = {
          commentsDisabled: res.ListItemAllFields.CommentsDisabled,
          numSections: clientSidePage.sections.length,
          numControls: numControls,
          title: res.ListItemAllFields.Title
        };

        if (res.ListItemAllFields.PageLayoutType) {
          page.layoutType = res.ListItemAllFields.PageLayoutType;
        }

        if (args.options.output === 'json') {
          page = Object.assign(res, page);
        }

        logger.log(page);

        if (this.verbose) {
          logger.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoPageGetCommand();
