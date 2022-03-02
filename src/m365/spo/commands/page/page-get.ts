import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { urlUtil, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
  metadataOnly?: boolean;
}

class SpoPageGetCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_GET;
  }

  public get description(): string {
    return 'Gets information about the specific modern page';
  }

  public defaultProperties(): string[] | undefined {
    return ['commentsDisabled', 'numSections', 'numControls', 'title', 'layoutType'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information about the page...`);
    }

    let pageName: string = args.options.name;
    if (args.options.name.indexOf('.aspx') < 0) {
      pageName += '.aspx';
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/SitePages/${encodeURIComponent(pageName)}')?$expand=ListItemAllFields/ClientSideApplicationId,ListItemAllFields/PageLayoutType,ListItemAllFields/CommentsDisabled`,
      headers: {
        'content-type': 'application/json;charset=utf-8',
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    let pageItemData: any = {};

    request
      .get(requestOptions)
      .then((res: any): Promise<{ CanvasContent1: string } | void> => {
        if (res.ListItemAllFields.ClientSideApplicationId !== 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
          return Promise.reject(`Page ${args.options.name} is not a modern page.`);
        }

        pageItemData = Object.assign({}, res);
        pageItemData.commentsDisabled = res.ListItemAllFields.CommentsDisabled;
        pageItemData.title = res.ListItemAllFields.Title;

        if (res.ListItemAllFields.PageLayoutType) {
          pageItemData.layoutType = res.ListItemAllFields.PageLayoutType;
        }

        if (args.options.metadataOnly) {
          return Promise.resolve();
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/SitePages/Pages(${res.ListItemAllFields.Id})`,
          headers: {
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get<{ CanvasContent1: string }>(requestOptions);
      })
      .then((res: { CanvasContent1: string } | void) => {
        if (res && res.CanvasContent1) {
          const canvasData: any[] = JSON.parse(res.CanvasContent1);
          pageItemData.canvasContentJson = res.CanvasContent1;
          if (canvasData && canvasData.length > 0) {
            pageItemData.numControls = canvasData.length;
            const sections = [...new Set(canvasData.filter(c => c.position).map(c => c.position.zoneIndex))];
            pageItemData.numSections = sections.length;
          }
        }

        logger.log(pageItemData);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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
        option: '--metadataOnly'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return validation.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoPageGetCommand();
