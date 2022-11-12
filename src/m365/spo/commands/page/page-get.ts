import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
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

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--metadataOnly'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information about the page...`);
    }

    let pageName: string = args.options.name;
    if (args.options.name.indexOf('.aspx') < 0) {
      pageName += '.aspx';
    }

    try {
      let requestOptions: any = {
        url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/SitePages/${formatting.encodeQueryParameter(pageName)}')?$expand=ListItemAllFields/ClientSideApplicationId,ListItemAllFields/PageLayoutType,ListItemAllFields/CommentsDisabled`,
        headers: {
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const page = await request.get<any>(requestOptions);

      if (page.ListItemAllFields.ClientSideApplicationId !== 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
        throw `Page ${args.options.name} is not a modern page.`;
      }

      let pageItemData: any = {};
      pageItemData = Object.assign({}, page);
      pageItemData.commentsDisabled = page.ListItemAllFields.CommentsDisabled;
      pageItemData.title = page.ListItemAllFields.Title;

      if (page.ListItemAllFields.PageLayoutType) {
        pageItemData.layoutType = page.ListItemAllFields.PageLayoutType;
      }

      if (!args.options.metadataOnly) {
        requestOptions = {
          url: `${args.options.webUrl}/_api/SitePages/Pages(${page.ListItemAllFields.Id})`,
          headers: {
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const res = await request.get<{ CanvasContent1: string }>(requestOptions);
        const canvasData: any[] = JSON.parse(res.CanvasContent1);
        pageItemData.canvasContentJson = res.CanvasContent1;
        if (canvasData && canvasData.length > 0) {
          pageItemData.numControls = canvasData.length;
          const sections = [...new Set(canvasData.filter(c => c.position).map(c => c.position.zoneIndex))];
          pageItemData.numSections = sections.length;
        }
      }

      logger.log(pageItemData);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoPageGetCommand();
