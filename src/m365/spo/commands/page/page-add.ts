import { Auth } from '../../../../Auth';
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
import { ClientSidePageProperties } from './ClientSidePageProperties';
import { Page } from './Page';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
  layoutType?: string;
  promoteAs?: string;
  commentsEnabled: boolean;
  publish: boolean;
  publishMessage?: string;
  description?: string;
  title?: string;
}

class SpoPageAddCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_ADD;
  }

  public get description(): string {
    return 'Creates modern page';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.layoutType = args.options.layoutType;
    telemetryProps.promoteAs = args.options.promoteAs;
    telemetryProps.commentsEnabled = args.options.commentsEnabled || false;
    telemetryProps.publish = args.options.publish || false;
    telemetryProps.publishMessage = typeof args.options.publishMessage !== 'undefined';
    telemetryProps.description = typeof args.options.description !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const resource = Auth.getResourceFromUrl(args.options.webUrl);
    let requestDigest: string = '';
    let itemId: string = '';
    let pageName: string = args.options.name;
    const serverRelativeSiteUrl: string = Utils.getServerRelativeSiteUrl(args.options.webUrl);
    const fileNameWithoutExtension: string = pageName.replace('.aspx', '');
    let bannerImageUrl: string = '';
    let canvasContent1: string = '';
    let layoutWebpartsContent: string = '';
    const pageTitle: string = args.options.title ? args.options.title : (args.options.name.indexOf('.aspx') > -1 ? args.options.name.substr(0, args.options.name.indexOf('.aspx')) : args.options.name);
    let pageId: number | null = null;
    const pageDescription: string = args.options.description || "";

    this
      .getRequestDigest(args.options.webUrl)
      .then((res: ContextInfo): Promise<{ UniqueId: string }> => {
        requestDigest = res.FormDigestValue;

        if (!pageName.endsWith('.aspx')) {
          pageName += '.aspx';
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfolderbyserverrelativeurl('${serverRelativeSiteUrl}/sitepages')/files/AddTemplateFile`,
          headers: {
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          data: {
            urlOfFile: `${serverRelativeSiteUrl}/sitepages/${pageName}`,
            templateFileType: 3
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((res: { UniqueId: string }): Promise<void> => {
        itemId = res.UniqueId;
        const layoutType: string = args.options.layoutType || 'Article';

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyid('${itemId}')/ListItemAllFields`,
          headers: {
            'X-RequestDigest': requestDigest,
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          data: {
            ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
            Title: pageTitle,
            ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
            PageLayoutType: layoutType
          },
          responseType: 'json'
        };

        if (layoutType === 'Article') {
          requestOptions.data.PromotedState = 0;
          requestOptions.data.BannerImageUrl = {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `${resource}/_layouts/15/images/sitepagethumbnail.png`
          };
        }

        return request.post(requestOptions);
      })
      .then((): Promise<ClientSidePageProperties> => {
        return Page.checkout(pageName, args.options.webUrl, logger, this.debug, this.verbose);
      })
      .then((res: ClientSidePageProperties): Promise<{ Id: string }> => {
        if (res) {
          pageId = res.Id;

          bannerImageUrl = res.BannerImageUrl;
          canvasContent1 = res.CanvasContent1;
          layoutWebpartsContent = res.LayoutWebpartsContent;
        }

        if (!args.options.promoteAs) {
          return Promise.resolve({ Id: '' });
        }

        const requestOptions: any = {
          responseType: 'json'
        };

        switch (args.options.promoteAs) {
          case 'HomePage':
            requestOptions.url = `${args.options.webUrl}/_api/web/rootfolder`;
            requestOptions.headers = {
              'X-RequestDigest': requestDigest,
              'X-HTTP-Method': 'MERGE',
              'IF-MATCH': '*',
              'content-type': 'application/json;odata=nometadata',
              accept: 'application/json;odata=nometadata'
            };
            requestOptions.data = {
              WelcomePage: `SitePages/${pageName}`
            };
            break;
          case 'NewsPage':
            requestOptions.url = `${args.options.webUrl}/_api/web/getfilebyid('${itemId}')/ListItemAllFields`;
            requestOptions.headers = {
              'X-RequestDigest': requestDigest,
              'X-HTTP-Method': 'MERGE',
              'IF-MATCH': '*',
              'content-type': 'application/json;odata=nometadata',
              accept: 'application/json;odata=nometadata'
            };
            requestOptions.data = {
              PromotedState: 2,
              FirstPublishedDate: new Date().toISOString().replace('Z', '')
            };
            break;
          case 'Template':
            requestOptions.url = `${args.options.webUrl}/_api/web/getfilebyid('${itemId}')/ListItemAllFields`;
            requestOptions.headers = {
              'X-RequestDigest': requestDigest,
              'content-type': 'application/json;odata=nometadata',
              accept: 'application/json;odata=nometadata'
            };
            break;
        }

        return request.post(requestOptions);
      })
      .then((res: { Id: string }): Promise<{ Id: number | null, BannerImageUrl: string, CanvasContent1: string, LayoutWebpartsContent: string, UniqueId: string }> => {
        if (args.options.promoteAs !== 'Template') {
          return Promise.resolve({ Id: null, BannerImageUrl: '', CanvasContent1: '', LayoutWebpartsContent: '', UniqueId: '' });
        }

        const requestOptions: any = {
          responseType: 'json',
          url: `${args.options.webUrl}/_api/SitePages/Pages(${res.Id})/SavePageAsTemplate`,
          headers: {
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            'X-HTTP-Method': 'POST',
            'IF-MATCH': '*',
            accept: 'application/json;odata=nometadata'
          }
        };

        return request.post(requestOptions);
      })
      .then((res: { Id: number | null, BannerImageUrl: string, CanvasContent1: string, LayoutWebpartsContent: string, UniqueId: string }): Promise<void> => {
        if (args.options.promoteAs !== 'Template') {
          return Promise.resolve();
        }

        bannerImageUrl = res.BannerImageUrl;
        canvasContent1 = res.CanvasContent1;
        layoutWebpartsContent = res.LayoutWebpartsContent;
        pageId = res.Id;

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyid('${res.UniqueId}')/ListItemAllFields/SetCommentsDisabled(${!args.options.commentsEnabled})`,
          headers: {
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((): Promise<void> => {
        const requestOptions: any = {
          responseType: 'json',
          url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePage`,
          headers: {
            'X-RequestDigest': requestDigest,
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          data: {
            BannerImageUrl: bannerImageUrl,
            CanvasContent1: canvasContent1,
            LayoutWebpartsContent: layoutWebpartsContent,
            Description: pageDescription
          }
        };

        return request.post(requestOptions);
      })
      .then((): Promise<void> => {
        if (args.options.promoteAs !== 'Template') {
          return Promise.resolve();
        }

        const requestOptions: any = {
          responseType: 'json',
          url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePageAsDraft`,
          headers: {
            'X-RequestDigest': requestDigest,
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          data: {
            Title: fileNameWithoutExtension,
            BannerImageUrl: bannerImageUrl,
            CanvasContent1: canvasContent1,
            LayoutWebpartsContent: layoutWebpartsContent,
            Description: pageDescription
          }
        };

        return request.post(requestOptions);
      })
      .then((): Promise<void> => {
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyid('${itemId}')/ListItemAllFields/SetCommentsDisabled(${!args.options.commentsEnabled})`,
          headers: {
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((): Promise<void> => {
        let requestOptions: any = {};

        if (!args.options.publish) {
          if (args.options.promoteAs === 'Template' || !pageId) {
            return Promise.resolve();
          }

          requestOptions = {
            responseType: 'json',
            url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePageAsDraft`,
            headers: {
              'content-type': 'application/json;odata=nometadata',
              'accept': 'application/json;odata=nometadata'
            },
            data: {
              Title: pageTitle,
              Description: pageDescription,
              BannerImageUrl: bannerImageUrl,
              CanvasContent1: canvasContent1,
              LayoutWebpartsContent: layoutWebpartsContent
            }
          };
        }
        else {
          requestOptions = {
            url: `${args.options.webUrl}/_api/web/getfilebyid('${itemId}')/CheckIn(comment=@a1,checkintype=@a2)?@a1='${encodeURIComponent(args.options.publishMessage || '').replace(/'/g, '%39')}'&@a2=1`,
            headers: {
              'X-RequestDigest': requestDigest,
              'content-type': 'application/json;odata=nometadata',
              accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };
        }

        return request.post(requestOptions);
      })
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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
        option: '-t, --title [title]'
      },
      {
        option: '-l, --layoutType [layoutType]',
        autocomplete: ['Article', 'Home']
      },
      {
        option: '-p, --promoteAs [promoteAs]',
        autocomplete: ['HomePage', 'NewsPage', 'Template']
      },
      {
        option: '--commentsEnabled'
      },
      {
        option: '--publish'
      },
      {
        option: '--publishMessage [publishMessage]'
      },
      {
        option: '--description [description]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (args.options.layoutType &&
      args.options.layoutType !== 'Article' &&
      args.options.layoutType !== 'Home') {
      return `${args.options.layoutType} is not a valid option for layoutType. Allowed values Article|Home`;
    }

    if (args.options.promoteAs &&
      args.options.promoteAs !== 'HomePage' &&
      args.options.promoteAs !== 'NewsPage' &&
      args.options.promoteAs !== 'Template') {
      return `${args.options.promoteAs} is not a valid option for promoteAs. Allowed values HomePage|NewsPage|Template`;
    }

    if (args.options.promoteAs === 'HomePage' && args.options.layoutType !== 'Home') {
      return 'You can only promote home pages as site home page';
    }

    if (args.options.promoteAs === 'NewsPage' && args.options.layoutType === 'Home') {
      return 'You can only promote article pages as news article';
    }

    return true;
  }
}

module.exports = new SpoPageAddCommand();
