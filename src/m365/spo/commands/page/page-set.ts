import * as chalk from 'chalk';
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
  commentsEnabled?: string;
  publish: boolean;
  publishMessage?: string;
  description?: string;
}

class SpoPageSetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_SET}`;
  }

  public get description(): string {
    return 'Updates modern page properties';
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
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let requestDigest: string = '';
    let pageName: string = args.options.name;
    let fileNameWithoutExtension: string = pageName.replace('.aspx', '');
    let bannerImageUrl: string = '';
    let canvasContent1: string = '';
    let layoutWebpartsContent: string = '';
    let pageTitle: string | null = null;
    let pageId: number | null = null;
    let pageDescription: string = args.options.description || "";

    if (!pageName.endsWith('.aspx')) {
      pageName += '.aspx';
    }
    const serverRelativeFileUrl: string = `${Utils.getServerRelativeSiteUrl(args.options.webUrl)}/sitepages/${pageName}`;

    this
      .getRequestDigest(args.options.webUrl)
      .then((res: ContextInfo): Promise<ClientSidePageProperties> => {
        requestDigest = res.FormDigestValue;

        return Page.checkout(args.options.name, args.options.webUrl, logger, this.debug, this.verbose);
      })
      .then((res: ClientSidePageProperties): Promise<void> => {
        if (res) {
          pageTitle = res.Title;
          pageId = res.Id;

          bannerImageUrl = res.BannerImageUrl;
          canvasContent1 = res.CanvasContent1;
          layoutWebpartsContent = res.LayoutWebpartsContent;
        }

        if (!args.options.layoutType) {
          return Promise.resolve();
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeFileUrl}')/ListItemAllFields`,
          headers: {
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            accept: 'application/json;odata=nometadata'
          },
          data: {
            PageLayoutType: args.options.layoutType
          },
          responseType: 'json'
        };

        if (args.options.layoutType === 'Article') {
          requestOptions.data.PromotedState = 0;
          requestOptions.data.BannerImageUrl = {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `${resource}/_layouts/15/images/sitepagethumbnail.png`
          };
        }

        return request.post(requestOptions);
      })
      .then((): Promise<{ Id: string }> => {
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
            requestOptions.url = `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeFileUrl}')/ListItemAllFields`;
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
            }
            break;
          case 'Template':
            requestOptions.url = `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeFileUrl}')/ListItemAllFields`;
            requestOptions.headers = {
              'X-RequestDigest': requestDigest,
              'content-type': 'application/json;odata=nometadata',
              accept: 'application/json;odata=nometadata'
            };
            break;
        }

        return request.post(requestOptions);
      })
      .then((res: { Id: string }): Promise<{ Id: number | null, BannerImageUrl: string, CanvasContent1: string, LayoutWebpartsContent: string }> => {
        if (args.options.promoteAs !== 'Template') {
          return Promise.resolve({ Id: null, BannerImageUrl: '', CanvasContent1: '', LayoutWebpartsContent: '' });
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
      .then((res: { Id: number | null, BannerImageUrl: string, CanvasContent1: string, LayoutWebpartsContent: string }): Promise<void> => {
        if (args.options.promoteAs !== 'Template') {
          return Promise.resolve();
        }

        bannerImageUrl = res.BannerImageUrl;
        canvasContent1 = res.CanvasContent1;
        layoutWebpartsContent = res.LayoutWebpartsContent;
        pageId = res.Id;

        let requestOptions: any = {
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
        let requestOptions: any = {};

        if (args.options.promoteAs !== 'Template') {
          if (!pageId) {
            return Promise.resolve();
          }

          requestOptions = {
            responseType: 'json',
            url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePage`,
            headers: {
              'content-type': 'application/json;odata=nometadata',
              'accept': 'application/json;odata=nometadata'
            },
            data: {
              Title: pageTitle,
              BannerImageUrl: bannerImageUrl,
              CanvasContent1: canvasContent1,
              LayoutWebpartsContent: layoutWebpartsContent,
              Description: pageDescription
            }
          };
        }
        else {
          requestOptions = {
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
        }

        return request.post(requestOptions);
      })
      .then((): Promise<void> => {
        if (typeof args.options.commentsEnabled === 'undefined') {
          return Promise.resolve();
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeFileUrl}')/ListItemAllFields/SetCommentsDisabled(${args.options.commentsEnabled === 'false'})`,
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
        if (!args.options.publish) {
          return Promise.resolve();
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeFileUrl}')/CheckIn(comment=@a1,checkintype=@a2)?@a1='${encodeURIComponent(args.options.publishMessage || '').replace(/'/g, '%39')}'&@a2=1`,
          headers: {
            'X-RequestDigest': requestDigest,
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
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'Name of the page to update'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to update is located'
      },
      {
        option: '-l, --layoutType [layoutType]',
        description: 'Layout of the page. Allowed values Article|Home',
        autocomplete: ['Article', 'Home']
      },
      {
        option: '-p, --promoteAs [promoteAs]',
        description: 'Update the page purpose. Allowed values HomePage|NewsPage|Template',
        autocomplete: ['HomePage', 'NewsPage', 'Template']
      },
      {
        option: '--commentsEnabled [commentsEnabled]',
        description: 'Set to true, to enable comments on the page. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--publish',
        description: 'Set to publish the page'
      },
      {
        option: '--publishMessage [publishMessage]',
        description: 'Message to set when publishing the page'
      },
      {
        option: '--description [description]',
        description: 'The description to set for the page'
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

    if (typeof args.options.commentsEnabled !== 'undefined' &&
      args.options.commentsEnabled !== 'true' &&
      args.options.commentsEnabled !== 'false') {
      return `${args.options.commentsEnabled} is not a valid value for commentsEnabled. Allowed values true|false`;
    }

    return true;
  }
}

module.exports = new SpoPageSetCommand();
