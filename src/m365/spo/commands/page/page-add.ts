import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
  title?: string;
}

class SpoPageAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_ADD}`;
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
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let resource = Auth.getResourceFromUrl(args.options.webUrl);
    let requestDigest: string = '';
    let itemId: string = '';
    let pageName: string = args.options.name;
    const serverRelativeSiteUrl: string = Utils.getServerRelativeSiteUrl(args.options.webUrl);
    const fileNameWithoutExtension: string = pageName.replace('.aspx', '');
    let bannerImageUrl: string = '';
    let canvasContent1: string = '';
    let layoutWebpartsContent: string = '';
    let templateListItemId: string = '';

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
          body: {
            urlOfFile: `${serverRelativeSiteUrl}/sitepages/${pageName}`,
            templateFileType: 3
          },
          json: true
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
          body: {
            ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
            Title: args.options.title ? args.options.title : (args.options.name.indexOf('.aspx') > -1 ? args.options.name.substr(0, args.options.name.indexOf('.aspx')) : args.options.name),
            ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
            PageLayoutType: layoutType
          },
          json: true
        };

        if (layoutType === 'Article') {
          requestOptions.body.PromotedState = 0;
          requestOptions.body.BannerImageUrl = {
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
          json: true
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
            requestOptions.body = {
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
            requestOptions.body = {
              PromotedState: 2,
              FirstPublishedDate: new Date().toISOString().replace('Z', '')
            }
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
      .then((res: { Id: string }): Promise<{ Id: string, BannerImageUrl: string, CanvasContent1: string, LayoutWebpartsContent: string, UniqueId: string }> => {
        if (args.options.promoteAs !== 'Template') {
          return Promise.resolve({ Id: '', BannerImageUrl: '', CanvasContent1: '', LayoutWebpartsContent: '', UniqueId: '' });
        }

        const requestOptions: any = {
          json: true,
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
      .then((res: { Id: string, BannerImageUrl: string, CanvasContent1: string, LayoutWebpartsContent: string, UniqueId: string }): Promise<void> => {
        if (args.options.promoteAs !== 'Template') {
          return Promise.resolve();
        }

        bannerImageUrl = res.BannerImageUrl;
        canvasContent1 = res.CanvasContent1;
        layoutWebpartsContent = res.LayoutWebpartsContent;
        templateListItemId = res.Id;

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyid('${res.UniqueId}')/ListItemAllFields/SetCommentsDisabled(${!args.options.commentsEnabled})`,
          headers: {
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((): Promise<void> => {
        if (args.options.promoteAs !== 'Template') {
          return Promise.resolve();
        }

        const requestOptions: any = {
          json: true,
          url: `${args.options.webUrl}/_api/SitePages/Pages(${templateListItemId})/SavePage`,
          headers: {
            'X-RequestDigest': requestDigest,
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          body: {
            BannerImageUrl: bannerImageUrl,
            CanvasContent1: canvasContent1,
            LayoutWebpartsContent: layoutWebpartsContent
          }
        };
        return request.post(requestOptions);
      })
      .then((): Promise<void> => {
        if (args.options.promoteAs !== 'Template') {
          return Promise.resolve();
        }

        const requestOptions: any = {
          json: true,
          url: `${args.options.webUrl}/_api/SitePages/Pages(${templateListItemId})/SavePageAsDraft`,
          headers: {
            'X-RequestDigest': requestDigest,
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          body: {
            Title: fileNameWithoutExtension,
            BannerImageUrl: bannerImageUrl,
            CanvasContent1: canvasContent1,
            LayoutWebpartsContent: layoutWebpartsContent
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
          json: true
        };

        return request.post(requestOptions);
      })
      .then((): Promise<void> => {
        if (!args.options.publish) {
          return Promise.resolve();
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyid('${itemId}')/Publish('${encodeURIComponent(args.options.publishMessage || '').replace(/'/g, '%39')}')`,
          headers: {
            'X-RequestDigest': requestDigest,
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
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'Name of the page to create'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page should be created'
      },
      {
        option: '-t, --title [title]',
        description: 'Title of the page to create. If not specified, will use the page name as its title'
      },
      {
        option: '-l, --layoutType [layoutType]',
        description: 'Layout of the page. Allowed values Article|Home. Default Article',
        autocomplete: ['Article', 'Home']
      },
      {
        option: '-p, --promoteAs [promoteAs]',
        description: 'Create the page for a specific purpose. Allowed values HomePage|NewsPage|Template',
        autocomplete: ['HomePage', 'NewsPage', 'Template']
      },
      {
        option: '--commentsEnabled',
        description: 'Set to enable comments on the page'
      },
      {
        option: '--publish',
        description: 'Set to publish the page'
      },
      {
        option: '--publishMessage [publishMessage]',
        description: 'Message to set when publishing the page'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
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
    };
  }
}

module.exports = new SpoPageAddCommand();
