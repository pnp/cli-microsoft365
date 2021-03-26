import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSideControl } from './ClientSideControl';
import { ClientSidePageProperties } from './ClientSidePageProperties';
import { Page } from './Page';
import { PageControl } from './PageControl';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  name: string;
  webUrl: string;
  webPartData?: string;
  webPartProperties?: string;
}

class SpoPageControlSetCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_CONTROL_SET;
  }

  public get description(): string {
    return 'Updates web part data or properties of a control on a modern page';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let pageName: string = args.options.name;
    if (args.options.name.indexOf('.aspx') < 0) {
      pageName += '.aspx';
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/SitePages/Pages/GetByUrl('sitepages/${encodeURIComponent(pageName)}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<ClientSidePageProperties>(requestOptions)
      .then((res: ClientSidePageProperties): Promise<ClientSidePageProperties> => {
        if (!res.CanvasContent1) {
          return Promise.reject(`Page ${pageName} doesn't contain canvas controls.`);
        }

        const canvasContent: PageControl[] = JSON.parse(res.CanvasContent1);
        const control: PageControl | undefined = canvasContent.find(control => control.id && control.id.toLowerCase() === args.options.id.toLowerCase());

        if (!control) {
          return Promise.reject(`Control with ID ${args.options.id} not found on page ${pageName}`);
        }

        if (this.verbose) {
          logger.logToStderr(`Control with ID ${args.options.id} found on the page`);
        }

        // Check out the page
        return Page.checkout(pageName, args.options.webUrl, logger, this.debug, this.verbose);
      })
      .then((page: ClientSidePageProperties) => {
        // Update the web part data
        const canvasContent: ClientSideControl[] = JSON.parse(page.CanvasContent1);
        if (this.debug) {
          logger.logToStderr(canvasContent);
        }

        const canvasControl = canvasContent.find(c => c.id.toLowerCase() === args.options.id.toLowerCase());
        if (!canvasControl) {
          return Promise.reject(`Control with ID ${args.options.id} not found on page ${pageName}`);
        }

        if (args.options.webPartData) {
          if (this.verbose) {
            logger.logToStderr('web part data:');
            logger.logToStderr(args.options.webPartData);
            logger.logToStderr('');
          }

          const webPartData = JSON.parse(args.options.webPartData);
          canvasControl.webPartData = {
            ...canvasControl.webPartData,
            ...webPartData,
            id: canvasControl.webPartData.id,
            instanceId: canvasControl.webPartData.instanceId
          };

          if (this.verbose) {
            logger.logToStderr('Updated web part data:');
            logger.logToStderr(canvasControl.webPartData);
            logger.logToStderr('');
          }
        }

        if (args.options.webPartProperties) {
          if (this.verbose) {
            logger.logToStderr('web part properties data:');
            logger.logToStderr(args.options.webPartProperties);
            logger.logToStderr('');
          }

          const webPartProperties = JSON.parse(args.options.webPartProperties);
          canvasControl.webPartData.properties = {
            ...canvasControl.webPartData.properties,
            ...webPartProperties
          };

          if (this.verbose) {
            logger.logToStderr('Updated web part properties:');
            logger.logToStderr(canvasControl.webPartData.properties);
            logger.logToStderr('');
          }
        }

        const pageData: any = {};

        if (page.AuthorByline) {
          pageData.AuthorByline = page.AuthorByline;
        }
        if (page.BannerImageUrl) {
          pageData.BannerImageUrl = page.BannerImageUrl;
        }
        if (page.Description) {
          pageData.Description = page.Description;
        }
        if (page.Title) {
          pageData.Title = page.Title;
        }
        if (page.TopicHeader) {
          pageData.TopicHeader = page.TopicHeader;
        }
        if (page.LayoutWebpartsContent) {
          pageData.LayoutWebpartsContent = page.LayoutWebpartsContent;
        }
        if (canvasContent) {
          pageData.CanvasContent1 = JSON.stringify(canvasContent);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${encodeURIComponent(pageName)}')/SavePageAsDraft`,
          headers: {
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          data: pageData,
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then(_ => cb())
      .catch((err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--webPartData [webPartData]'
      },
      {
        option: '--webPartProperties [webPartProperties]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    if (args.options.webPartData && args.options.webPartProperties) {
      return 'Specify webPartProperties or webPartData but not both';
    }

    if (args.options.webPartProperties) {
      try {
        JSON.parse(args.options.webPartProperties);
      }
      catch (e) {
        return `Specified webPartProperties is not a valid JSON string. Input: ${args.options.webPartData}. Error: ${e}`;
      }
    }

    if (args.options.webPartData) {
      try {
        JSON.parse(args.options.webPartData);
      }
      catch (e) {
        return `Specified webPartData is not a valid JSON string. Input: ${args.options.webPartData}. Error: ${e}`;
      }
    }

    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoPageControlSetCommand();
