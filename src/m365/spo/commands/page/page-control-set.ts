import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ClientSideControl } from './ClientSideControl.js';
import { ClientSidePageProperties } from './ClientSidePageProperties.js';
import { Page } from './Page.js';
import { PageControl } from './PageControl.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  pageName: string;
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

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        webPartData: typeof args.options.webPartData !== 'undefined',
        webPartProperties: typeof args.options.webPartProperties !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-n, --pageName <pageName>'
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
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
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

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let pageName: string = args.options.pageName;
    if (args.options.pageName.indexOf('.aspx') < 0) {
      pageName += '.aspx';
    }

    try {
      let requestOptions: any = {
        url: `${args.options.webUrl}/_api/SitePages/Pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageName)}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.get<ClientSidePageProperties>(requestOptions);
      if (!res.CanvasContent1) {
        throw `Page ${pageName} doesn't contain canvas controls.`;
      }

      const pageControls: PageControl[] = JSON.parse(res.CanvasContent1);
      const control: PageControl | undefined = pageControls.find(control => control.id && control.id.toLowerCase() === args.options.id.toLowerCase());

      if (!control) {
        throw `Control with ID ${args.options.id} not found on page ${pageName}`;
      }

      if (this.verbose) {
        await logger.logToStderr(`Control with ID ${args.options.id} found on the page`);
      }

      // Check out the page
      const page = await Page.checkout(pageName, args.options.webUrl, logger, this.verbose);
      // Update the web part data
      const canvasContent: ClientSideControl[] = JSON.parse(page.CanvasContent1);
      if (this.debug) {
        await logger.logToStderr(canvasContent);
      }

      const canvasControl = canvasContent.find(c => c.id.toLowerCase() === args.options.id.toLowerCase());
      if (!canvasControl) {
        throw `Control with ID ${args.options.id} not found on page ${pageName}`;
      }

      if (args.options.webPartData) {
        if (this.verbose) {
          await logger.logToStderr('web part data:');
          await logger.logToStderr(args.options.webPartData);
          await logger.logToStderr('');
        }

        const webPartData = JSON.parse(args.options.webPartData);
        canvasControl.webPartData = {
          ...canvasControl.webPartData,
          ...webPartData,
          id: canvasControl.webPartData.id,
          instanceId: canvasControl.webPartData.instanceId
        };

        if (this.verbose) {
          await logger.logToStderr('Updated web part data:');
          await logger.logToStderr(canvasControl.webPartData);
          await logger.logToStderr('');
        }
      }

      if (args.options.webPartProperties) {
        if (this.verbose) {
          await logger.logToStderr('web part properties data:');
          await logger.logToStderr(args.options.webPartProperties);
          await logger.logToStderr('');
        }

        const webPartProperties = JSON.parse(args.options.webPartProperties);
        canvasControl.webPartData.properties = {
          ...canvasControl.webPartData.properties,
          ...webPartProperties
        };

        if (this.verbose) {
          await logger.logToStderr('Updated web part properties:');
          await logger.logToStderr(canvasControl.webPartData.properties);
          await logger.logToStderr('');
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

      requestOptions = {
        url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageName)}')/SavePageAsDraft`,
        headers: {
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': '*',
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        data: pageData,
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageControlSetCommand();
