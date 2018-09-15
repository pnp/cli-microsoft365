import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import { CommandOption, CommandValidate } from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';
import {
  ClientSidePage,
  CanvasSection,
  ClientSideWebpart,
  ClientSidePageComponent,
  CanvasColumn
} from './clientsidepages';
import { StandardWebPart, StandardWebPartUtils } from '../../common/StandardWebPartTypes';
import { ContextInfo } from '../../spo';
import { isNumber } from 'util';
import { PageItem } from './PageItem';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  pageName: string;
  webUrl: string;
  standardWebPart?: StandardWebPart;
  webPartId?: string;
  webPartProperties?: string;
  section?: number;
  column?: number;
  order?: number;
}

class SpoPageClientSideWebPartAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_CLIENTSIDEWEBPART_ADD}`;
  }

  public get description(): string {
    return 'Adds a client-side web part to a modern page';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.standardWebPart = args.options.standardWebPart;
    telemetryProps.webPartId = typeof args.options.webPartId !== 'undefined';
    telemetryProps.webPartProperties = typeof args.options.webPartProperties !== 'undefined';
    telemetryProps.section = typeof args.options.section !== 'undefined';
    telemetryProps.column = typeof args.options.column !== 'undefined';
    telemetryProps.order = typeof args.options.order !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';
    let requestDigest: string = '';
    let clientSidePage: ClientSidePage | null = null;

    let pageName: string = args.options.pageName;
    if (args.options.pageName.indexOf('.aspx') < 0) {
      pageName += '.aspx';
    }

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        siteAccessToken = accessToken;

        if (this.verbose) {
          cmd.log(`Retrieving request digest...`);
        }

        return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
      })
      .then((res: ContextInfo): Promise<ClientSidePage> => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        // Keep the reference of request digest for subsequent requests
        requestDigest = res.FormDigestValue;

        if (this.verbose) {
          cmd.log(`Retrieving modern page ${pageName}...`);
        }
        // Get Client Side Page
        return this.getClientSidePage(pageName, cmd, args, siteAccessToken);
      })
      .then((page: ClientSidePage): Promise<ClientSideWebpart> => {
        // Keep the reference of client side page for subsequent requests
        clientSidePage = page;

        if (this.verbose) {
          cmd.log(
            `Retrieving definition for web part ${args.options.webPartId ||
            args.options.standardWebPart}...`
          );
        }
        // Get the WebPart according to arguments
        return this.getWebPart(cmd, args, siteAccessToken);
      })
      .then((webPart: ClientSideWebpart): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved WebPart definition:`);
          cmd.log(webPart);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(`Setting client-side web part layout and properties...`);
        }
        // Set the WebPart properties and layout (section, column and order)
        this.setWebPartPropertiesAndLayout(clientSidePage as ClientSidePage, webPart, cmd, args);
        if (this.verbose) {
          cmd.log(`Saving modern page...`);
        }
        // Save the Client Side Page with updated information
        return this.saveClientSidePage(clientSidePage as ClientSidePage, cmd, args, pageName, siteAccessToken, requestDigest);
      })
      .then((res: any): void => {
        if (this.debug) {
          cmd.log(`Response`);
          cmd.log(res);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
        cb();
      })
      .catch((err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getClientSidePage(
    pageName: string,
    cmd: CommandInstance,
    args: CommandArgs,
    accessToken: string
  ): Promise<ClientSidePage> {
    return new Promise<ClientSidePage>((resolve: (page: ClientSidePage) => void, reject: (error: any) => void): void => {
      if (this.verbose) {
        cmd.log(`Retrieving information about the page...`);
      }

      const webUrl: string = args.options.webUrl;
      const requestOptions: any = {
        url: `${webUrl}/_api/web/getfilebyserverrelativeurl('${webUrl.substr(
          webUrl.indexOf('/', 8)
        )}/sitepages/${encodeURIComponent(pageName)}')?$expand=ListItemAllFields/ClientSideApplicationId`,
        headers: Utils.getRequestHeaders({
          authorization: `Bearer ${accessToken}`,
          accept: 'application/json;odata=nometadata'
        }),
        json: true
      };

      if (this.debug) {
        cmd.log('Executing web request...');
        cmd.log(requestOptions);
        cmd.log('');
      }

      request
        .get(requestOptions)
        .then((res: PageItem): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          if (res.ListItemAllFields.ClientSideApplicationId !== 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
            reject(new Error(`Page ${args.options.pageName} is not a modern page.`));
            return;
          }

          const clientSidePage: ClientSidePage = ClientSidePage.fromHtml(
            res.ListItemAllFields.CanvasContent1
          );

          resolve(clientSidePage);
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  private getWebPart(cmd: CommandInstance, args: CommandArgs, accessToken: string): Promise<ClientSideWebpart> {
    return new Promise<ClientSideWebpart>((resolve: (webPart: ClientSideWebpart) => void, reject: (error: any) => void): void => {
      const standardWebPart: string | undefined = args.options.standardWebPart;

      const webPartId = standardWebPart
        ? StandardWebPartUtils.getWebPartId(standardWebPart as StandardWebPart)
        : args.options.webPartId;

      if (this.debug) {
        cmd.log(`StandardWebPart: ${standardWebPart}`);
        cmd.log(`WebPartId: ${webPartId}`);
      }

      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/web/getclientsidewebparts()`,
        headers: Utils.getRequestHeaders({
          authorization: `Bearer ${accessToken}`,
          accept: 'application/json;odata=nometadata'
        }),
        json: true
      };

      if (this.debug) {
        cmd.log('Executing web request...');
        cmd.log(requestOptions);
        cmd.log('');
      }

      request
        .get(requestOptions)
        .then((res: { value: ClientSidePageComponent[] }): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          const webPartDefinition = res.value.filter((c) => c.Id === webPartId);
          if (webPartDefinition.length === 0) {
            reject(new Error(`There is no available WebPart with Id ${webPartId}.`));
            return;
          }

          if (this.debug) {
            cmd.log('WebPart definition:');
            cmd.log(webPartDefinition);
            cmd.log('');
          }

          if (this.verbose) {
            cmd.log(`Creating instance from definition of WebPart ${webPartId}...`);
          }
          const webPart = ClientSideWebpart.fromComponentDef(webPartDefinition[0]);
          resolve(webPart);
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  private setWebPartPropertiesAndLayout(
    clientSidePage: ClientSidePage,
    webPart: ClientSideWebpart,
    cmd: CommandInstance,
    args: CommandArgs
  ): void {
    let actualSectionIndex: number | undefined = args.options.section && args.options.section - 1;
    // If the section arg is not specified, set to first section
    if (typeof actualSectionIndex === 'undefined') {
      if (this.debug) {
        cmd.log(`No section argument specified, The component will be added to default section`);
      }
      actualSectionIndex = 0;
    }
    // Make sure the section is in the range, other add a new section
    if (actualSectionIndex >= clientSidePage.sections.length) {
      throw new Error(`Invalid Section '${args.options.section}'`);
    }

    // Get the target section
    const section: CanvasSection = clientSidePage.sections[actualSectionIndex];

    let actualColumnIndex: number | undefined = args.options.column && args.options.column - 1;
    // If the column arg is not specified, set to first column
    if (typeof actualColumnIndex === 'undefined') {
      if (this.debug) {
        cmd.log(`No column argument specified, The component will be added to default column`);
      }
      actualColumnIndex = 0;
    }
    // Make sure the column is in the range of the current section
    if (actualColumnIndex >= section.columns.length) {
      throw new Error(`Invalid Column '${args.options.column}'`);
    }

    // Get the target column
    const column: CanvasColumn = section.columns[actualColumnIndex];

    if (args.options.webPartProperties) {
      if (this.debug) {
        cmd.log('WebPart properties: ');
        cmd.log(args.options.webPartProperties);
        cmd.log('');
      }

      try {
        const properties: any = JSON.parse(args.options.webPartProperties);
        webPart.setProperties(properties);
      }
      catch {
      }
    }

    // Add the WebPart to to appropriate location
    column.addControl(webPart);
  }

  private saveClientSidePage(
    clientSidePage: ClientSidePage,
    cmd: CommandInstance,
    args: CommandArgs,
    pageName: string,
    accessToken: string,
    requestDigest: string
  ): request.RequestPromise {
    const serverRelativeSiteUrl: string = `${args.options.webUrl.substr(
      args.options.webUrl.indexOf('/', 8)
    )}/sitepages/${pageName}`;

    const updatedContent: string = clientSidePage.toHtml();

    if (this.debug) {
      cmd.log('Updated canvas content: ');
      cmd.log(updatedContent);
      cmd.log('');
    }

    const requestOptions: any = {
      url: `${args.options
        .webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeSiteUrl}')/ListItemAllFields`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${accessToken}`,
        'X-RequestDigest': requestDigest,
        'content-type': 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*',
        accept: 'application/json;odata=nometadata'
      }),
      body: {
        CanvasContent1: updatedContent
      },
      json: true
    };

    if (this.debug) {
      cmd.log('Executing web request...');
      cmd.log(requestOptions);
      cmd.log('');
    }

    return request.post(requestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to add the web part to is located'
      },
      {
        option: '-n, --pageName <pageName>',
        description: 'Name of the page to which add the web part'
      },
      {
        option: '--standardWebPart [standardWebPart]',
        description: `Name of the standard web part to add (see the possible values below)`
      },
      {
        option: '--webPartId [webPartId]',
        description: 'ID of the custom web part to add'
      },
      {
        option: '--webPartProperties [webPartProperties]',
        description: 'JSON string with web part properties to set on the web part'
      },
      {
        option: '--section [section]',
        description: 'Number of the section to which the web part should be added (1 or higher)'
      },
      {
        option: '--column [column]',
        description: 'Number of the column in which the web part should be added (1 or higher)'
      },
      {
        option: '--order [order]',
        description: 'Order of the web part in the column'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      if (!args.options.pageName) {
        return 'Required option pageName is missing';
      }

      if (!args.options.standardWebPart && !args.options.webPartId) {
        return 'Specify either the standardWebPart or the webPartId option';
      }

      if (args.options.standardWebPart && args.options.webPartId) {
        return 'Specify either the standardWebPart or the webPartId option but not both';
      }

      if (args.options.webPartId && !Utils.isValidGuid(args.options.webPartId)) {
        return `The webPartId '${args.options.webPartId}' is not a valid GUID`;
      }

      if (args.options.standardWebPart && !StandardWebPartUtils.isValidStandardWebPartType(args.options.standardWebPart)) {
        return `${args.options.standardWebPart} is not a valid standard web part type`;
      }

      if (args.options.webPartProperties) {
        try {
          JSON.parse(args.options.webPartProperties);
        }
        catch (e) {
          return `Specified webPartProperties is not a valid JSON string. Input: ${args.options
            .webPartProperties}. Error: ${e}`;
        }
      }

      if (args.options.section && (!isNumber(args.options.section) || args.options.section < 1)) {
        return 'The value of parameter section must be 1 or higher';
      }

      if (args.options.column && (!isNumber(args.options.column) || args.options.column < 1)) {
        return 'The value of parameter column must be 1 or higher';
      }

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site
    using the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To add a client-side web part to a modern page, you have to first log in
    to a SharePoint site using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    If the specified ${chalk.grey('pageName')} doesn't refer to an existing modern page,
    you will get a ${chalk.grey("File doesn't exists")} error.

    To add a standard web part to the page, specify one of the following values:
    ${chalk.grey("ContentRollup, BingMap, ContentEmbed, DocumentEmbed, Image,")}
    ${chalk.grey("ImageGallery, LinkPreview, NewsFeed, NewsReel, PowerBIReportEmbed,")}
    ${chalk.grey("QuickChart, SiteActivity, VideoEmbed, YammerEmbed, Events,")}
    ${chalk.grey("GroupCalendar, Hero, List, PageTitle, People, QuickLinks,")}
    ${chalk.grey("CustomMessageRegion, Divider, MicrosoftForms, Spacer")}.

    When specifying the JSON string with web part properties on Windows, you
    have to escape double quotes in a specific way. Considering the following
    value for the _webPartProperties_ option: {"Foo":"Bar"},
    you should specify the value as \`"{""Foo"":""Bar""}"\`. In addition,
    when using PowerShell, you should use the --% argument.

  Examples:

    Add the standard Bing Map web part to a modern page in the first available location on the page
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap

    Add the standard Bing Map web part to a modern page in the third column of the second section
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap --section 2 --column 3

    Add a custom web part to the modern page
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --webPartId 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1

    Using PowerShell, add the standard Bing Map web part with the specific properties to a modern page
      ${chalk.grey(config.delimiter)} --% ${this.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap --webPartProperties \`"{""Title"":""Foo location""}"\`

    Using Windows command line, add the standard Bing Map web part with the specific properties to a modern page
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --standardWebPart BingMap --webPartProperties \`"{""Title"":""Foo location""}"\`
      `
    );
  }
}

module.exports = new SpoPageClientSideWebPartAddCommand();
