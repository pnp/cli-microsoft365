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
  ClientSideText
} from './clientsidepages';
import { ContextInfo } from '../../spo';
import { isNumber } from 'util';
import { Page } from './Page';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  column?: number;
  order?: number;
  pageName: string;
  section?: number;
  text: string;
  webUrl: string;
}

class SpoPageTextAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_TEXT_ADD}`;
  }

  public get description(): string {
    return 'Adds text to a modern page';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.section = typeof args.options.section !== 'undefined';
    telemetryProps.column = typeof args.options.column !== 'undefined';
    telemetryProps.order = typeof args.options.order !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';
    let requestDigest: string = '';

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
        return Page.getPage(pageName, args.options.webUrl, siteAccessToken, cmd, this.debug, this.verbose);
      })
      .then((page: ClientSidePage): request.RequestPromise => {
        const section: number = (args.options.section || 1) - 1;
        const column: number = (args.options.column || 1) - 1;

        // Make sure the section is in range
        if (section >= page.sections.length) {
          throw new Error(`Invalid section '${section + 1}'`);
        }

        // Make sure the column is in range
        if (column >= page.sections[section].columns.length) {
          throw new Error(`Invalid column '${column + 1}'`);
        }

        const text: ClientSideText = new ClientSideText(args.options.text);
        if (typeof args.options.order === 'undefined') {
          page.sections[section].columns[column].addControl(text);
        }
        else {
          const order: number = args.options.order - 1;
          page.sections[section].columns[column].insertControl(text, order);
        }

        // Save the Client Side Page with updated information
        return this.saveClientSidePage(page, cmd, args, pageName, siteAccessToken, requestDigest);
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
    )}/SitePages/${pageName}`;

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
        description: 'URL of the site where the page to add the text to is located'
      },
      {
        option: '-n, --pageName <pageName>',
        description: 'Name of the page to which add the text'
      },
      {
        option: '-t, --text <text>',
        description: 'Text to add to the page'
      },
      {
        option: '--section [section]',
        description: 'Number of the section to which the text should be added (1 or higher)'
      },
      {
        option: '--column [column]',
        description: 'Number of the column in which the text should be added (1 or higher)'
      },
      {
        option: '--order [order]',
        description: 'Order of the text in the column'
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

      if (!args.options.text) {
        return 'Required option text is missing';
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

    To add text to a modern page, you have to first log in to a SharePoint site
    using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    If the specified ${chalk.grey('pageName')} doesn't refer to an existing modern page,
    you will get a ${chalk.grey("File doesn't exists")} error.

  Examples:

    Add text to a modern page in the first available location on the page
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --text 'Hello world'

    Add text to a modern page in the third column of the second section
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --text 'Hello world' --section 2 --column 3

    Add text at the beginning of the default column on a modern page
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --text 'Hello world' --order 1
      `
    );
  }
}

module.exports = new SpoPageTextAddCommand();
