import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';
import { ClientSidePage, CanvasSectionTemplate } from './clientsidepages';
import { ContextInfo } from '../../spo';
import { Page } from './Page';
import { isNumber } from 'util';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  pageName: string;
  webUrl: string;
  sectionTemplate: CanvasSectionTemplate;
  order?: number;
}

class SpoPageSectionAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_SECTION_ADD}`;
  }

  public get description(): string {
    return 'Add section to modern page';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';
    let requestDigest: string = '';

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
        cmd.log(`Retrieving modern page ${args.options.pageName}...`);
      }
      // Get Client Side Page
      return Page.getPage(args.options.pageName, args.options.webUrl, siteAccessToken, cmd, this.debug, this.verbose);
    })
    .then((clientSidePage: ClientSidePage): request.RequestPromise => {

        clientSidePage.addSection(args.options.sectionTemplate, args.options.order);

        // Save the Client Side Page with updated section
        return this.saveClientSidePage(clientSidePage as ClientSidePage, cmd, args, args.options.pageName, siteAccessToken, requestDigest);
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

    }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, cmd, cb)
    });
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

    const updatedContent: string = clientSidePage.  toHtml();

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
        option: '-n, --pageName <pageNname>',
        description: 'Name of the page to to add section'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to retrieve is located'
      },
      {
        option: '-t, --sectionTemplate <sectionTemplate>',
        description: 'type of section to add. Allowed values OneColumn|OneColumnFullWidth| TwoColumn|ThreeColumn|TwoColumnLeft|TwoColumnRight'
      },
      {
        option: '-o, --order <order>',
        description: 'order of the section'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.pageName) {
        return 'Required parameter pageName missing';
      }

      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      if (!args.options.sectionTemplate) {
        return 'Required parameter sectionTemplate missing';
      }
      else {
        if (!(args.options.sectionTemplate in CanvasSectionTemplate)) {
          return `${args.options.sectionTemplate} is not a valid section template`;
        }
      }

      if (!isNumber(args.options.order) || (args.options.order != undefined && args.options.order < 1)) {
        return 'The value of parameter order must be 1 or higher';
      }

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site
    using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To add a section to the modern page, you have to first connect to
    a SharePoint site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    If the specified ${chalk.grey('pageName')} doesn't refer to an existing modern 
    page, you will get a ${chalk.grey('File doesn\'t exists')} error.

  Examples:
  
    Get information about adding section to the modern page
    named ${chalk.grey('home.aspx')}
      ${chalk.grey(config.delimiter)} ${this.name} --pageName home.aspx --webUrl https://contoso.sharepoint.com/sites/team-a  --sectionTemplate OneColumn --order 1
`);
  }
}

module.exports = new SpoPageSectionAddCommand();