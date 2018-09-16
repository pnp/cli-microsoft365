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
  name: string;
  webUrl: string;
  sectionTemplate: CanvasSectionTemplate;
  order?: number;
}

class SpoPageSectionAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_SECTION_ADD}`;
  }

  public get description(): string {
    return 'Adds section to modern page';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';
    let requestDigest: string = '';
    let pageFullName: string = args.options.name.toLowerCase();
    if (pageFullName.indexOf('.aspx') < 0) {
      pageFullName += '.aspx';
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
          cmd.log(`Retrieving modern page ${args.options.name}...`);
        }
        // Get Client Side Page
        return Page.getPage(pageFullName, args.options.webUrl, siteAccessToken, cmd, this.debug, this.verbose);
      })
      .then((clientSidePage: ClientSidePage): request.RequestPromise => {
        clientSidePage.addSection(args.options.sectionTemplate, args.options.order);

        // Save the Client Side Page with updated section
        return this.saveClientSidePage(clientSidePage as ClientSidePage, cmd, args, pageFullName, siteAccessToken, requestDigest);
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
    name: string,
    accessToken: string,
    requestDigest: string
  ): request.RequestPromise<any> {
    const serverRelativeSiteUrl: string = `${args.options.webUrl.substr(
      args.options.webUrl.indexOf('/', 8)
    )}/sitepages/${name}`;

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
        option: '-n, --name <name>',
        description: 'Name of the page to add section to'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to retrieve is located'
      },
      {
        option: '-t, --sectionTemplate <sectionTemplate>',
        description: 'Type of section to add. Allowed values OneColumn|OneColumnFullWidth|TwoColumn|ThreeColumn|TwoColumnLeft|TwoColumnRight'
      },
      {
        option: '--order [order]',
        description: 'Order of the section to add'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name) {
        return 'Required parameter name missing';
      }

      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      if (!args.options.sectionTemplate) {
        return 'Required parameter sectionTemplate missing';
      }
      else {
        if (!(args.options.sectionTemplate in CanvasSectionTemplate)) {
          return `${args.options.sectionTemplate} is not a valid section template. Allowed values are OneColumn|OneColumnFullWidth|TwoColumn|ThreeColumn|TwoColumnLeft|TwoColumnRight`;
        }
      }

      if (args.options.order) {
        if (!isNumber(args.options.order) || args.options.order < 1) {
          return 'The value of parameter order must be 1 or higher';
        }
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

    To add a section to the modern page, you have to first log in to
    a SharePoint site using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    If the specified ${chalk.grey('name')} doesn't refer to an existing modern 
    page, you will get a ${chalk.grey('File doesn\'t exists')} error.

  Examples:
  
    Add section to the modern page named ${chalk.grey('home.aspx')}
      ${chalk.grey(config.delimiter)} ${this.name} --name home.aspx --webUrl https://contoso.sharepoint.com/sites/newsletter  --sectionTemplate OneColumn --order 1
`);
  }
}

module.exports = new SpoPageSectionAddCommand();