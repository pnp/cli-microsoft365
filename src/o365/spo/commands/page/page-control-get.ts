import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import {
  CommandOption, CommandValidate, CommandError
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';
import { PageItem } from './PageItem';
import { ClientSidePage, ClientSidePart } from './clientsidepages';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  name: string;
  webUrl: string;
}

class SpoPageControlGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_CONTROL_GET}`;
  }

  public get description(): string {
    return 'Gets information about the specific control on a modern page';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving information about the page...`);
        }

        let pageName: string = args.options.name;
        if (args.options.name.indexOf('.aspx') < 0) {
          pageName += '.aspx';
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${args.options.webUrl.substr(args.options.webUrl.indexOf('/', 8))}/SitePages/${encodeURIComponent(pageName)}')?$expand=ListItemAllFields/ClientSideApplicationId`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: PageItem): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (res.ListItemAllFields.ClientSideApplicationId !== 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
          cmd.log(new CommandError(`Page ${args.options.name} is not a modern page.`));
          cb();
          return;
        }

        const clientSidePage: ClientSidePage = ClientSidePage.fromHtml(res.ListItemAllFields.CanvasContent1);
        const control: ClientSidePart | null = clientSidePage.findControlById(args.options.id);

        if (control) {
          // remove the column property to be able to serialize the object to JSON
          delete control.column;

          if (args.options.output !== 'json') {
            (control as any).controlType = SpoPageControlGetCommand.getControlTypeDisplayName((control as any).controlType);
          }

          cmd.log(control);

          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }
        }
        else {
          if (this.verbose) {
            cmd.log(`Control with ID ${args.options.id} not found on page ${args.options.name}`);
          }
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private static getControlTypeDisplayName(controlType: number): string {
    switch (controlType) {
      case 0:
        return 'Empty column';
      case 3:
        return 'Client-side web part';
      case 4:
        return 'Client-side text';
      default:
        return '' + controlType;
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'ID of the control to retrieve information for'
      },
      {
        option: '-n, --name <name>',
        description: 'Name of the page where the control is located'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to retrieve is located'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (!args.options.name) {
        return 'Required parameter name missing';
      }

      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
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

    To get information about a control on a modern page, you have to first
    connect to a SharePoint site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    If the specified ${chalk.grey('name')} doesn't refer to an existing modern page, you will get
    a ${chalk.grey('File doesn\'t exists')} error.

  Examples:
  
    Get information about the control with ID
    ${chalk.grey('3ede60d3-dc2c-438b-b5bf-cc40bb2351e1')} placed on a modern page
    with name ${chalk.grey('home.aspx')}
      ${chalk.grey(config.delimiter)} ${this.name} --id 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1 --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx
`);
  }
}

module.exports = new SpoPageControlGetCommand();