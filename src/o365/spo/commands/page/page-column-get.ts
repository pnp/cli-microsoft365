import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';
import { ClientSidePage, CanvasSection } from './clientsidepages';
import { Page } from './Page';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  section: number;
  column: number;
  webUrl: string;
}

class SpoPageColumnGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_COLUMN_GET}`;
  }

  public get description(): string {
    return 'Get information about a specific column of a modern page';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): Promise<ClientSidePage> => {
        return Page.getPage(args.options.name, args.options.webUrl, accessToken, cmd, this.debug, this.verbose);
      })
      .then((clientSidePage: ClientSidePage): void => {
        const sections: CanvasSection[] = clientSidePage.sections
          .filter(section => section.order === args.options.section);

        if (sections.length) {
          const isJSONOutput = args.options.output === 'json';
          const columns = sections[0].columns.filter(col => col.order === args.options.column);
          if (columns.length) {
            const column = Page.getColumnsInformation(columns[0], isJSONOutput);
            column.controls = columns[0].controls.map(control => Page.getControlsInformation(control, isJSONOutput));
            cmd.log(column);
          }
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to retrieve is located'
      },
      {
        option: '-n, --name <name>',
        description: 'Name of the page to list columns of'
      },
      {
        option: '-s, --section <sectionId>',
        description: 'ID of the section for which to list columns'
      },
      {
        option: '-c, --column <column>',
        description: 'ID of the column for which to retrieve more information'
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

      if (!args.options.section) {
        return 'Required parameter section missing';
      }
      else {
        if (isNaN(args.options.section)) {
          return `${args.options.section} is not a number`;
        }
      }

      if (!args.options.column) {
        return 'Required parameter column missing';
      }
      else {
        if (isNaN(args.options.column)) {
          return `${args.options.column} is not a number`;
        }
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

    To get information about the specific column of a modern page, you have to first
    connect to a SharePoint site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    If the specified ${chalk.grey('name')} doesn't refer to an existing modern
    page, you will get a ${chalk.grey('File doesn\'t exists')} error.

  Examples:
  
    Get information about the specific column of a modern page with name ${chalk.grey('home.aspx')}
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx --section 1 --column 1
`);
  }
}

module.exports = new SpoPageColumnGetCommand();