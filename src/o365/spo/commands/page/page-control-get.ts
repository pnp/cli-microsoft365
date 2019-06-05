import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import { ClientSidePage, ClientSidePart } from './clientsidepages';
import { Page } from './Page';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    Page
      .getPage(args.options.name, args.options.webUrl, cmd, this.debug, this.verbose)
      .then((clientSidePage: ClientSidePage): void => {
        const control: ClientSidePart | null = clientSidePage.findControlById(args.options.id);

        if (control) {
          const isJSONOutput = args.options.output === 'json';

          cmd.log(JSON.parse(JSON.stringify(Page.getControlsInformation(control, isJSONOutput))));

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
      `  Remarks:

    If the specified ${chalk.grey('name')} doesn't refer to an existing modern page, you will get
    a ${chalk.grey('File doesn\'t exists')} error.

  Examples:
  
    Get information about the control with ID
    ${chalk.grey('3ede60d3-dc2c-438b-b5bf-cc40bb2351e1')} placed on a modern page
    with name ${chalk.grey('home.aspx')}
      ${this.name} --id 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1 --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx
`);
  }
}

module.exports = new SpoPageControlGetCommand();