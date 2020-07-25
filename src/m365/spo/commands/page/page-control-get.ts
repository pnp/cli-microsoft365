import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import { ClientSidePage, ClientSidePart } from './clientsidepages';
import { Page } from './Page';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
            cmd.log(chalk.green('DONE'));
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
      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoPageControlGetCommand();