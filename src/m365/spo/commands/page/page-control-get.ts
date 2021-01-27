import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSidePage, ClientSidePart } from './clientsidepages';
import { Page } from './Page';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    Page
      .getPage(args.options.name, args.options.webUrl, logger, this.debug, this.verbose)
      .then((clientSidePage: ClientSidePage): void => {
        const control: ClientSidePart | null = clientSidePage.findControlById(args.options.id);

        if (control) {
          const isJSONOutput = args.options.output === 'json';

          logger.log(JSON.parse(JSON.stringify(Page.getControlsInformation(control, isJSONOutput))));

          if (this.verbose) {
            logger.logToStderr(chalk.green('DONE'));
          }
        }
        else {
          if (this.verbose) {
            logger.logToStderr(`Control with ID ${args.options.id} not found on page ${args.options.name}`);
          }
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoPageControlGetCommand();