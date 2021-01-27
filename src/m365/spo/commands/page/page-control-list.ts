import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSidePage, ClientSidePart } from './clientsidepages';
import { Page } from './Page';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
}

class SpoPageControlListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_CONTROL_LIST}`;
  }

  public get description(): string {
    return 'Lists controls on the specific modern page';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'type', 'title'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    Page
      .getPage(args.options.name, args.options.webUrl, logger, this.debug, this.verbose)
      .then((clientSidePage: ClientSidePage): void => {
        let controls: ClientSidePart[] = [];
        clientSidePage.sections.forEach(s => {
          s.columns.forEach(c => {
            controls = controls.concat(c.controls);
          });
        });
        // remove the column property to be able to serialize the array to JSON
        controls.forEach(c => delete c.column);

        // remove the dynamicDataValues and dynamicDataPaths properties if they are null
        controls.forEach(c => {
          if (!c.dynamicDataPaths) {
            delete c.dynamicDataPaths
          }
          if (!c.dynamicDataValues) {
            delete c.dynamicDataValues
          }
        });

        controls.forEach(c => {
          (c as any).type = SpoPageControlListCommand.getControlTypeDisplayName((c as any).controlType);
        });
        // drop the information about original classes from clientsidepages.ts
        logger.log(JSON.parse(JSON.stringify(controls)));

        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoPageControlListCommand();