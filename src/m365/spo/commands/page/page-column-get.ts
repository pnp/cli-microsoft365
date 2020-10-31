import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { CanvasSection, ClientSidePage } from './clientsidepages';
import { Page } from './Page';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  column: number;
  name: string;
  section: number;
  webUrl: string;
}

class SpoPageColumnGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_COLUMN_GET}`;
  }

  public get description(): string {
    return 'Get information about a specific column of a modern page';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    Page
      .getPage(args.options.name, args.options.webUrl, logger, this.debug, this.verbose)
      .then((clientSidePage: ClientSidePage): void => {
        const sections: CanvasSection[] = clientSidePage.sections
          .filter(section => section.order === args.options.section);

        if (sections.length) {
          const isJSONOutput = args.options.output === 'json';
          const columns = sections[0].columns.filter(col => col.order === args.options.column);
          if (columns.length) {
            const column = Page.getColumnsInformation(columns[0], isJSONOutput);
            column.controls = columns[0].controls
              .map(control => Page.getControlsInformation(control, isJSONOutput));
            if (!isJSONOutput) {
              column.controls = (column.controls as any[])
                .map(control => `${control.id} (${control.title})`)
                .join(', ');
            }
            logger.log(column);
          }
        }

        if (this.verbose) {
          logger.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to retrieve is located'
      },
      {
        option: '-n, --name <name>',
        description: 'Name of the page to get column information of'
      },
      {
        option: '-s, --section <section>',
        description: 'ID of the section where the column is located'
      },
      {
        option: '-c, --column <column>',
        description: 'ID of the column for which to retrieve more information'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (isNaN(args.options.section)) {
      return `${args.options.section} is not a number`;
    }

    if (isNaN(args.options.column)) {
      return `${args.options.column} is not a number`;
    }

    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoPageColumnGetCommand();