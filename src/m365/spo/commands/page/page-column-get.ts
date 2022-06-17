import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils';
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
    return commands.PAGE_COLUMN_GET;
  }

  public get description(): string {
    return 'Get information about a specific column of a modern page';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '-s, --section <section>'
      },
      {
        option: '-c, --column <column>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (isNaN(args.options.section)) {
          return `${args.options.section} is not a number`;
        }

        if (isNaN(args.options.column)) {
          return `${args.options.column} is not a number`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
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

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoPageColumnGetCommand();