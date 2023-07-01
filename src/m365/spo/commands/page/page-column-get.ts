import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { CanvasSection } from './clientsidepages.js';
import { Page } from './Page.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  column: number;
  pageName: string;
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
        option: '-n, --pageName <pageName>'
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const clientSidePage = await Page.getPage(args.options.pageName, args.options.webUrl, logger, this.debug, this.verbose);

      const sections: CanvasSection[] = clientSidePage.sections
        .filter(section => section.order === args.options.section);

      if (sections.length) {
        const isJSONOutput = !Cli.shouldTrimOutput(args.options.output);
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
          await logger.log(column);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageColumnGetCommand();