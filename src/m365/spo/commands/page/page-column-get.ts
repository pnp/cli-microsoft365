import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { CanvasSection } from './clientsidepages';
import { Page } from './Page';

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
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoPageColumnGetCommand();