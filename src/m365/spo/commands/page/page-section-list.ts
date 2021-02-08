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
  name: string;
  section: number;
  webUrl: string;
}

class SpoPageSectionListCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_SECTION_LIST;
  }

  public get description(): string {
    return 'List sections in the specific modern page';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    Page
      .getPage(args.options.name, args.options.webUrl, logger, this.debug, this.verbose)
      .then((clientSidePage: ClientSidePage): void => {
        const sections: CanvasSection[] = clientSidePage.sections;

        const isJSONOutput = args.options.output === 'json';
        if (sections.length) {
          let output = sections.map(section => Page.getSectionInformation(section, isJSONOutput));
          if (isJSONOutput) {
            logger.log(output);
          }
          else {
            logger.log(output.map(s => {
              return {
                order: s.order,
                columns: s.columns.length
              }
            }));
          }
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-n, --name <name>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoPageSectionListCommand();