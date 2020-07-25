import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { ClientSidePage, CanvasSection } from './clientsidepages';
import { Page } from './Page';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  section: number;
  webUrl: string;
}

class SpoPageSectionGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_SECTION_GET}`;
  }

  public get description(): string {
    return 'Get information about the specified modern page section';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    Page
      .getPage(args.options.name, args.options.webUrl, cmd, this.debug, this.verbose)
      .then((clientSidePage: ClientSidePage): void => {
        const sections: CanvasSection[] = clientSidePage.sections
          .filter(section => section.order === args.options.section);

        const isJSONOutput = args.options.output === 'json';
        if (sections.length) {
          cmd.log(Page.getSectionInformation(sections[0], isJSONOutput));
        }

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
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
        description: 'Name of the page to get section information of'
      },
      {
        option: '-s, --section <sectionId>',
        description: 'ID of the section for which to retrieve information'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (isNaN(args.options.section)) {
        return `${args.options.section} is not a number`;
      }

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoPageSectionGetCommand();