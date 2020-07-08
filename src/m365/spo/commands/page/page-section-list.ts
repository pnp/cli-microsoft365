import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { ClientSidePage, CanvasSection } from './clientsidepages';
import { Page } from './Page';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
    return `${commands.PAGE_SECTION_LIST}`;
  }

  public get description(): string {
    return 'List sections in the specific modern page';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    Page
      .getPage(args.options.name, args.options.webUrl, cmd, this.debug, this.verbose)
      .then((clientSidePage: ClientSidePage): void => {
        const sections: CanvasSection[] = clientSidePage.sections;

        const isJSONOutput = args.options.output === 'json';
        if (sections.length) {
          let output = sections.map(section => Page.getSectionInformation(section, isJSONOutput));
          if (isJSONOutput) {
            cmd.log(output);
          }
          else {
            cmd.log(output.map(s => {
              return {
                order: s.order,
                columns: s.columns.length
              }
            }));
          }
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
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
        description: 'Name of the page to list sections of'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
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

    If the specified ${chalk.grey('name')} doesn't refer to an existing modern 
    page, you will get a ${chalk.grey('File doesn\'t exists')} error.

  Examples:
  
    List sections of a modern page named ${chalk.grey('home.aspx')}
      ${this.name} --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx
`);
  }
}

module.exports = new SpoPageSectionListCommand();